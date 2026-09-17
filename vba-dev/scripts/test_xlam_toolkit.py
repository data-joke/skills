# -*- coding: utf-8 -*-
"""
vba-dev 回归测试 —— 默认纯静态，不启动 Excel。

用途：改动同目录的 xlam_toolkit.py 后自验，防止把既有行为碰坏。

覆盖：
  - _join_continuations()              续行合并（修复过：曾每次多吞一个字符）
  - _check_early_binding_in_module()   前期绑定检测（分发安全规范）
  - _check_blocks_in_module()          块结构配对
  - _check_structure_in_module()       过程外代码 / 未闭合 Type·Enum
  - get_file_format()                  格式判定
  - build_check / _static_check_project  签名（防止 check_binding 参数被误删）
  - render_ribbon_preview()            功能区预览渲染 + 临时目录清理
  - get/set/clear_preferences()        跨会话偏好读写

用法：
    python scripts/test_xlam_toolkit.py
    python -m unittest discover -s scripts -v

Excel 端到端用例默认跳过；需要时设环境变量 VBA_DEV_EXCEL_TESTS=1 启用
（会启动隔离 Excel 实例，**不会**触碰你自己打开的 Excel）。

设计约束：静态用例全部为纯文本，不触碰 COM。导入 toolkit 不会建 Excel 实例，
atexit 注册的 shutdown_excel() 在无缓存实例时直接 return，不会反向拉起 Excel。
"""
import inspect
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from xlam_toolkit import (  # noqa: E402
    build_check,
    get_file_format,
    _check_blocks_in_module,
    _check_early_binding_in_module,
    _check_structure_in_module,
    _join_continuations,
)


def _binding_errors(lines):
    errs = []
    _check_early_binding_in_module("Mod1", lines, errs)
    return errs


def _block_errors(lines):
    errs = []
    _check_blocks_in_module("Mod1", lines, errs)
    return errs


def _struct_errors(lines):
    errs = []
    _check_structure_in_module("Mod1", lines, errs)
    return errs


# ---------------------------------------------------------------------------
# 续行合并
# ---------------------------------------------------------------------------
class TestJoinContinuations(unittest.TestCase):
    """_join_continuations 曾经每次拼接多吞掉前一段的末字符
    （And→An、As→A、+ 与 = 直接消失），这里锁死正确行为。"""

    def test_single_continuation(self):
        self.assertEqual(
            _join_continuations(["Dim d As _", "    Scripting.Dictionary"]),
            ["Dim d As Scripting.Dictionary"])

    def test_preserves_trailing_keyword(self):
        # 回归点：旧实现把此处的 "And" 啃成 "An"
        self.assertEqual(
            _join_continuations(["If a = 1 And _", "    b = 2 Then"]),
            ["If a = 1 And b = 2 Then"])

    def test_keyword_at_fragment_end(self):
        # 回归点：片段以关键字结尾时会被啃掉末字符（"Then" -> "The"），
        # 这类损坏会直接改变结构判定，最危险
        self.assertEqual(
            _join_continuations(["If a Then _", "x = 1"]),
            ["If a Then x = 1"])

    def test_preserves_operators(self):
        # 回归点：旧实现把 "+" 吃掉
        self.assertEqual(
            _join_continuations(["x = 1 + _", "    2 + _", "    3"]),
            ["x = 1 + 2 + 3"])

    def test_multi_line_condition(self):
        self.assertEqual(
            _join_continuations(["If a = 1 And _", "   b = 2 _", "   Then"]),
            ["If a = 1 And b = 2 Then"])

    def test_non_continuation_untouched(self):
        self.assertEqual(
            _join_continuations(["a = 1", "b = 2"]),
            ["a = 1", "b = 2"])

    def test_other_lines_survive(self):
        self.assertEqual(
            _join_continuations(
                ["Sub S()", "    Dim d As _", "        Scripting.Dictionary", "End Sub"]),
            ["Sub S()", "    Dim d As Scripting.Dictionary", "End Sub"])


# ---------------------------------------------------------------------------
# 前期绑定检测
# ---------------------------------------------------------------------------
class TestEarlyBindingDetected(unittest.TestCase):

    def test_new_keyword(self):
        errs = _binding_errors(["Dim d As New Scripting.Dictionary"])
        self.assertEqual(len(errs), 1)
        self.assertIn('CreateObject("Scripting.Dictionary")', errs[0]["error"])

    def test_as_declaration(self):
        errs = _binding_errors(["Dim d As Scripting.Dictionary"])
        self.assertEqual(len(errs), 1)
        self.assertIn('CreateObject("Scripting.Dictionary")', errs[0]["error"])

    def test_procedure_signature(self):
        errs = _binding_errors(["Sub F(d As Outlook.Application)", "End Sub"])
        self.assertEqual(len(errs), 1)
        self.assertIn("Outlook.Application", errs[0]["error"])

    def test_progid_alias_regexp(self):
        # 类型库是 VBScript_RegExp_55，但 ProgID 是 VBScript.RegExp
        errs = _binding_errors(["Dim re As New VBScript_RegExp_55.RegExp"])
        self.assertEqual(len(errs), 1)
        self.assertIn('CreateObject("VBScript.RegExp")', errs[0]["error"])

    def test_progid_alias_msxml(self):
        errs = _binding_errors(["Dim x As New MSXML2.XMLHTTP60"])
        self.assertEqual(len(errs), 1)
        self.assertIn('CreateObject("MSXML2.XMLHTTP")', errs[0]["error"])

    def test_progid_same_as_libtype(self):
        for decl, progid in [
            ("Dim f As New Scripting.FileSystemObject", "Scripting.FileSystemObject"),
            ("Dim s As New WScript.Shell", "WScript.Shell"),
            ("Dim a As New Shell.Application", "Shell.Application"),
            ("Dim cn As New ADODB.Connection", "ADODB.Connection"),
            ("Dim o As New Outlook.Application", "Outlook.Application"),
        ]:
            with self.subTest(decl=decl):
                errs = _binding_errors([decl])
                self.assertEqual(len(errs), 1)
                self.assertIn(f'CreateObject("{progid}")', errs[0]["error"])

    def test_withevents_hint_differs(self):
        # WithEvents 无法后期绑定，文案不应给出做不到的 CreateObject 建议
        errs = _binding_errors(["Dim WithEvents o As Outlook.Application"])
        self.assertEqual(len(errs), 1)
        self.assertIn("WithEvents", errs[0]["error"])
        self.assertNotIn("CreateObject", errs[0]["error"])

    def test_continued_declaration_detected(self):
        errs = _binding_errors(["Dim d As _", "    Scripting.Dictionary"])
        self.assertEqual(len(errs), 1)
        self.assertIn('CreateObject("Scripting.Dictionary")', errs[0]["error"])

    def test_one_error_per_type_per_module(self):
        errs = _binding_errors([
            "Dim a As Scripting.Dictionary",
            "Dim b As New Scripting.Dictionary",
        ])
        self.assertEqual(len(errs), 1)

    def test_distinct_types_each_reported(self):
        errs = _binding_errors([
            "Dim a As New Scripting.Dictionary",
            "Dim b As New ADODB.Connection",
        ])
        self.assertEqual(len(errs), 2)

    def test_carries_kind_field(self):
        errs = _binding_errors(["Dim d As New Scripting.Dictionary"])
        self.assertEqual(errs[0]["kind"], "early_binding")
        self.assertIn("line", errs[0])
        self.assertIn("module", errs[0])


class TestEarlyBindingAllowed(unittest.TestCase):

    def test_late_binding_ok(self):
        self.assertEqual(
            _binding_errors(['Set d = CreateObject("Scripting.Dictionary")']), [])

    def test_host_libraries_ok(self):
        for decl in [
            "Sub F(c As IRibbonControl)",
            "Dim ui As IRibbonUI",
            "Dim r As Excel.Range",
            "Dim t As MSForms.TextBox",
            "Dim o As Office.IRibbonUI",
            "Dim a As stdole.IPictureDisp",
        ]:
            with self.subTest(decl=decl):
                self.assertEqual(_binding_errors([decl, "End Sub"]), [])

    def test_plain_types_ok(self):
        self.assertEqual(
            _binding_errors(["Dim s As String", "Dim x As Object",
                             "Dim c As Collection", "Dim v As Variant"]), [])

    def test_comment_ignored(self):
        self.assertEqual(
            _binding_errors(["' Dim d As Scripting.Dictionary"]), [])

    def test_custom_class_ok(self):
        self.assertEqual(_binding_errors(["Dim o As MyHelperClass"]), [])

    def test_exemption_marker(self):
        self.assertEqual(
            _binding_errors(["' @allow-early-binding 理由：WithEvents sink",
                             "Dim d As New Scripting.Dictionary"]), [])


# ---------------------------------------------------------------------------
# 块结构检查（既有行为）
# ---------------------------------------------------------------------------
class TestBlockChecks(unittest.TestCase):

    def test_valid_block_if_with_continuation(self):
        # 续行写在 Then 之前的块 If，应被识别为块结构（而不是单行 If）。
        # 注：此用例对旧 _join_continuations 的损坏不敏感（旧算法拼成
        # "...An b = 2 Then"，Then 仍在行尾），它守的是块检查器本身。
        self.assertEqual(_block_errors([
            "Sub S()",
            "    If a = 1 And _",
            "       b = 2 Then",
            "        x = 1",
            "    End If",
            "End Sub"]), [])

    def test_missing_end_if(self):
        errs = _block_errors(["Sub S()", "    If a Then", "        x = 1", "End Sub"])
        self.assertEqual(len(errs), 1)

    def test_single_line_if_is_not_block_opener(self):
        self.assertEqual(
            _block_errors(["Sub S()", "    If a Then x = 1", "End Sub"]), [])

    def test_balanced_constructs(self):
        self.assertEqual(_block_errors([
            "Sub S()",
            "    For i = 1 To 3", "    Next",
            "    For Each c In r", "    Next",
            "    Do While x", "    Loop",
            "    While y", "    Wend",
            "    With rng", "    End With",
            "    Select Case v", "    End Select",
            "End Sub"]), [])

    def test_one_line_for_ok(self):
        self.assertEqual(
            _block_errors(["Sub S()", "    For i = 1 To 3: x = i: Next", "End Sub"]), [])

    def test_end_if_without_if(self):
        errs = _block_errors(["Sub S()", "    End If", "End Sub"])
        self.assertEqual(len(errs), 1)

    def test_conditional_compilation(self):
        self.assertEqual(_block_errors([
            "Sub S()",
            "    #If VBA7 Then", "        x = 1", "    #End If",
            "End Sub"]), [])

    def test_unclosed_procedure(self):
        errs = _block_errors(["Sub S()", "    x = 1"])
        self.assertEqual(len(errs), 1)


# ---------------------------------------------------------------------------
# 结构检查（既有行为）
# ---------------------------------------------------------------------------
class TestStructureChecks(unittest.TestCase):

    def test_module_level_declarations_ok(self):
        self.assertEqual(_struct_errors([
            "Option Explicit",
            "Private Const MAX As Long = 10",
            "Dim g As Integer",
            "Sub S()", "End Sub"]), [])

    def test_executable_code_outside_proc(self):
        errs = _struct_errors(["Sub S()", "End Sub", "x = 1"])
        self.assertEqual(len(errs), 1)
        self.assertIn("过程外", errs[0]["error"])

    def test_type_enum_blocks_ok(self):
        self.assertEqual(_struct_errors([
            "Type T", "    a As Long", "End Type",
            "Enum E", "    A = 1", "End Enum"]), [])

    def test_unclosed_type_block(self):
        errs = _struct_errors(["Type T", "    a As Long"])
        self.assertEqual(len(errs), 1)
        self.assertIn("Type/Enum", errs[0]["error"])


# ---------------------------------------------------------------------------
# 文件格式
# ---------------------------------------------------------------------------
class TestFileFormat(unittest.TestCase):

    def test_formats(self):
        for path, want in [("a.xlam", "xlam"), ("a.xlsm", "xlsm"),
                           ("a.xlsx", "xlsx"), ("a.txt", "unknown"),
                           ("A.XLAM", "xlam")]:
            with self.subTest(path=path):
                self.assertEqual(get_file_format(path), want)


# ---------------------------------------------------------------------------
# API 签名（防止参数被误删）
# ---------------------------------------------------------------------------
class TestApiSurface(unittest.TestCase):

    def test_build_check_has_check_binding(self):
        params = inspect.signature(build_check).parameters
        self.assertIn("check_binding", params)
        self.assertTrue(params["check_binding"].default)

    def test_static_check_project_has_check_binding(self):
        from xlam_toolkit import _static_check_project
        params = inspect.signature(_static_check_project).parameters
        self.assertIn("check_binding", params)


# ---------------------------------------------------------------------------
# Ribbon 预览渲染（纯绘图，不启 Excel）
# ---------------------------------------------------------------------------
class TestRibbonPreview(unittest.TestCase):

    CUSTOMUI = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui"><ribbon><tabs>\n'
        '<tab id="tabAI" label="数你皮">\n'
        '  <group id="grpQuick" label="快速工具">\n'
        '    <button id="b1" label="清理文本" size="large" imageMso="Copy" onAction="A_Click"/>\n'
        '    <button id="b2" label="小按钮" onAction="B_Click"/>\n'
        '  </group>\n'
        '  <group id="grpEmpty" label="空分组"/>\n'
        '</tab>\n'
        '<tab id="tabAbout" label="关于"/>\n'
        '</tabs></ribbon></customUI>')

    def _pkg(self, d):
        os.makedirs(os.path.join(d, "customUI", "_rels"))
        with open(os.path.join(d, "customUI", "customUI.xml"), "w", encoding="utf-8") as f:
            f.write(self.CUSTOMUI)
        return d

    def test_renders_directory_source(self):
        import tempfile
        from PIL import Image
        from xlam_toolkit import render_ribbon_preview
        with tempfile.TemporaryDirectory() as d:
            self._pkg(d)
            out = render_ribbon_preview(d, os.path.join(d, "p.png"))
            self.assertTrue(os.path.exists(out))
            with Image.open(out) as im:
                self.assertGreater(im.width, 0)
                self.assertGreater(im.height, 0)

    def test_tab_index_selects_other_tab(self):
        import tempfile
        from xlam_toolkit import render_ribbon_preview
        with tempfile.TemporaryDirectory() as d:
            self._pkg(d)
            out = render_ribbon_preview(d, os.path.join(d, "p1.png"), tab_index=1)
            self.assertTrue(os.path.exists(out))

    def test_renders_xlam_source_and_cleans_temp(self):
        import tempfile
        import zipfile
        import glob
        from xlam_toolkit import render_ribbon_preview
        with tempfile.TemporaryDirectory() as d:
            self._pkg(d)
            xlam = os.path.join(d, "fake.xlam")
            with zipfile.ZipFile(xlam, "w") as z:
                z.write(os.path.join(d, "customUI", "customUI.xml"),
                        "customUI/customUI.xml")
            before = set(glob.glob(os.path.join(tempfile.gettempdir(), "ribbon_preview_*")))
            out = render_ribbon_preview(xlam, os.path.join(d, "p.png"))
            after = set(glob.glob(os.path.join(tempfile.gettempdir(), "ribbon_preview_*")))
            self.assertTrue(os.path.exists(out))
            self.assertEqual(before, after, "临时解包目录未清理")

    def test_missing_path_raises(self):
        from xlam_toolkit import render_ribbon_preview
        with self.assertRaises(FileNotFoundError):
            render_ribbon_preview("no_such_file_xyz")

    def test_no_ribbon_raises(self):
        import tempfile
        from xlam_toolkit import render_ribbon_preview
        with tempfile.TemporaryDirectory() as d:
            os.makedirs(os.path.join(d, "customUI"))
            with open(os.path.join(d, "customUI", "customUI.xml"), "w",
                      encoding="utf-8") as f:
                f.write("<customUI/>")
            with self.assertRaises(ValueError):
                render_ribbon_preview(d)


# ---------------------------------------------------------------------------
# 用户偏好（用临时文件，绝不碰用户真实配置）
# ---------------------------------------------------------------------------
class TestPreferences(unittest.TestCase):

    def setUp(self):
        import tempfile
        import xlam_toolkit as xt
        self.xt = xt
        # 沙箱化迁移源与目标：_migrate_legacy_config 在 get_preferences 内
        # 运行，不隔离就会动到用户真实配置（~/.vba-dev、~/.claude/vba-dev、scripts/）
        self._d = tempfile.TemporaryDirectory()
        sandbox = self._d.name
        self._orig = xt.PREF_FILE
        self._orig_config_dir = xt.CONFIG_DIR
        self._orig_legacy_dir = xt._LEGACY_CONFIG_DIR
        self._orig_legacy_file = xt._LEGACY_ICON_CONFIG_FILE
        xt.CONFIG_DIR = sandbox
        xt._LEGACY_CONFIG_DIR = sandbox
        xt._LEGACY_ICON_CONFIG_FILE = os.path.join(sandbox, "icon_config.json")
        xt.PREF_FILE = os.path.join(sandbox, "preferences.json")

    def tearDown(self):
        self.xt.PREF_FILE = self._orig
        self.xt.CONFIG_DIR = self._orig_config_dir
        self.xt._LEGACY_CONFIG_DIR = self._orig_legacy_dir
        self.xt._LEGACY_ICON_CONFIG_FILE = self._orig_legacy_file
        self._d.cleanup()

    def test_empty_by_default(self):
        self.assertEqual(self.xt.get_preferences(), {})

    def test_set_get_roundtrip(self):
        self.xt.set_preference("icon_type", "imageMso")
        self.xt.set_preference("theme_color", "#1F4E92")
        self.assertEqual(self.xt.get_preferences(),
                         {"icon_type": "imageMso", "theme_color": "#1F4E92"})

    def test_merge_not_replace(self):
        self.xt.set_preference("a", 1)
        self.xt.set_preference("b", 2)
        self.assertEqual(self.xt.get_preferences(), {"a": 1, "b": 2})

    def test_none_deletes_key(self):
        self.xt.set_preference("a", 1)
        self.xt.set_preference("a", None)
        self.assertEqual(self.xt.get_preferences(), {})

    def test_clear(self):
        self.xt.set_preference("a", 1)
        self.xt.clear_preferences()
        self.assertEqual(self.xt.get_preferences(), {})

    def test_corrupt_file_returns_empty(self):
        with open(self.xt.PREF_FILE, "w", encoding="utf-8") as f:
            f.write("{ not json")
        self.assertEqual(self.xt.get_preferences(), {})

    def test_does_not_touch_real_pref_file(self):
        # 守卫：本测试类必须写到临时路径，不得污染用户真实偏好
        # （当前真实位置 ~/.vba-dev/，历史位置 ~/.claude/vba-dev/ 一并看住）
        home = os.path.expanduser("~")
        for real in (os.path.join(home, ".vba-dev", "preferences.json"),
                     os.path.join(home, ".claude", "vba-dev", "preferences.json")):
            self.assertNotEqual(self.xt.PREF_FILE, real)


# ---------------------------------------------------------------------------
class TestConfigMigration(unittest.TestCase):
    """~/.claude/vba-dev/ 与 scripts/ 内的旧配置 → ~/.vba-dev/ 自动迁移"""

    def setUp(self):
        import tempfile
        import xlam_toolkit as xt
        self.xt = xt
        self._d = tempfile.TemporaryDirectory()
        root = self._d.name
        self._new_dir = os.path.join(root, "new")
        self._legacy_dir = os.path.join(root, "claude", "vba-dev")
        self._skill_dir = os.path.join(root, "skill", "scripts")
        self._orig = (xt.CONFIG_DIR, xt._LEGACY_CONFIG_DIR, xt._LEGACY_ICON_CONFIG_FILE)
        xt.CONFIG_DIR = self._new_dir
        xt._LEGACY_CONFIG_DIR = self._legacy_dir
        xt._LEGACY_ICON_CONFIG_FILE = os.path.join(self._skill_dir, "icon_config.json")

    def tearDown(self):
        self.xt.CONFIG_DIR, self.xt._LEGACY_CONFIG_DIR, self.xt._LEGACY_ICON_CONFIG_FILE = self._orig
        self._d.cleanup()

    def _write(self, path, content='"unit-test"'):
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            f.write(content)

    def test_migrates_agent_specific_dir(self):
        # 旧 agent 专属目录（~/.claude/vba-dev）里的两个文件都迁过去
        self._write(os.path.join(self._legacy_dir, "preferences.json"), '{"k":1}')
        self._write(os.path.join(self._legacy_dir, "icon_config.json"), '{"m":2}')
        self.xt._migrate_legacy_config()
        self.assertTrue(os.path.exists(os.path.join(self._new_dir, "preferences.json")))
        self.assertTrue(os.path.exists(os.path.join(self._new_dir, "icon_config.json")))
        self.assertFalse(os.path.exists(os.path.join(self._legacy_dir, "preferences.json")))

    def test_migrates_in_skill_legacy(self):
        # 更早的 skill 目录内配置也迁走
        self._write(self.xt._LEGACY_ICON_CONFIG_FILE, '{"in_skill":true}')
        self.xt._migrate_legacy_config()
        self.assertTrue(os.path.exists(os.path.join(self._new_dir, "icon_config.json")))
        self.assertFalse(os.path.exists(self.xt._LEGACY_ICON_CONFIG_FILE))

    def test_never_overwrites_existing(self):
        # 新位置已有文件时，旧文件原样保留（不覆盖）
        self._write(os.path.join(self._new_dir, "preferences.json"), '{"keep":1}')
        self._write(os.path.join(self._legacy_dir, "preferences.json"), '{"old":2}')
        self.xt._migrate_legacy_config()
        with open(os.path.join(self._new_dir, "preferences.json"), encoding="utf-8") as f:
            self.assertIn("keep", f.read())
        self.assertTrue(os.path.exists(os.path.join(self._legacy_dir, "preferences.json")))

    def test_noop_when_nothing_legacy(self):
        # 没有旧文件时不创建任何东西、不抛错
        self.xt._migrate_legacy_config()
        self.assertFalse(os.path.exists(self._new_dir))


# ---------------------------------------------------------------------------
class TestProcessAlive(unittest.TestCase):
    """_process_alive 必须 bytes 判断：PYTHONUTF8=1 进程里 text=True 会按 utf-8
    解码 tasklist 的 GBK 本地化输出，读线程崩溃且恒返回 False，导致 shutdown_excel
    的残留强杀保险失效（2026-09-12 修复，曾真实产生僵尸 EXCEL.EXE）"""

    def test_dead_pid_false_no_crash(self):
        import xlam_toolkit as xt
        self.assertFalse(xt._process_alive(999999))

    def test_returns_bool_not_none(self):
        import xlam_toolkit as xt
        self.assertIsInstance(xt._process_alive(999999), bool)


# ---------------------------------------------------------------------------
# pack_xlam（纯 zip 操作，不需要 Excel）
# ---------------------------------------------------------------------------
class TestPackXlam(unittest.TestCase):

    def _pkg(self, d, with_ct=True):
        os.makedirs(os.path.join(d, "customUI"), exist_ok=True)
        with open(os.path.join(d, "customUI", "customUI.xml"), "w", encoding="utf-8") as f:
            f.write("<customUI/>")
        if with_ct:
            with open(os.path.join(d, "[Content_Types].xml"), "w", encoding="utf-8") as f:
                f.write("<Types/>")
        return d

    def test_content_types_is_first_entry(self):
        # Office 要求 [Content_Types].xml 是归档里的第一个条目
        import tempfile
        import zipfile
        from xlam_toolkit import pack_xlam
        with tempfile.TemporaryDirectory() as d:
            src = self._pkg(os.path.join(d, "pkg"))
            out = os.path.join(d, "out.xlam")
            pack_xlam(src, out)
            with zipfile.ZipFile(out) as z:
                self.assertEqual(z.namelist()[0], "[Content_Types].xml")

    def test_output_inside_source_dir_is_not_self_packed(self):
        # 回归点：输出路径落在 source_dir 内时，输出文件曾被递归打进自己
        import tempfile
        import zipfile
        from xlam_toolkit import pack_xlam
        with tempfile.TemporaryDirectory() as d:
            self._pkg(d)
            out = os.path.join(d, "out.xlam")
            pack_xlam(d, out)
            with zipfile.ZipFile(out) as z:
                names = z.namelist()
            self.assertNotIn("out.xlam", names)
            self.assertIn("customUI/customUI.xml", names)

    def test_missing_content_types_warns(self):
        # 回归点：缺 [Content_Types].xml 时曾静默产出 Excel 打不开的包
        import tempfile
        import io
        import contextlib
        from xlam_toolkit import pack_xlam
        with tempfile.TemporaryDirectory() as d:
            self._pkg(d, with_ct=False)
            with tempfile.TemporaryDirectory() as d2:
                buf = io.StringIO()
                with contextlib.redirect_stdout(buf):
                    pack_xlam(d, os.path.join(d2, "out.xlam"))
                self.assertIn("warning", buf.getvalue())

    def test_refuses_output_equals_vba_source(self):
        # 输出与 vba_source 同路径会先删掉输出→丢掉 VBA 源，必须拒绝
        import tempfile
        from xlam_toolkit import pack_xlam
        with tempfile.TemporaryDirectory() as d:
            src = self._pkg(os.path.join(d, "pkg"))
            same = os.path.join(d, "same.xlam")
            with self.assertRaises(ValueError):
                pack_xlam(src, same, vba_source=same)


# ---------------------------------------------------------------------------
# 2026-09 五视角审查修复的回归测试
# ---------------------------------------------------------------------------

class TestXmlEscaping(unittest.TestCase):
    """generate_button_xml / generate_group_xml 属性转义——label 含 & 或 "
    曾产出非法 customUI.xml，Excel 拒载整个功能区（三个审查视角独立发现）。"""

    def test_button_label_escaped(self):
        from xlam_toolkit import generate_button_xml
        xml = generate_button_xml("b1", '复制 & 粘贴"全部', "My_Click")
        self.assertIn("&amp;", xml)
        self.assertIn("&quot;", xml)
        # 转义后不应存在裸 &（所有 & 都必须是实体起始）
        import re
        for m in re.finditer(r"&(?!amp;|lt;|gt;|quot;|apos;|#)", xml):
            self.fail(f"未转义的 & at {m.start()}: {xml}")

    def test_group_label_escaped(self):
        from xlam_toolkit import generate_group_xml
        xml = generate_group_xml("g1", "AT&T 工具", [])
        self.assertIn("AT&amp;T", xml)

    def test_plain_values_unchanged(self):
        from xlam_toolkit import generate_button_xml
        xml = generate_button_xml("b1", "普通按钮", "My_Click", screentip="提示")
        self.assertIn('label="普通按钮"', xml)


class TestProcTypeOfDecl(unittest.TestCase):
    """list_procedures 曾用 startswith 把 'Public Function ...' 全部误判为 Sub。"""

    def test_public_function(self):
        from xlam_toolkit import _proc_type_of_decl
        self.assertEqual(
            _proc_type_of_decl("Public Function AddTwo(a, b) As Long"), "Function")

    def test_private_property_get(self):
        from xlam_toolkit import _proc_type_of_decl
        self.assertEqual(
            _proc_type_of_decl("Private Property Get Name() As String"), "Property")

    def test_static_sub(self):
        from xlam_toolkit import _proc_type_of_decl
        self.assertEqual(_proc_type_of_decl("Static Sub Tick()"), "Sub")

    def test_bare_function(self):
        from xlam_toolkit import _proc_type_of_decl
        self.assertEqual(_proc_type_of_decl("Function F() As Long"), "Function")

    def test_unknown_falls_back_to_sub(self):
        from xlam_toolkit import _proc_type_of_decl
        self.assertEqual(_proc_type_of_decl("' comment"), "Sub")


class TestProcDeclReBoundary(unittest.TestCase):
    """add_form_event_handler 判重曾用裸子串：已存在 btnOK_Click2 时
    btnOK_Click 被误判"已存在"而静默跳过，事件永不触发。"""

    def test_no_false_positive_on_longer_name(self):
        from xlam_toolkit import _proc_decl_re
        code = "Private Sub btnOK_Click2()\nEnd Sub\n"
        self.assertIsNone(_proc_decl_re("btnOK_Click").search(code))

    def test_matches_with_modifiers(self):
        from xlam_toolkit import _proc_decl_re
        code = "Private Sub btnOK_Click()\nEnd Sub\n"
        self.assertIsNotNone(_proc_decl_re("btnOK_Click").search(code))

    def test_not_matching_comment_mention(self):
        from xlam_toolkit import _proc_decl_re
        code = "' TODO: wire btnOK_Click later\nSub Other()\nEnd Sub\n"
        self.assertIsNone(_proc_decl_re("btnOK_Click").search(code))


class TestFormInitHandler(unittest.TestCase):
    """曾生成 '{form_name}_Initialize'——UserForm 的 Initialize 事件过程名
    固定为 UserForm_Initialize（与窗体名无关），带窗体名的版本是永不触发
    的死代码。"""

    def test_fixed_event_name(self):
        from xlam_toolkit import generate_form_init_handler
        code = generate_form_init_handler("frmMain")
        self.assertIn("Private Sub UserForm_Initialize()", code)
        self.assertNotIn("frmMain_Initialize", code)

    def test_additem_via_me(self):
        from xlam_toolkit import generate_form_init_handler
        code = generate_form_init_handler(
            "frmMain", {"cmbType": {"AddItem": ["a", "b"]}})
        self.assertIn('Me.cmbType.AddItem "a"', code)


class TestMotwGate(unittest.TestCase):
    """带 Zone.Identifier（网络下载标记）的文件拒绝以宏启用方式打开——
    隔离实例不是沙箱，Workbook_Open/Auto_Open 以用户完整权限执行。"""

    def _touch(self, d, name="book.xlsm"):
        p = os.path.join(d, name)
        with open(p, "wb") as f:
            f.write(b"PK\x03\x04")
        return p

    def test_no_marker_passes(self):
        import tempfile
        from xlam_toolkit import _has_mark_of_the_web, _refuse_untrusted
        with tempfile.TemporaryDirectory() as d:
            p = self._touch(d)
            self.assertFalse(_has_mark_of_the_web(p))
            _refuse_untrusted(p)  # 不抛即通过

    def test_marker_refused_then_trusted_ok(self):
        import tempfile
        from xlam_toolkit import _has_mark_of_the_web, _refuse_untrusted
        with tempfile.TemporaryDirectory() as d:
            p = self._touch(d)
            with open(p + ":Zone.Identifier", "w") as f:
                f.write("[ZoneTransfer]\nZoneId=3\n")
            self.assertTrue(_has_mark_of_the_web(p))
            with self.assertRaises(ValueError):
                _refuse_untrusted(p)
            _refuse_untrusted(p, assume_trusted=True)  # 显式信任放行

    def test_missing_file_no_marker(self):
        from xlam_toolkit import _has_mark_of_the_web
        self.assertFalse(_has_mark_of_the_web(r"C:\__no_such_file__.xlsm"))


class TestUnpackGuard(unittest.TestCase):
    """unpack_xlam 曾对已存在目录无条件 rmtree——误传父目录即整目录销毁。
    护栏：只清空含 [Content_Types].xml 的解包产物或空目录。"""

    def _zip(self, path):
        import zipfile
        with zipfile.ZipFile(path, "w") as z:
            z.writestr("[Content_Types].xml", "<Types/>")
            z.writestr("xl/workbook.xml", "<wb/>")

    def test_refuses_foreign_dir(self):
        import tempfile
        from xlam_toolkit import unpack_xlam
        with tempfile.TemporaryDirectory() as d:
            victim = os.path.join(d, "victim")
            os.makedirs(victim)
            with open(os.path.join(victim, "keepme.txt"), "w") as f:
                f.write("data")
            z = os.path.join(d, "src.xlam")
            self._zip(z)
            with self.assertRaises(ValueError):
                unpack_xlam(z, victim)
            # 目录原样保留，未被清空
            self.assertTrue(os.path.exists(os.path.join(victim, "keepme.txt")))

    def test_allows_empty_and_unpack_dirs(self):
        import tempfile
        from xlam_toolkit import unpack_xlam
        with tempfile.TemporaryDirectory() as d:
            z = os.path.join(d, "src.xlam")
            self._zip(z)
            # 空目录放行（mkdtemp 预建目录场景，render_ribbon_preview 内部用）
            empty = os.path.join(d, "empty")
            os.makedirs(empty)
            unpack_xlam(z, empty)
            self.assertTrue(
                os.path.exists(os.path.join(empty, "[Content_Types].xml")))
            # 已是解包产物（重复解包）放行
            unpack_xlam(z, empty)


class TestPackBackupOnOverwrite(unittest.TestCase):
    """pack_xlam 覆盖已有输出曾直接 os.remove 无备份，与模块头部
    "Every write operation backs the file up" 的宣称不符。"""

    def test_existing_output_backed_up(self):
        import tempfile
        from xlam_toolkit import pack_xlam, list_backups
        with tempfile.TemporaryDirectory() as d:
            src = os.path.join(d, "pkg")
            os.makedirs(os.path.join(src, "customUI"), exist_ok=True)
            with open(os.path.join(src, "customUI", "customUI.xml"), "w",
                      encoding="utf-8") as f:
                f.write("<customUI/>")
            with open(os.path.join(src, "[Content_Types].xml"), "w",
                      encoding="utf-8") as f:
                f.write("<Types/>")
            out = os.path.join(d, "out.xlam")
            with open(out, "wb") as f:
                f.write(b"OLD-CONTENT")
            pack_xlam(src, out)
            self.assertTrue(list_backups(out))  # 旧内容进了备份目录
            bak = os.path.join(out + ".bak", list_backups(out)[0])
            with open(bak, "rb") as f:
                self.assertEqual(f.read(), b"OLD-CONTENT")


class TestButtonIdDedup(unittest.TestCase):
    """add_button_to_ribbon 曾不查重 id——AI 重试追加同 id 按钮后产出
    重复 id 的 customUI，Excel 拒载整个功能区。"""

    def _dir(self, d):
        os.makedirs(os.path.join(d, "customUI"), exist_ok=True)
        with open(os.path.join(d, "customUI", "customUI.xml"), "w",
                  encoding="utf-8") as f:
            f.write('<customUI><ribbon><tabs><tab id="tabCustom">'
                    '<group id="grp"><button id="b1" label="old"/>'
                    '</group></tab></tabs></ribbon></customUI>')
        return d

    def test_duplicate_id_skipped(self):
        import tempfile
        from xlam_toolkit import add_button_to_ribbon, get_ribbon_xml
        with tempfile.TemporaryDirectory() as d:
            self._dir(d)
            add_button_to_ribbon(d, "grp", '<button id="b1" label="new"/>')
            xml = get_ribbon_xml(d)
            self.assertEqual(xml.count('id="b1"'), 1)
            self.assertNotIn('label="new"', xml)

    def test_new_id_appended(self):
        import tempfile
        from xlam_toolkit import add_button_to_ribbon, get_ribbon_xml
        with tempfile.TemporaryDirectory() as d:
            self._dir(d)
            add_button_to_ribbon(d, "grp", '<button id="b2" label="second"/>')
            xml = get_ribbon_xml(d)
            self.assertEqual(xml.count('id="b1"'), 1)
            self.assertEqual(xml.count('id="b2"'), 1)


# ---------------------------------------------------------------------------
# 端到端（需 Excel，默认跳过）
# ---------------------------------------------------------------------------
@unittest.skipUnless(os.environ.get("VBA_DEV_EXCEL_TESTS") == "1",
                     "端到端用例需 Excel：设 VBA_DEV_EXCEL_TESTS=1 启用")
class TestExcelEndToEnd(unittest.TestCase):
    """会启动隔离 Excel 实例（不触碰用户自己打开的 Excel）。"""

    def _make_addin(self, tmpdir, module_code):
        """建一个临时 .xlam 并写入给定模块，返回路径。"""
        import xlam_toolkit as xt
        path = os.path.join(tmpdir, "smoke.xlam")
        xt.create_addin(path)
        xt.add_vba_module(path, "modSmoke", module_code)
        return path

    def test_early_binding_blocks_build(self):
        import tempfile
        with tempfile.TemporaryDirectory() as d:
            p = self._make_addin(d, "Sub S()\n    Dim d As New Scripting.Dictionary\nEnd Sub")
            res = build_check(p, compile=False)
            self.assertFalse(res["ok"])
            self.assertTrue(any(e.get("kind") == "early_binding"
                                for e in res["static"]["errors"]))

    def test_late_binding_passes(self):
        import tempfile
        with tempfile.TemporaryDirectory() as d:
            p = self._make_addin(
                d,
                'Sub S()\n    Dim d As Object\n    Set d = CreateObject("Scripting.Dictionary")\nEnd Sub')
            res = build_check(p, compile=False)
            self.assertTrue(res["ok"], res["static"]["errors"])

    def test_check_binding_can_be_disabled(self):
        import tempfile
        with tempfile.TemporaryDirectory() as d:
            p = self._make_addin(d, "Sub S()\n    Dim d As New Scripting.Dictionary\nEnd Sub")
            res = build_check(p, compile=False, check_binding=False)
            self.assertFalse(any(e.get("kind") == "early_binding"
                                 for e in res["static"]["errors"]))

    # --- 真实 VBE 编译层（compile=True）-----------------------------------
    # 上面的用例都跑 compile=False，只覆盖静态层；编译探针是本工具最复杂、
    # 最容易"假通过"的一环，必须有真用例压着。

    def test_compile_layer_passes_clean_code(self):
        import tempfile
        with tempfile.TemporaryDirectory() as d:
            p = self._make_addin(
                d,
                "Option Explicit\n\n"
                "Public Function AddTwo(a As Long, b As Long) As Long\n"
                "    AddTwo = a + b\n"
                "End Function")
            res = build_check(p, compile=True)
            self.assertTrue(
                res["ok"],
                f"干净代码应编译通过；static={res['static']['errors']} compile={res['compile']}")

    def test_compile_layer_catches_syntax_error(self):
        import tempfile
        with tempfile.TemporaryDirectory() as d:
            # 静态检查不会报这个（Dim 在过程内），只有真编译能发现
            p = self._make_addin(d, "Sub Broken()\n    Dim\nEnd Sub")
            res = build_check(p, compile=True)
            comp = res["compile"]
            self.assertIsNotNone(comp, "未触发编译探针")
            self.assertIs(comp.get("ok"), False,
                          f"真编译未捕获语法错误（假通过）: {comp}")
            self.assertFalse(res["ok"])
            self.assertTrue(comp.get("error_text") or comp.get("timed_out"),
                            f"应带错误文本: {comp}")

    # --- run_test(keep=True)：曾走 write 模式在 OPEN 前强禁宏必败 ------------
    def test_run_test_keep_true_runs_and_saves(self):
        import tempfile
        import xlam_toolkit as xt
        with tempfile.TemporaryDirectory() as d:
            p = self._make_addin(
                d,
                "Public Function Ping() As String\n"
                "    Ping = \"pong\"\n"
                "End Function")
            r = xt.run_test(
                p,
                "Public Function RunTest() As String\n"
                "    RunTest = \"k=\" & Ping()\n"
                "End Function",
                keep=True)
            self.assertTrue(r["ok"], f"keep=True 应能执行: {r['error']}")
            self.assertEqual(r["result"], "k=pong")
            # keep=True 保存了注入模块（此前 write 模式下 Run 必报"宏被禁用"）
            code = xt.read_vba_module(p, r["module"])
            self.assertIn("RunTest", code)

    def test_run_test_default_leaves_file_unchanged(self):
        import tempfile
        import xlam_toolkit as xt
        with tempfile.TemporaryDirectory() as d:
            p = self._make_addin(
                d,
                "Public Function Ping() As String\n"
                "    Ping = \"pong\"\n"
                "End Function")
            r = xt.run_test(p, "Public Function RunTest() As String\n"
                               "    RunTest = \"k=\" & Ping()\n"
                               "End Function")
            self.assertTrue(r["ok"], r["error"])
            self.assertEqual(r["result"], "k=pong")
            # 默认 keep=False：注入模块跑完即删，文件不含测试模块
            self.assertNotIn("RunTest", xt.read_vba_module(p, "modSmoke"))

    def tearDown(self):
        import xlam_toolkit as xt
        xt.shutdown_excel()


if __name__ == "__main__":
    unittest.main(verbosity=2)
