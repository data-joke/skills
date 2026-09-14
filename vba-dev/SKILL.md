---
name: vba-dev
description: "开发与修改 Excel 插件/加载项及宏工作簿（.xlam / .xlsm）：编写 VBA 模块与回调、定制功能区（Ribbon / customUI）、加图标、创建用户窗体，并在隔离的 Excel 实例中做真实编译验证与运行测试。当用户要做 Excel 插件/加载项、给插件加按钮或功能区、写/改/调试 VBA 宏代码、加图标或窗体、或升级已安装的插件（文件被锁无法直接修改）时使用。"
---

# VBA 开发技能 - Excel 文件操作

## 能力总览（签名/参数/返回值查 `references/function-reference.md`）

- **环境**：`check_environment` / `enable_vba_trust` / `shutdown_excel`
- **源码同步**：`create_addin`（0→1 空 `.xlam`）；`export_vba_source` / `import_vba_source`（VBA ↔ .bas/.cls，AI 编辑源码主通道）
- **验证**：`build_check`（静态+真编译）；`run_macro` / `run_test`（注入临时宏，文件不变）
- **VBA/窗体**：`add_vba_module` 等 13 个过程级函数；`create_userform` / `lint_form`
- **Ribbon**：customUI 全链路 + `render_ribbon_preview`（装进 Excel 前唯一可见的验收材料）
- **备份/升级/偏好**：写操作自动备份、`restore_backup` 回滚；已装插件升级见「升级工作流」；`get_preferences` / `set_preference`

## 执行契约（AI 调用方式，务必遵守）

本工具是纯 Python 单文件，用 Bash 调用。**关键约束：Excel 实例缓存在 Python 进程内，
进程退出时 `atexit` 关闭它——每次 Bash 调用都是一次全新的 Excel 冷启动（秒级开销），
而同进程内追加操作近乎免费。首要规则：把多个操作合并进同一个脚本。**

```python
# 一个脚本跑完 → 只开 1 次 Excel
import sys, os
SKILL_DIR = r"<本技能根目录>"   # 用技能加载方公布的 base directory（即本 SKILL.md 所在目录，勿写死绝对路径）
sys.path.insert(0, os.path.join(SKILL_DIR, "scripts"))
from xlam_toolkit import create_addin, add_vba_module, build_check, run_test

create_addin("MyTools.xlam")
add_vba_module("MyTools.xlam", "modA", CODE_A)
add_vba_module("MyTools.xlam", "modB", CODE_B)
build_check("MyTools.xlam")
run_test("MyTools.xlam", TEST_CODE)
```

- 需要中途等用户确认时：**读操作**合并成一次调用，**写操作**集中成一次调用
- 控制台为 GBK 时先 `set PYTHONIOENCODING=utf-8`；**长脚本一律落盘成 .py 文件再执行**——多行 `python -c` 在 GBK 控制台可能输出被整体吞掉且退出码仍为 0，造成"假成功"误判
- **批量脚本中途抛异常**：已完成的写操作各自落盘、不回滚。要退掉整批，用 `list_backups()` 找到**批次开始前**那份再 `restore_backup(file, backup="...")`——⚠️ 不带参数恢复的是**最后一次写操作前**，不是整批之前。

## 安全机制（务必了解）

- **独立 Excel 实例**：COM 全在 `DispatchEx` 后台 Excel 中进行，**绝不动用户打开的 Excel**；`attach=True` 可显式复用用户实例，但**永不 Quit、永不强杀用户 Excel**
- **自动备份**：所有写操作前自动备份；`restore_backup()` 恢复（恢复前也先备份当前状态，可逆）
- **弹窗看门狗**：模态弹窗（MsgBox/编译错/运行时错）自动关闭、错误文本完整捕获——VBA 运行时错误的 `com_error` 本身不带描述，**真实错误信息在结果 dict 的 `error`/`dialogs` 字段里**
- **卡死保护 + 失败原子性**：宏卡死超 `timeout` 秒或弹窗风暴时强杀**本工具的隔离实例**并报告（不影响用户 Excel）；写操作中途抛异常则不保存关闭，文件保持原状

## 环境准备（首次使用时执行一次）

```python
env = check_environment()                  # {pywin32, excel_installed, vba_trust_enabled, ok}
if not env["vba_trust_enabled"]:
    enable_vba_trust()                     # 写注册表 AccessVBOM=1，对新启动的 Excel 生效
```

依赖：`pip install pywin32 Pillow`（AI 生成图标另需 `requests`）。

---

## 交互引导规范（多可能性决策点：先征询、带推荐、再实施）

**原则**：凡用户未明确指定、且存在多种合理方案并影响最终体验的决策点，**先征询再实施**——用 AskUserQuestion 罗列约 3 个选项，标注 ⭐推荐并附一句理由，保留「自定义输入」出口（AskUserQuestion 的 "Other" 天然支持；纯文本征询时显式写明「以上都不满意？直接说您的偏好」）。**相互独立的决策合并进同一次征询**（单次最多 4 个问题，不拆轮串行）；只有前一答案会改变后一问题选项时才分轮。用户说「你定/随便」直接采用推荐项，不再追问。

**自定义输入**：具体值（"用扫帚图标"）直接采用；偏好描述（"我想要极简风格"）转化为 1-2 个符合偏好的具体方案，一次确认后执行（收窄即止），并适用于后续同类决策。

**新增功能区按钮的图标/命名/分组三决策选项模板见 `references/interaction.md`**（先 `get_ribbon_xml()` 读现状，三问合一，一次问全）。

### 记住用户偏好（跨会话）

开工做图标/功能区前先 `get_preferences()` 读一次（纯 JSON，不启 Excel），有值直接沿用并说明；
用户明确表达风格偏好时 `set_preference()` 顺手落盘（约定键见 `references/function-reference.md`），
改了口径就覆盖——**不替用户臆测**。同类风格调整发生第二次时，主动问一句「以后都按这个风格来？」，确认后落盘。

### 问与不问的边界

- **要问**：宏实现方式、窗体布局与尺寸、目标文件格式（xlam vs xlsm）、图标主题色——凡多方案且用户未指定
- **不问**：只影响内部实现的细节（如模块拆分）；用户指令已含明确值——直接执行，交付时一句话说明可调整处
- **批量操作**（一次加 N 个按钮）只征询一次统一风格，确认后批量执行；**小改动/修 bug**（换图标、改 label）直接做

---

## 工作流路由（按用户需求选路，可叠加）

| 用户需求 | 走哪条 |
|---------|--------|
| 从零做一个新插件 | `create_addin()` 建空 `.xlam` → 写 VBA（示例见「执行契约」）→ 需要功能区接 Ribbon 工作流。**全程用同一个 `.xlam` 路径**——中途写成 `.xlsm` 是另一个文件，会各改各的 |
| 改现有文件，涉及多模块/批量改动 | 标准开发工作流（导出 → 编辑源码 → 导回） |
| 单点小改（加个过程、改个回调） | 快速路径（免导出） |
| 只动功能区/按钮/图标 | Ribbon 定制工作流 |
| 写/改文件时报「文件被锁定」 | 先走已装插件升级工作流释放锁，再回原工作流 |

任何工作流开工前先过一眼「开发质量基线」：性能（数组化/状态恢复）、UI（对齐/主题色）、产品（首次使用动线）三条视角贯穿始终，不是收尾检查项。

## 标准开发工作流（推荐：多改动、批量任务）

**核心思想：让 AI 直接编辑文本文件（.bas/.cls），而不是逐段调函数注入**——改文件是 AI 的原生能力，最后一次性导回并验证。

```python
from xlam_toolkit import (
    export_vba_source, import_vba_source, build_check, run_test, run_macro
)

# 1. 导出全部 VBA 为源码（含 manifest.json，目录默认 <文件名>_vba/）
export_vba_source("DataJoke.xlsm")          # -> ./DataJoke_vba/modMain.bas 等

# 2. 用 Read/Edit/Write 直接编辑导出的 .bas/.cls 文件
#    （UTF-8 编码，中文注释往返安全；.frm 是窗体二进制，只读不动；
#     文档模块 Sheet1/ThisWorkbook.cls 可改代码）

# 3. 导回（写前自动备份；sync=True 删除工作簿里有但源码目录没有的模块/类/窗体）
import_vba_source("DataJoke.xlsm", "DataJoke_vba", sync=True)

# 4. 构建验证：静态结构检查 + 真实 VBE 编译
result = build_check("DataJoke.xlsm")
# result["ok"] / result["static"]["errors"] / result["compile"]["error_text"|"error_location"]

# 5. 冒烟测试：注入临时宏运行（文件不变，函数返回值可捕获）
r = run_test("DataJoke.xlsm", '''
Public Function RunTest() As String
    RunTest = "AddTwo=" & AddTwo(2, 3)
End Function
''')
# r["ok"], r["result"] == "AddTwo=5", r["error"], r["dialogs"]

# 6. 失败则修复后重复 2-5；需要回滚时：
# from xlam_toolkit import restore_backup, list_backups
# list_backups("DataJoke.xlsm");  restore_backup("DataJoke.xlsm")
```

**验证闭环规则**：任何改动后必须 `build_check` 通过（自己新写的模块传 `strict=True`，见代码风格指南）；涉及运行逻辑的再 `run_test` 冒烟。编译错误结果包含 `error_location = {component, line}`，直接定位修复；若报的是前期绑定（错误项带 `kind == "early_binding"`），按提示改成 `CreateObject` 后期绑定。

## 快速路径与运行（单点小改 / 跑宏 / 冒烟）

```python
add_vba_module("D.xlsm", "modF", CODE)      # 免导出单点编辑；另有 update_vba_callback /
read_vba_module("D.xlsm", "modF")           # replace_vba_lines / list_procedures 等 13 个（见 function-reference）
run_macro("D.xlsm", "AddTwo", args=[3, 4])  # 跑已有宏；save=True 时保存（写前自动备份）；Function 返回值进 result
run_test("D.xlsm", TEST_CODE, timeout=120)  # 注入临时宏：code 须定义 Public RunTest；结束自动删模块、默认不保存
build_check("D.xlsm")                       # 改完必查
```

弹窗由看门狗自动关闭、错误文本进 `error`/`dialogs`（机制见「安全机制」）。**测试代码不要写依赖人工交互的逻辑**——MsgBox 会被自动点掉，InputBox 只拿到空串。**结果含 NBSP/零宽字符等特殊字符时，控制台可能显示成 `?`**——勿轻信显示，用 `run_macro` 单独取回结果按字符码复核再下结论。

## 文件格式约定

| 格式 | 功能区支持 | VBA 操作 |
|------|-----------|---------|
| `.xlam` 加载项 | ✅（推荐，加载即生效） | 直接操作 |
| `.xlsm` 宏工作簿 | ⚠️ 技术可行（同 xlam 链路，打开工作簿时生效） | 直接操作 |
| `.xlsx` 普通工作簿 | ❌ | **写操作**自动转 `.xlsm`（原文件不变）；**读操作**报错（xlsx 无 VBA） |

---

## .xlam Ribbon 定制工作流（解包目录操作）

**执行本工作流新增按钮/分组前，先按「交互引导规范」完成决策征询**（图标 / 命名 / 分组位置，选项模板见 `references/interaction.md`，一次问全；用户已明确的项跳过）。
**XML 写法、回调模式、imageMso 清单、图标管线见 `references/ribbon.md`。**

```python
from xlam_toolkit import (
    unpack_xlam, init_custom_ui, pack_xlam,
    get_ribbon_xml, set_ribbon_xml,
    add_button_to_ribbon, add_group_to_ribbon, register_icon,
    generate_button_xml, generate_group_xml, render_ribbon_preview,
)

# 1. 解包（本质是 zip；.xlsm 同样可用此链路，技术上 xlsm 也支持 customUI）
unpack_xlam("DataJoke.xlam", "DataJoke/")

# 2. 从零初始化 customUI 全套设施（幂等，未定制过的包第一步必调）
init_custom_ui("DataJoke/")

# 3. 加分组/按钮（自闭合分组自动展开；tab_id 支持 id/idMso/idQ，可加到内置标签页）
add_group_to_ribbon("DataJoke/", '<group id="grpTools" label="工具"/>',
                    tab_id="tabCustom")            # 或 tab_id="TabHome"（内置页签）
add_button_to_ribbon("DataJoke/", "grpTools",
                     generate_button_xml(id="btnGo", label="执行",
                                         on_action="Go_Click", image_mso="Copy"))

# 4. 自定义图标（自动补 png 的 Content-Types 声明和 rels）
register_icon("DataJoke/", "my_icon", "path/to/icon.png")   # 16x16/32x32 透明 PNG

# 5. 预览检查 —— 功能区装进 Excel 前完全不可见，先渲染一张模拟图确认布局
render_ribbon_preview("DataJoke/", "preview.png")   # 交付前给用户看

# 6. 重新打包 —— 改过 VBA 就必须传 vba_source，否则旧 vbaProject.bin 会静默回退 VBA 改动
pack_xlam("DataJoke/", "DataJoke_new.xlam", vba_source="DataJoke.xlam")
```

**顺序规则**：`unpack → (改 XML/Ribbon) → (需要改 VBA 用 COM API 改原文件) → 预览 → pack(vba_source=原文件)`。输出路径不能与 vba_source 相同。

---

## 已安装插件的升级工作流（文件被用户 Excel 锁定时）

已加载的 `.xlam` 被用户正在运行的 Excel 锁定，直接写 VBA 或重新打包会失败。标准链路：**卸载释放锁 → 修改 → 重载**：

```python
from xlam_toolkit import is_file_locked, unload_addin, reload_addin
import os

if is_file_locked("MyTools.xlam"):      # 纯文件系统检测，不启 COM
    unload_addin("MyTools")             # 在用户 Excel 中卸载（默认 attach）

# ... 正常修改 VBA（add_vba_module 等直接写原文件）；改 Ribbon 时：
# unpack → 改 customUI → pack 到新路径 → os.replace 覆盖原文件

reload_addin("MyTools")                 # 重载（⚠️ 同会话卸载后需重启才真正恢复，见下）
```

⚠️ **本机构建（Office 16）实测：COM 方式安装/重载（`AddIns.Add` + `Installed=True`）只完成注册，不即时装载工程**。因此 `reload_addin` 返回的 `loaded=True` 只代表标志置位，**是否真正装载必须看 `workbook_loaded` 字段**（或用 `xl.Workbooks` 复核）；为 False 时让用户**重启 Excel** 或在加载项对话框手动勾选验收（已注册的加载项在 Excel 启动时会正常自动加载）。另：Backstage 界面（未开工作簿）的 Excel 不注册 ROT，`GetActiveObject` 附着不上。

要点：

- 文件锁由「加载它的实例」持有：用户 Excel 未运行 → 本就无锁直接改；运行中且加载了该插件 → 才需 `unload_addin()`（默认 attach 用户实例，永不 Quit）。它只取消勾选不删注册。
- 插件首次安装：`reload_addin("名字", path="完整路径")`。

---

## 排错速查（高频 5 条；完整 18 条见 `references/troubleshooting.md`）

| 症状 | 处理 |
|------|------|
| 文件被锁定（Open 失败） | 已装插件被用户 Excel 加载 → `is_file_locked()` 确认后 `unload_addin()` 释放锁，改完 `reload_addin()`；工具自身残留实例 → `shutdown_excel()` |
| `无法访问 Visual Basic 工程` | 信任中心未开 AccessVBOM → `enable_vba_trust()` 后**重启所有 Excel** |
| 按钮图标不显示 | imageMso 名无效 → `validate_imagemso()` 验证或换 `COMMON_IMAGEMSO` 名单内图标 |
| 重载/安装后功能区没出现 | COM 安装不即时装载工程 → 看 `workbook_loaded` 字段，为 False 让用户重启 Excel 验收 |
| 改坏了文件 | `list_backups()` + `restore_backup()` 回滚 |

## 开发质量基线（写代码与 UI 时自带的专业口径）

- **性能**：Range 批量读写数组化、循环外关 `ScreenUpdating`/`Calculation` 且出口必恢复（`references/vba-pitfalls.md` §2/§9）
- **UI**：`lint_form` 布局 warning 清零、配色一套主题色不超 3 色、主按钮 `Default`/`Cancel` 就位（口径见 `references/userform.md`）
- **产品动线**：交付前替用户走一遍首次使用路径——装好后第一次点击看到什么、screentip 是否说人话、出错文案是否告诉用户怎么办

## 代码风格指南

**VBA：** `Option Explicit`；类型前缀命名（`strName`/`rngSel`）；`On Error GoTo`；测试宏用 Function 返回结果字符串便于 `run_test` 捕获。
**自己新写的模块，`build_check` 一律传 `strict=True`**（拦截 Option Explicit 缺失与过程重名）。AI 生成 VBA 的高频坑（`Declare` 缺 `PtrSafe`、`Select` 模式、状态不恢复、插件语境 `ThisWorkbook` 误用等）见 `references/vba-pitfalls.md`——多数能通过编译、只在 64 位/插件/分发场景爆炸，写时自查。

**外部类型库一律后期绑定**（`CreateObject`；需要常量时在模块顶部 `Const` 本地重声明），完整口径见 `references/binding-rules.md`，`build_check` 会兜底拦截。

**Ribbon XML 与回调：** 优先 `imageMso` 内置图标；同时提供 `screentip` 和 `supertip`。**回调第一行必须 `On Error GoTo`，出口统一 `MsgBox` 报告**——裸奔的回调出错会给用户弹 VBA 调试框（分发场景大忌）。

**兼容基准：** Office 2016+；更新的能力（`XLOOKUP`、动态数组）默认不用，确需使用向用户注明版本要求。

**Python 调用：** 批量改动走「导出→编辑文件→导回」而非逐函数注入；每次写后 `build_check`；**多操作合并进同一个脚本**（见「执行契约」——分开调用每次都要冷启 Excel）。

**交付收尾：** 向用户报告文件路径、验证结果一句话（build_check/run_test）、preview.png 位置（有功能区时）、安装/升级方式（如 `reload_addin`）、一句话说明可调整处；**图标若因 API 未配置用了 `draw_icon` 兜底，必须说明并附 `set_icon_config` 配置方法**（见 `references/ribbon.md`）。收尾即走一遍「开发质量基线」的产品视角：替用户描述首次使用动线（装好后去哪点、点了看到什么）。

---

## 参考文档（按需读取，**不要预读**）

| 文件 | 什么时候读 |
|------|-----------|
| `references/function-reference.md` | 需要确认函数**签名/参数/返回值**时（本文档只给常用调用示例） |
| `references/ribbon.md` | 实际写 customUI.xml、写 Ribbon 回调、做图标时（XML 写法 / 回调模式 / imageMso 清单 / 图标管线 / 预览用法） |
| `references/binding-rules.md` | 写 VBA 前读一次（前期绑定禁令、ProgID 对照、白名单口径）；`build_check` 会兜底拦截 |
| `references/vba-pitfalls.md` | 写 VBA 前读一次（PtrSafe/64 位、插件语境、`Select` 模式、状态恢复等 AI 高频坑）；编译或运行报错时回来对照 |
| `references/userform.md` | 创建/修改用户窗体、查控件 ProgID 与事件名时 |
| `references/interaction.md` | 征询用户决策时（图标/命名/分组三决策的选项模板） |
| `references/troubleshooting.md` | 出错排查时（完整 18 条排错速查表） |
