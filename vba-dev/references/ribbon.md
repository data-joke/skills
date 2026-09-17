# 功能区（Ribbon）参考：XML 写法 / 回调模式 / 图标管线

> 由 SKILL.md 按需引用。**实际动手写 customUI.xml、写回调、或做图标时**读本文件。
> 工作流顺序（unpack → init → 改 → pack）在 SKILL.md。

## 按钮基本结构

```xml
<button id="btnMyButton"
        label="我的按钮"
        size="large"
        imageMso="Copy"
        onAction="MyCallback"
        screentip="提示文本"
        supertip="详细说明"/>
```

`size` 省略即小按钮（Office 默认行为）；`size="large"` 才是大图标按钮。

## 常用内置图标 imageMso（已在 Office 16 进程内实测验证 2026-09）

| 类别 | 有效图标名 |
|------|-----------|
| 编辑 | `Copy` `Cut` `Paste` `PasteValues` `Undo` `Redo` `Clear` `ClearContents` `Delete` `FormatPainter` `Spelling` |
| 格式 | `Font` `FontDialog` `Bold` `Underline` `WrapText` `MergeCells` `AlignLeft` `TextBoxInsert` |
| 数据 | `Sort` `Filter` `AutoSum` `Refresh` `RefreshAll` `TableInsertDialog` `InsertTable` `InsertChart` `ChartInsert` |
| 其他 | `Save` `HyperlinkInsert` `Camera` `Calculator` `CalendarInsert` `VisualBasic` `MacroPlay` `AddInManager` |

⚠️ **以下常见名字实测无效（按钮会不显示图标）**：`Replace` `Find` `Print` `Symbol` `AdvancedFilter` `SortAscending` `SortDescending` `Subtotal` 等。
名单外/存疑的名字先用 `validate_imagemso()` 验证（进程内探针，可靠）；也可直接读运行时的 `COMMON_IMAGEMSO`，不必背这份表。

## 分组 / 标签页结构

```xml
<tab id="tabAI" label="数你皮">
  <group id="grpMyGroup" label="我的工具">
    <button id="btnFirst" label="第一个" size="large" imageMso="Copy" onAction="First_Click"/>
  </group>
</tab>
```

**改已有标签页的 label（如 `init_custom_ui` 默认的 `tabCustom`/「自定义」）**：必须做**属性级替换**，不能拿 `str.replace('<tab id="tabCustom"', '<tab id="tabCustom" label="新名"')` 往标签上再插一个 label——同一标签出现两个 `label` 属性是非法 XML，功能区整体失效。安全写法：

```python
import re
p = os.path.join(xlam_dir, "customUI", "customUI.xml")
xml = open(p, encoding="utf-8").read()
xml, n = re.subn(r'(<tab id="tabCustom"[^>]*?)\s+label="[^"]*"', r'\1 label="效率工具"', xml)
if n == 0:   # 原标签没有 label 属性，才补插
    xml = xml.replace('<tab id="tabCustom"', '<tab id="tabCustom" label="效率工具"')
open(p, "w", encoding="utf-8").write(xml)
```

改完任何 XML 手改都必须 `render_ribbon_preview` 验证（解析失败会直接报错，别跳过这步）。

## 其他功能区控件

```xml
<toggleButton id="tglMode" label="模式" onAction="ToggleMode_Click"/>
<dropDown id="ddOptions" label="选项" onChange="Options_Changed">
  <item id="opt1" label="选项 1"/>
</dropDown>
<comboBox id="cmbInput" label="输入" onChange="Input_Changed"/>
<menu id="mnuTools" label="工具" size="large" imageMso="Copy">
  <button id="mnuTool1" label="工具 1" onAction="Tool1_Click"/>
  <separator id="sep1"/>
</menu>
<splitButton id="splRecent" label="最近使用">
  <button id="btnRecent1" label="最近 1" onAction="Recent1_Click"/>
  <menu id="splMenu"/>
</splitButton>
<gallery id="galColors" label="颜色" columns="4" onAction="Color_Selected"/>
```

动态回调：`getLabel` / `getEnabled` / `getVisible` / `getImage` 属性 + 对应 VBA 过程（见下方回调模式）。

## VBA 回调模式

### 先接线 onLoad（否则刷新类回调全部失效）

动态回调（`getLabel`/`getVisible`/`getImage`）的返回值在功能区**加载时求值并缓存**，
之后要 `IRibbonUI.Invalidate` 才重新求值——而拿到 `IRibbonUI` 的唯一途径是根元素
的 `onLoad`（`init_custom_ui` 生成的骨架**不带**它，需要动态刷新时自己补上）：

```xml
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui"
          onLoad="InitializeRibbon">
```

不接 `onLoad`，`ribUI` 恒为 `Nothing`，`RefreshRibbon` 一调就运行时错误 91。
纯静态按钮（无 get* 动态属性）不需要这步。

### 回调签名与模板

```vba
' 按钮回调——On Error 必须在回调内吃掉错误，否则用户点按钮会弹 VBA 调试框
Sub ButtonName_Click(control As IRibbonControl)
    On Error GoTo ErrHandler
    MsgBox "按钮被点击了！"
    Exit Sub
ErrHandler:
    MsgBox "执行失败: " & Err.Description, vbExclamation
End Sub

' 切换按钮（注意块形式：单行 If 到行尾结束，"If x Then Else End If" 编译不过）
Sub ToggleMode_Click(control As IRibbonControl, pressed As Boolean)
    If pressed Then
        ' TODO: 开启模式
    Else
        ' TODO: 关闭模式
    End If
End Sub

' 动态标签/可见性
Sub GetButtonLabel(control As IRibbonControl, ByRef label)
    label = "动态标签"
End Sub
Sub GetButtonVisible(control As IRibbonControl, ByRef visible)
    visible = True
End Sub

' 下拉框数量/项目/选择
Sub GetItemCount(control As IRibbonControl, ByRef count): count = 3: End Sub
Sub GetItemLabel(control As IRibbonControl, index As Integer, ByRef label): label = "选项 " & index: End Sub
' ⚠️ onChange 签名按控件区分，抄错即参数不匹配、运行时报错：
'   dropDown  → (control, id As String, index As Integer)
'   comboBox  → (control, text As String)   ← 收到的是用户输入的文本
Sub Options_Changed(control As IRibbonControl, id As String, index As Integer)
End Sub
Sub Input_Changed(control As IRibbonControl, text As String)
End Sub

' 刷新功能区缓存（依赖上方 onLoad 接线）
Dim ribUI As IRibbonUI
Sub InitializeRibbon(ribbon As IRibbonUI): Set ribUI = ribbon: End Sub
Sub RefreshRibbon(): ribUI.Invalidate: End Sub
```

> `IRibbonControl` / `IRibbonUI` 来自 Office 库，是 Excel 工程的默认引用，**不需要后期绑定**（见 `binding-rules.md`）。

## 自定义图标管线（预检 → 提示词 → 生成 → 尺寸规整 → 嵌入）

```
check_icon_api() 预检（≤2 行诊断：缺配置/key 无效/URL 不通，fail-fast）
  ├─ 通过 → [提示词（按下方规范）] → generate_icon()（图像 API）
  └─ 未配置/不可用 → draw_icon() 离线兜底（勿反复重试 API）
                       ↓
              prepare_icon()（1024→32×32 圆角，防大图毁按钮）
                       ↓
              register_icon() / add_icon_button()（嵌入）
```

**AI 使用规则**：调 `generate_icon()` 前先 `check_icon_api()`；预检失败或生成报错时，**直接转 `draw_icon()` 兜底，不要重试 API**——错误消息已含完整修复指引（缺什么、配置示例、验证方法），照做即可。**即使用户点名要 AI 图标也先兜底交付，交付时必须说明**：本次因图像 API 未配置/不可用改用了字符图标；并附一句配置方法（`set_icon_config(base_url=..., api_key=...)`，key 建议用环境变量 `ICON_API_KEY`），说明这是一次性配置、配好后以后可随时重新生成 AI 图标替换。

**图标提示词规范**（AI 生成提示词时遵循，保证同插件风格统一）：
- 扁平简约单色/双色线条图标，粗描边，几何造型，32px 下可辨识
- **透明背景**（或纯色背景+圆角），画布居中占 ~70%
- 与插件主题色一致（如 #1F4E92）；一组图标用同一种风格
- 负向：照片写实、复杂渐变、阴影、图标内文字、过多细节

⚠️ **实测（2026-09）：图像模型基本会无视提示词里的「透明背景」，输出白底 RGB 图**（三个图标全部如此，透明只剩圆角 ~4%）。提示词照写（偶有效果），但**可靠兜底只有 `white_to_alpha=True`**——AI 生成的图标嵌入时一律带上它。

**配置图像 API**（OpenAI 兼容 `/v1/images/generations`，国内多数服务商/网关可用）。只需 `base_url + api_key` 两项，**model 默认 `gpt-image-1`**（可随时 `set_icon_config(model=...)` 覆盖）：

```python
set_icon_config(base_url="https://your-gateway/v1", api_key="sk-...")   # model 自动默认
get_icon_config()        # 回显（key 自动打码）
check_icon_api()         # 预检连通性（生成前必调）
```

**配置存储与发布安全**：配置文件存于 `~/.vba-dev/icon_config.json`——**在 skill 目录之外**，因此把 skill 上传 GitHub 等永远不会携带真实配置；`api_key` 优先用环境变量 `ICON_API_KEY`（完全不落盘）。skill 内仅随附无敏感信息的 `scripts/icon_config.example.json` 字段说明模板 + `.gitignore` 双保险。

**典型用法**：

```python
# 有 API：从提示词到按钮一步成型（out_png 必填：生成图的存盘路径）
src = generate_icon("扁平风格橡皮擦图标，透明背景，主题蓝色 #1F4E92", "icon_eraser.png")
# 无 API：离线兜底（圆角底 + 1-2 个中英文字符；out_png 同样必填）
src = draw_icon("清", "icon_clean.png", bg=(31, 78, 146))
# 嵌入（自动规整为 32×32 圆角 + 注册 + 加按钮，一条调用）
# AI 生成的图标必须传 white_to_alpha=True（模型会无视透明背景要求输出白底）
add_icon_button(xlam_dir, "grpQuick", "btnClean", "清理文本",
                on_action="CleanSelection_Click", icon_png=src, white_to_alpha=True,
                screentip="清理选区文本")
# 手动分步：prepare_icon(src, out, size=32, rounded=True, white_to_alpha=True)
#           + register_icon(xlam_dir, icon_id, out) + add_button_to_ribbon(image=...)
```

## 预览检查（改完必看）

功能区装进 Excel 之前完全不可见。改完 customUI 后渲染一张模拟图确认布局：

```python
render_ribbon_preview("DataJoke/", "preview.png")        # 解包目录
render_ribbon_preview("MyTools.xlam", "preview.png")     # 也可直接给 .xlam
render_ribbon_preview("DataJoke/", "preview.png", tab_index=1)   # 看第 2 个标签页
```

能看出：按钮有没有落进正确的分组、label 会不会被截断、图标有没有生效（未注册的图标画成灰色占位方块）、空分组。
**交付前把 preview.png 给用户看**——这是功能区唯一可见的验收材料。
