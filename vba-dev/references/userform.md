# 用户窗体（UserForm）参考

> 由 SKILL.md 按需引用。**创建/修改窗体、查控件 ProgID 或事件名时**读本文件。

## 控件 ProgID 表

| 控件 | ProgID | 关键属性 |
|------|--------|---------|
| 标签 | `Forms.Label.1` | Caption |
| 文本框 | `Forms.TextBox.1` | Value、MultiLine、PasswordChar |
| 命令按钮 | `Forms.CommandButton.1` | Caption、Default、Cancel |
| 组合框 | `Forms.ComboBox.1` | Value、List |
| 列表框 | `Forms.ListBox.1` | Value、MultiSelect |
| 复选框 | `Forms.CheckBox.1` | Value (0/1/-1) |
| 单选按钮 | `Forms.OptionButton.1` | Value |
| 框架 | `Forms.Frame.1` | Caption（作 container 参数可嵌套放控件） |
| 图片 | `Forms.Image.1` | Picture、PictureSizeMode |
| 多页 | `Forms.MultiPage.1` | Pages（页内容放置见下文「多页窗体」） |
| 滚动条/数值调节 | `Forms.ScrollBar.1` / `Forms.SpinButton.1` | Min、Max、Value |
| 单元格选择 | `Forms.RefEdit.1` | Value |

## 用法

```python
add_control("f.xlsm", "frmMyDialog", "Forms.CommandButton.1", "btnOK",
            {"Caption": "确定", "Left": 10, "Top": 65, "Width": 80, "Default": True})
lint_form("f.xlsm", "frmMyDialog")   # 检查重叠/越界/负坐标
# 事件绑定：用 add_form_event_handler("f.xlsm", "frmMyDialog", "btnOK_Click", code) 写处理器。
# add_control 的 event_handler 参数对 MSForms 无效（MSForms 设计器控件没有 OnClick/OnChange
# 属性，那是 Access/VB6 约定），事件由 VBE 按"控件名_事件名"自动绑定，该参数已废弃不用。
```

常用事件：`Click`（按钮/复选/单选）、`Change`（文本框/组合框）、`Initialize`（窗体）。

## 多页窗体（MultiPage）

⚠️ `add_control(container=...)` 定位 MultiPage 时**必须配合 `page` 参数**：MSForms 的 MultiPage
在 IDispatch 层**不暴露 `Controls`**（实测 `IMultiPage` 上 `GetIDsOfNames("Controls")` 返回
"未知名称"），容器内容只能经 `Pages(i)` 访问。

```python
add_control("f.xlam", "frmMy", "Forms.MultiPage.1", "mpgMain",
            {"Left": 8, "Top": 6, "Width": 300, "Height": 200})
add_control("f.xlam", "frmMy", "Forms.Label.1", "lblP0",
            {"Caption": "第一页", "Left": 8, "Top": 6, "Width": 120, "Height": 14},
            container="mpgMain", page=0)   # page：页索引（0 起）或页名/页标题
```

默认只有 2 页；页数不够或改页标题时走一小段原生 COM：`mp.Pages.Add()`、
`mp.Pages(i).Caption = "..."`（`mp` 经 `VBComponents(form).Designer.Controls("mpgMain")` 取得）。

MultiPage 相关怪癖（均已实测）：

- `designer.Controls` 是**扁平集合**（含页内子控件，坐标为页内相对坐标）→ 旧版 lint_form
  会对 MultiPage 窗体产生大量"重叠" warning（warning 级不阻断）；新版已按父容器分组修复。
- `set_form_properties` 前若未实例化 `comp.Designer`，组件 Properties 写入报"发生意外"
  （工具已修；自己写 COM 时先访问一次 `comp.Designer`）。
- 设计态经 designer 设 Caption 报成功但不持久 → 运行时 `Me.Caption = "标题"` 兜底。
- ListBox 的 `ColumnWidths` 设计时设置报"类型不匹配" → 运行时赋值即可。

## 窗体设计怎么改

`.frm` 导出文件本身是**文本**（VERSION 头 + 控件几何 + 代码），伴生的 `.frx` 才是
**二进制**（图片等属性），且两者必须同步——所以不要试图用文本编辑器改窗体布局，走 COM API：
`add_control()` / `set_form_properties()` / `add_form_event_handler()` 增删控件与属性，
改完 `lint_form()` 检查几何，再 `build_check()` 验证。

`MSForms.*` 类型来自 **Microsoft Forms 2.0 Object Library**——含 UserForm 的工程会自动携带该引用，
接收方无需手动勾选，因此**不需要后期绑定**（见 `binding-rules.md`）。

## 窗体怎么验收（自动化的边界）

`lint_form` 只查几何；**模态窗体无法自动冒烟**——在 `run_test` 里写 `UserForm1.Show`
会一直挂到超时强杀（模态窗体不是 MsgBox，看门狗不认，也无预警）。验收口径：

- **逻辑抽出来测**：窗体背后的处理逻辑写成模块级纯函数（与事件处理器分离），
  `run_test` 直接调它——与 SKILL.md「可测试结构」口径一致
- **初始化可以不 Show 地验证**（`UserForms.Add` 加载并触发 `UserForm_Initialize`
  但不显示，不会挂起）：

  ```vba
  ' run_test 注入的代码：
  Public Function RunTest() As String
      Dim f As Object
      Set f = VBA.UserForms.Add("frmMy")      ' 触发 Initialize，窗体不显示
      RunTest = "items=" & f.cmbType.ListCount ' 断言初始化效果
      Unload f
  End Function
  ```

- **视觉与交互人工验收**：交付时让用户跑一次真窗体确认布局与动线

## 布局与视觉口径（交付前的自查清单）

窗体美学的硬规则——工具只能查几何（`lint_form`），对齐留白配色靠这套口径自查：

- **间距用网格**：所有外边距统一（建议 10-12px），控件间距统一（6-8px）；同组控件左对齐或顶对齐，不要手工目测错位
- **字号两档为限**：正文 9pt（Tahoms/Segoe UI），标题 11-12pt 加粗；更多层级说明信息架构该简化了
- **配色三色为限**：一个主题色（按钮/强调）+ 中性灰阶 + 一个警示色；颜色只用于表达状态，不装饰
- **主次分明**：主按钮（确定/执行）视觉突出且设 `Default=True`，取消/关闭设 `Cancel=True`；危险操作（删除/覆盖）与普通按钮拉开距离或换警示色
- **动线顺序**：Tab 键序 = 视觉阅读序（左→右、上→下）；输入框带有意义的标签，不靠占位文本
- **尺寸克制**：窗体宁小勿大，一屏放下；内容多就分组（Frame/MultiPage）或分步，不拉滚动条
