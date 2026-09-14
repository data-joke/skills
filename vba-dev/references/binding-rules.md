# 引用与绑定规范（分发安全）

> 由 SKILL.md 按需引用。**写 VBA 代码前读一次**；`build_check()` 会默认拦截违规写法。

**.xlam 分发的头号坑是前期绑定（Early Binding）。** 插件一旦引用外部类型库
（Scripting / ADODB / VBScript_RegExp_55 / Outlook …），接收方没勾选那个引用，
打开即编译失败「**用户定义类型未定义**」——对方在插件里无从自救。

**规则：只要能用后期绑定（`CreateObject`）实现，一律后期绑定；禁止前期绑定。**

| 需求 | ❌ 前期绑定（禁止） | ✅ 后期绑定 |
|------|-------------------|-----------|
| 字典 | `Dim d As New Scripting.Dictionary` | `Dim d As Object: Set d = CreateObject("Scripting.Dictionary")` |
| 文件系统 | `Dim f As New Scripting.FileSystemObject` | `CreateObject("Scripting.FileSystemObject")` |
| 正则 | `Dim re As New VBScript_RegExp_55.RegExp` | `CreateObject("VBScript.RegExp")` ⚠️ ProgID 不带 `_55` |
| 数据库 | `Dim cn As New ADODB.Connection` | `CreateObject("ADODB.Connection")` |
| HTTP | `Dim x As New MSXML2.XMLHTTP60` | `CreateObject("MSXML2.XMLHTTP")` ⚠️ ProgID 不带 `60` |
| 命令行 | `Dim s As New WScript.Shell` | `CreateObject("WScript.Shell")` |
| 发邮件 | `Dim ol As New Outlook.Application` | `CreateObject("Outlook.Application")` |

**枚举/常量**：后期绑定取不到类型库常量，在模块顶部用 `Const` 本地重声明
（如 `Const adOpenStatic = 3`），**不要为此退回前期绑定**。

**可以放心用前期绑定的（宿主自带引用，接收方一定有）**：
`IRibbonControl` / `IRibbonUI`（Office 库，Excel 工程默认引用）、`MSForms.*`（含窗体的工程自动引用）、
`Excel.*` / `Office.*` / `stdole.*` / `VBA.*`、同工程内的自定义类模块。

**无法后期绑定的例外**（`build_check` 仍会报，需人工决定）：
`Dim WithEvents X As <外部类型>`（VBA 语法要求编译期类型）、`Implements <外部接口>`、类型库自定义类型（UDT）。

**豁免出口**：模块顶部加一行 `' @allow-early-binding`（并注明理由），该模块跳过检查；
或 `build_check(..., check_binding=False)` 整体关闭。

**自动检查**：`build_check()` 默认把外部类型库的前期绑定判为**错误（阻断构建）**，
错误项带 `"kind": "early_binding"`，错误信息直接给出对应的 `CreateObject` 写法。

## 白名单的判定口径（为什么不是无条件一刀切）

- **可后期绑定**（上表左列）：`Scripting` / `ADODB` / `ADOX` / `MSXML2` / `MSXML` / `WScript` /
  `Shell` / `VBScript_RegExp_55` / `Outlook` / `Word` / `PowerPoint` / `CDO` / `WinHttp` /
  `InternetExplorer` / `WIA` / `SAPI` —— 全部可通过 `CreateObject` 拿到，没有理由前期绑定。
- **宿主库**：不是用户手动勾选的引用，随 Office / 窗体自动携带，不构成分发风险。
- **真例外**：VBA 语法层面无法后期绑定，一刀切会把它们判成无法修复的死锁，
  所以留了 `' @allow-early-binding` 豁免口。
