# VBA 高频坑（AI 生成代码约束清单）

> 由 SKILL.md 按需引用。**写 VBA 代码前读一次**；`build_check` 报编译错误、或 `run_test`
> 行为怪异时也回来对照。收录的是 AI 生成 VBA 的高频错误——多数**能通过编译**，
> 只在特定环境（64 位 Office / 插件场景 / 大数据量 / 别人机器）爆炸，所以工具拦不住，写时自查。

## 1. 64 位 Office：`Declare` 必须 `PtrSafe`（编译失败第一来源）

新机器几乎都是 64 位 Office，而 AI 训练语料里大量 32 位示例。缺 `PtrSafe` 在 64 位上**直接编译失败**：

```vba
' ❌ 32 位写法（64 位 Office 编译失败）
Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ...) As Long

' ✅ 64 位安全写法
Declare PtrSafe Function SetTimer Lib "user32" (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr, ...) As Long
```

- 规则：`Declare` 一律加 `PtrSafe`；**句柄/指针/DC/HWND 参数用 `LongPtr`**（32/64 位自适应），
  只有确知是纯数值（非指针）才用 `Long`
- 高频 API：`SetTimer`/`KillTimer`、`FindWindow`、`SHGetFolderPath`、`URLDownloadToFile`、`CopyFile`、`SendMessage`
- 兼容全部位宽时用条件编译：`#If VBA7 Then ... #Else ... #End If`（VBA7+ 有 `LongPtr`）

## 2. 插件场景：改了界面状态就必须在错误路径上恢复

插件代码运行在**用户正开着的工作簿**里——出错后不恢复 `ScreenUpdating`/`Calculation`，
用户的 Excel 界面就永久残废（不是一次性宏，关掉重开还在）。恢复代码要放在
错误处理标签上，不是过程末尾：

```vba
Public Sub DoWork()
    Dim prevCalc As XlCalculation
    prevCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo Cleanup
    ' ... 实际工作 ...
Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = prevCalc
    If Err.Number <> 0 Then MsgBox "执行失败: " & Err.Description
End Sub
```

同理：自己 `Set` 的全局/模块级对象变量，用完置 `Nothing`；改过 `Application.EnableEvents`
的必须配对还原（漏还原 = 用户从此点不到任何工作表事件）。

## 3. 禁用 `Select` / `Activate`（宏录制器风格）

AI 很爱生成录制器风格代码；插件场景下慢、闪屏、且脆弱（激活失败/改错对象）：

```vba
' ❌ 录制器风格
Range("A1").Select
Selection.Copy
Range("B1").Select
ActiveSheet.Paste

' ✅ 直接操作对象
Range("B1").Value = Range("A1").Value          ' 传值
rngSrc.Copy Destination:=rngDst                ' 或带目标的 Copy
```

**剪贴板不是数据通道**：能赋值/传参就不要 Copy+Paste——那会清掉用户剪贴板里的内容（插件大忌）。

## 4. `On Error Resume Next` 必须收窄到三行

全局铺开 `Resume Next` 会吞掉后续所有错误，bug 不可诊断。只包"预期可能失败"的那一句，立即检查并关闭：

```vba
On Error Resume Next
Set ws = ThisWorkbook.Worksheets("Config")   ' 可能不存在
On Error GoTo 0
If ws Is Nothing Then MsgBox "缺少 Config 工作表": Exit Sub
```

## 5. 插件语境：`ActiveWorkbook` vs `ThisWorkbook`

`.xlam` 的 `ThisWorkbook` 是**插件自己**（`IsAddin=True`，没有可见工作表）——
在插件代码里写 `ThisWorkbook.Worksheets(1)` 会**运行时下标越界**。

| 语境 | 用 | 原因 |
|------|-----|------|
| 操作用户当前数据 | `ActiveWorkbook` / `ActiveSheet` / `Selection` | 指向用户正在看的工作簿 |
| 读插件自带配置 | `ThisWorkbook` | 固定指向插件文件，不受用户切换工作簿影响 |

操作用户数据前加守卫，别假设用户选中了什么：

```vba
If TypeName(Selection) <> "Range" Then MsgBox "请先选中单元格区域": Exit Sub
```

## 6. `Integer` → `Long`；循环拼字符串 → `Join`

- VBA 的 `Integer` 是 16 位（最大 32767）——行数/计数器用它必然溢出。**计数一律 `Long`**。
- 循环内 `s = s & x` 是 O(n²)（每次全量复制）；大文本改数组 + `Join`：

```vba
' ❌ 大数据量下越来越慢
For i = 1 To n: s = s & cells(i) & vbCrLf: Next

' ✅ 等价于 StringBuilder
Dim arr() As String: ReDim arr(1 To n)
For i = 1 To n: arr(i) = cells(i): Next
s = Join(arr, vbCrLf)
```

## 7. 分发到别人机器的 Locale 陷阱

- **公式**：`.Formula` 固定 en-US 语法（逗号分隔），`.FormulaLocal` 随用户区域变——
  分发场景一律 `.Formula`
- **日期**：别用 `CDate("2026-01-02")`（结果随区域变）；构造日期用 `DateSerial(2026, 1, 2)`
- **数值转文本**：`CStr(1.5)` 的小数点随区域；拼 SQL/文件内容用 `Format$(v, "0.##")` 明确格式

## 8. 可测试结构（配合 `run_test`）

入口回调只做「取状态 → 调核心 → 报告」，核心逻辑写成**纯函数**，`run_test` 注入的
临时宏才能直接验证：

```vba
Sub CleanSelection_Click(control As IRibbonControl)   ' 入口：薄壳
    If TypeName(Selection) <> "Range" Then Exit Sub
    CleanRange Selection
End Sub

Function CleanText(s As String) As String             ' 核心：纯函数，可测
    CleanText = Trim$(s)
End Function
```

测试宏用 Function 返回结果字符串（如 `"Case1=ok;Case2=5"`），**别用 MsgBox 报告断言结果**——
看门狗会点掉它但拿不到内容。

## 9. Range 批量读写必须数组化（首要性能规则）

逐格 `.Value` 读写是 VBA 最常见的性能灾难（每格一次 COM 往返，1 万格肉眼可见地卡）：

```vba
' ❌ 逐格读
For i = 1 To n: For j = 1 To m: arr(i, j) = cells(i, j).Value: Next: Next

' ✅ 一次交换
Dim v As Variant
v = rng.Value                  ' 读：整个区域 → 二维数组（1 基）
' ... 内存中处理 ...
rng.Value = v                  ' 写：数组 → 整个区域
```

要点：数组化后循环在纯内存里跑，速度差 2-3 个数量级；区域只有 1 格时 `v` 不是数组（`IsArray` 判断）；配合 §2 的 ScreenUpdating/Calculation 关闭使用。逐格操作仅保留给**必须逐格改格式/注释**的场景。

## 检查落点

| 坑 | `build_check` 能否拦截 |
|----|----------------------|
| 缺 `PtrSafe` | ✅ 64 位机器上编译失败会报出（对照本文 §1 修） |
| 前期绑定（分发编译失败） | ✅ 默认拦截（见 `binding-rules.md`） |
| 界面状态不恢复 / `Select` 模式 / `Resume Next` 滥用 / `ThisWorkbook` 误用 / Locale | ❌ 静态检查不覆盖——写时按本文自查 |
