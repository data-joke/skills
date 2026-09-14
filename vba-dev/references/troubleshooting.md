# 排错速查表（完整版）

主文档只保留高频 5 条；此表为全量。按症状定位，右列是处理动作。

| 症状 | 原因 | 处理 |
|------|------|------|
| 按钮图标不显示 | imageMso 名无效（如 Replace/Find/Print） | 用 `validate_imagemso()` 验证或换 `COMMON_IMAGEMSO` 名单内图标（见 `ribbon.md`） |
| 自定义图标显示异常/按钮变形 | 图片尺寸过大（模型出 1024px 直接嵌入） | 用 `add_icon_button()`（自动 `prepare_icon` 规整到 32×32） |
| 图标白底方丑 | 模型输出实底白背景（实测提示词「透明背景」基本被无视） | 嵌入时传 `white_to_alpha=True`（`add_icon_button` / `prepare_icon` 均支持）；这是唯一可靠兜底 |
| `无法访问 Visual Basic 工程` / 找不到编译菜单项 | 信任中心未开启 AccessVBOM | `enable_vba_trust()` 后**重启所有 Excel**；或 `check_environment()` 确认 |
| `宏被禁用`（运行时报） | 文件受保护视图（网络下载） | 文件右键→属性→解除锁定；run 系列已自动用 Low 安全级 |
| `excel.VBE` 报 0x800A03EC | 新实例 VBE 未就绪的瞬时失败 | 已内置 3 次重试；仍失败则 `shutdown_excel()` 后重试 |
| 宏运行卡死/弹窗不断 | 代码死循环或弹窗风暴 | 看门狗自动关闭弹窗；超时/风暴自动强杀并报告，检查 `r["dialogs"]` |
| 文件被锁定（Open 失败） | 其他 Excel 实例占用 | 已装插件被用户 Excel 加载 → `is_file_locked()` 确认后 `unload_addin()` 释放锁（改完 `reload_addin()`）；工具自身残留实例 → `shutdown_excel()`；用户打开了同名工作簿 → 请用户关闭，或 `attach=True`（attach 下运行宏不做超时强杀，卡死需人工结束） |
| 控制台中文/符号乱码 | GBK 控制台 | 运行脚本前 `set PYTHONIOENCODING=utf-8`；工具已内置 replace 兜底 |
| 读 xlsx 报"不含 VBA" | xlsx 无 VBA 工程 | 属预期；写操作会自动转 xlsm |
| 别人装了插件报「用户定义类型未定义」/「找不到工程或库」 | 前期绑定：对方未勾选该类型库引用 | 改后期绑定 `CreateObject`（见 `binding-rules.md`）；`build_check()` 已默认拦截此类写法 |
| 编译过了但运行行为怪 | 编译只查语法/结构 | 用 `run_test` 覆盖运行逻辑；关键路径都写冒烟测试 |
| 重载/安装后功能区没出现，`Workbooks` 里也没有该工程 | 本 Office 构建 COM 安装不即时装载工程 | 以 `reload_addin` 返回的 `workbook_loaded` 为准；让用户重启 Excel 或在加载项对话框手动勾选验收 |
| `set_form_properties` 报「发生意外」 | 会话内未实例化窗体 Designer | 已修复（内部先触碰 Designer）；自写 COM 时先访问一次 `comp.Designer` |
| MultiPage 窗体控件放不进/落错页 | MultiPage 不暴露 `Controls`，必须走 `Pages(i)` | `add_control(..., container="mpgX", page=页索引或页名)`；详见 `userform.md` |
| lint_form 报大量「控件重叠」warning | `designer.Controls` 为扁平集合，嵌套子控件跨容器误比 | 已修复（按父容器分组）；旧版产物均为 warning 级，可忽略 |
| 导出的 .bas/.cls 中文乱码 | 旧版工具导出为 GBK | 2026-09 起导出/导入统一 UTF-8（导入自动兼容旧 GBK 文件）；`.frm` 仍为二进制不要文本编辑 |
| 改坏了文件 | — | `list_backups()` + `restore_backup()` 回滚 |
