# 函数参考

> 由 SKILL.md 按需引用。需要确认某个函数的**确切签名/参数/返回值**时读本文件。
> 工作流与决策规则在 SKILL.md，本文件只做查表。

所有函数均支持 `attach=False`（默认独立实例）/ `attach=True`（复用用户 Excel，永不 Quit）。

## 创建与格式

| 函数 | 说明 |
|------|------|
| `create_addin(xlam_path)` | **从零创建空 .xlam**（0→1 起点），返回绝对路径，后续用同路径叠加 VBA/Ribbon |
| `get_file_format(file)` / `convert_to_xlsm(file)` | 判定格式 / xlsx→xlsm 转换（写操作会自动转，一般无需手动调用） |

## 环境与进程

| 函数 | 说明 |
|------|------|
| `check_environment(verbose=True)` | 返回环境诊断 dict |
| `enable_vba_trust()` | 写注册表启用 VBA 工程访问（HKCU） |
| `shutdown_excel(force_wait=5)` | 关闭工具缓存的隔离实例（Quit 后验证，存活则强杀；脚本结尾可显式调用） |
| `is_file_locked(path)` / `unload_addin(addin)` / `reload_addin(addin, path=None)` | 检测文件锁 / 在用户 Excel 卸载插件释放锁 / 重新加载（已装插件升级链路，默认 attach 用户实例）；返回含 `workbook_loaded` 复核字段——本 Office 构建实测 COM 安装/重载不即时装载工程（UI 对话框可、启动时自动加载可），装后建议重启 Excel 验收 |

## 源码同步

| 函数 | 说明 |
|------|------|
| `export_vba_source(file, out_dir=None)` | 导出全部组件 + manifest.json；返回 manifest |
| `import_vba_source(file, source_dir, sync=False)` | 导回；返回 {imported, updated, deleted, skipped} |

## 验证

| 函数 | 说明 |
|------|------|
| `build_check(file, compile=True, strict=False, timeout=60, check_binding=True)` | 静态 + 真编译；strict 附加 Option Explicit/重名检查；check_binding 拦截外部类型库的前期绑定（默认开） |
| `run_macro(file, name, args=None, timeout=120, save=False)` | 运行宏捕获错误 |
| `run_test(file, code, proc_name="RunTest", timeout=120, keep=False)` | 临时测试宏；keep=True 保留并保存 |
| `lint_form(file, form_name)` | 窗体几何检查 {findings, ok} |

## 备份

| 函数 | 说明 |
|------|------|
| `backup_file(file, reason)` / `list_backups(file)` / `restore_backup(file, backup=None)` | 手动备份/列出/恢复（恢复自动先备份当前）。⚠️ `restore_backup` 不带参恢复的是**最后一次写操作前**的状态，不是整批之前 |

## VBA 代码编辑

| 函数 | 说明 |
|------|------|
| `add_vba_module(file, name, code)` | 添加/整体替换模块（保留模块属性） |
| `add_vba_class(file, name, code)` | 添加/替换类模块 |
| `add_vba_callback(file, module, name, code)` | 追加过程（同名已存在则跳过） |
| `update_vba_callback(file, module, name, code)` | 替换单个过程（无则追加），返回是否替换 |
| `delete_vba_callback(file, module, name)` / `delete_vba_module(file, name)` | 删除过程/模块 |
| `replace_vba_module(file, name, code)` | 清空重写模块内容（原位，保属性） |
| `replace_vba_lines(file, module, start, count, code)` | 行级替换 |
| `insert_vba_code(file, module, after_proc, code)` | 在某过程后插入 |
| `read_vba_module(file, name)` / `list_procedures(file, name)` / `list_vba_modules(file)` / `get_procedure_info(file, module, proc)` | 读取类 |

## Ribbon（操作解包目录；.xlam/.xlsm 均可用）

| 函数 | 说明 |
|------|------|
| `init_custom_ui(dir, xml=None)` | **从零初始化** customUI 全套设施（幂等），未定制过的文件第一步必调 |
| `unpack_xlam(file, dir)` / `pack_xlam(dir, out, vba_source=None)` | 解包/打包；改过 VBA 后打包必须传 `vba_source` 防止代码回退 |
| `get_ribbon_xml(dir)` / `set_ribbon_xml(dir, xml)` | 读写 customUI.xml（set 对未初始化的包自动初始化） |
| `add_button_to_ribbon(dir, group_id, button_xml, after_button_id=None)` | 加按钮（支持自闭合分组、id/idQ 定位） |
| `add_group_to_ribbon(dir, group_xml, after_group_id=None, tab_id=None)` | 加分组；`tab_id` 支持 id/**idMso**/idQ，可加到内置标签页 |
| `register_icon(dir, icon_id, png_path)` / `unregister_icon(dir, icon_id)` | 注册/注销 PNG 图标（自动维护 rels + Content-Types） |
| `get_icon_rels(dir)` / `set_icon_rels(dir, rels_content)` | 读写 customUI.xml.rels（底层；`register_icon` 已封装，一般无需直接调用） |
| `generate_button_xml(...)` / `generate_group_xml(...)` | XML 生成 |
| `generate_vba_callback(name, control_param, doc)` / `generate_full_feature_module(...)` | VBA 模板生成 |
| `render_ribbon_preview(source, out_png, width=1000, tab_index=0)` | **把 customUI.xml 渲染成模拟功能区 PNG**（纯绘图不启 Excel）；source 可为解包目录或 .xlam。装进 Excel 前先看一眼 |

## 用户窗体

| 函数 | 说明 |
|------|------|
| `create_userform(file, name)` | 创建（默认 300×200） |
| `add_control(file, form, progid, name, properties, event_handler, container, page)` | 加控件，属性直接映射（Caption/Left/Top/Width/Height/List/...）；MultiPage 容器用 `page=` 指定页（索引/页名，见 `references/userform.md`），Frame 走 `container` 即可；`event_handler` 已废弃（MSForms 靠命名约定绑定） |
| `add_form_event_handler(file, form, event_name, code)` | 加事件处理 |
| `set_form_properties(file, form, properties)` | 设置窗体属性 |
| `generate_form_init_handler(form_name, form_properties)` | 生成初始化代码 |

## 图标

| 函数 | 说明 |
|------|------|
| `check_icon_api(quiet=False)` | **一站式预检**（生成前必调）：配置完整性 + URL 格式 + 端点连通（unreachable/auth_failed/not_compatible/ok），输出 ≤2 行 |
| `set_icon_config(...)` / `get_icon_config()` | 图像 API 配置（存 `~/.vba-dev/icon_config.json`，key 环境变量优先） |
| `generate_icon(prompt, out_png, size=1024)` | 调图像模型生成（b64/url 双兼容；未配置 fail-fast，运行期错误已翻译为中文诊断+修复指引） |
| `draw_icon(text, out_png, bg, fg, canvas=64)` | 离线兜底绘制（中英文均可） |
| `prepare_icon(src, out, size=32, rounded=True, white_to_alpha=False)` | 尺寸规整：居中裁切缩放 + 圆角 alpha + 可选白底转透明 |
| `add_icon_button(xlam_dir, group_id, button_id, label, on_action, icon_png, screentip="", supertip="", size="large", rounded=True, white_to_alpha=False, after_button_id=None)` | **一站式**：规整→注册→加按钮；AI 生成的图标必须传 `white_to_alpha=True` |
| `validate_imagemso(names)` | 进程内验证 imageMso 有效性（pywin32 进程外调用全部假失败，勿用） |
| `COMMON_IMAGEMSO` | 实测有效图标名清单 dict（按类别） |

## 用户偏好（跨会话）

| 函数 | 说明 |
|------|------|
| `get_preferences()` | 读 `~/.vba-dev/preferences.json`（不存在返回 `{}`，永不抛错） |
| `set_preference(key, value)` | 记住一条偏好（合并写入）；`value=None` 删除该键 |
| `clear_preferences()` | 清空全部偏好 |

约定键：`icon_type`(imageMso/draw/ai/user_png)、`icon_style`(自由文本)、`theme_color`("#1F4E92")、
`naming_style`(verb/noun/prefixed)、`default_group`、`default_tab`。
