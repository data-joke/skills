# vba-dev — AI 编程助手的 Excel/VBA 插件开发技能

让 AI 全流程自动化开发 Excel 插件：从 0 创建 `.xlam`、编写 VBA、定制功能区（Ribbon）、生成图标，到**真实编译验证和运行测试**——全部闭环完成，无需人工在 Excel 里点一下。

> 一个 AI 编程助手技能（Agent Skills 格式，主流 AI 编程工具通用；配置存于工具无关的 `~/.vba-dev/`）。安装后对 AI 说"帮我做一个 XX 插件"即可。

## AI 安装提示词

安装最简单的方式：把下面这段提示词整段复制、发给你的 AI 编程助手即可。

```text
请帮我安装 vba-dev 技能（Excel/VBA 插件自动化开发技能，仅限 Windows + 桌面版 Excel）：
1. 环境检查：本技能仅限 Windows（依赖 Windows 独有的 Excel COM 接口，macOS/Linux 无法使用）。
   若当前系统不是 Windows，请直接停止安装并告知我
2. 浅克隆仓库到临时目录：git clone --depth 1 https://github.com/data-joke/skills.git
3. 把其中的 vba-dev 子目录完整复制到你的技能（skills）目录下，目录名保持 vba-dev
4. 安装 Python 依赖：pip install pywin32 Pillow（如需 AI 生成图标，另装 requests）
5. 验证：运行 python <技能目录>/vba-dev/scripts/test_xlam_toolkit.py，
   89 项测试全部通过即安装成功（纯静态测试，不会启动 Excel；其中 7 个 Excel
   端到端用例默认跳过，设 VBA_DEV_EXCEL_TESTS=1 可启用）
6. 清理临时目录，然后简要告诉我 vba-dev 能做什么、怎么用
```

## 核心能力

| 能力 | 说明 |
|------|------|
| 🧩 **0→1 创建插件** | `create_addin()` 从空文件创建 `.xlam`，VBA/Ribbon/图标逐层搭建 |
| ✍️ **VBA 编辑** | 模块/类/窗体/过程级增删改查；同名模块原位替换保留属性 |
| 🔄 **源码同步** | `export/import_vba_source()`：VBA 工程 ↔ `.bas/.cls/.frm` 文件双向同步——AI 直接编辑源码文件，改完一键导回 |
| ✅ **验证闭环** | `build_check()` 静态结构检查 + **引用绑定检查** + **真实 VBE 编译**（错误文本+出错行号）；`run_test()` 注入临时测试宏运行（文件零修改），错误/弹窗文本自动捕获 |
| 🎀 **Ribbon 定制** | 从零初始化 customUI（全新插件也能建）、按钮/分组/标签页、`idMso` 内置页签、自闭合分组自动展开 |
| 👁️ **Ribbon 预览** | `render_ribbon_preview()` 把 customUI 渲染成模拟功能区 PNG——功能区装进 Excel 前完全不可见，先看一眼再交付 |
| 🎨 **图标管线** | AI 图像生成（OpenAI 兼容 + MiniMax 自适应）→ 尺寸规整（1024→32×32 圆角透明）→ 一键嵌入；另有零依赖的离线字符图标兜底；内置图标名清单经**进程内实测验证**（附校验函数） |
| 🪟 **用户窗体** | 程序化创建控件 + `lint_form()` 几何检查（重叠/越界/负坐标） |
| 🛡️ **安全机制** | 所有自动化在**隔离 Excel 实例**中进行（绝不动用户打开的 Excel）；每次写操作自动备份、可回滚；弹窗看门狗自动读关模态框、卡死自动强杀 |
| 🔧 **环境自检** | `check_environment()` / `enable_vba_trust()` 一键检测/修复 VBA 信任设置 |
| 🔖 **用户偏好** | `get_preferences()` / `set_preference()` 跨会话记住用户的风格选择（图标类型/主题色/命名风格），下次不必重问 |

## 快速开始

对 AI 说：

> 帮我做一个插件：清理选区文本首尾空格、文本数字转数值，各一个按钮

AI 会就**图标 / 命名 / 分组位置**一次性给出选项和推荐（单轮征询）征询确认，然后走标准链路：

```python
from xlam_toolkit import (
    create_addin, add_vba_module, build_check, run_test,   # 0→1 + 验证
    unpack_xlam, init_custom_ui, add_group_to_ribbon,       # Ribbon
    add_button_to_ribbon, register_icon, pack_xlam,         # 按钮/图标/打包
    render_ribbon_preview,                                  # 功能区预览
)

create_addin("MyTools.xlam")                    # 1. 从零创建插件
add_vba_module("MyTools.xlam", "modUtils", ...) # 2. 写 VBA（自动备份；全程用同一个 .xlam 路径）
build_check("MyTools.xlam")                     # 3. 静态检查 + 引用绑定检查 + 真编译
run_test("MyTools.xlam", test_code)             # 4. 冒烟测试（文件不变）

# 5. Ribbon + 图标（支持 AI 生成/离线绘制/内置图标/用户图片）
unpack_xlam("MyTools.xlam", "pkg/")
init_custom_ui("pkg/")                          #    从零初始化功能区
add_button_to_ribbon("pkg/", "grp", btn_xml)    #    加按钮
render_ribbon_preview("pkg/", "preview.png")    #    先看一眼布局（功能区装进 Excel 前唯一可见的验收材料）
pack_xlam("pkg/", "MyTools_new.xlam", vba_source="MyTools.xlam")  # 6. 打包（输出必须换新路径）
backup_file("MyTools.xlam", reason="pack-replace")                #    留底（os.replace 不自带备份）
os.replace("MyTools_new.xlam", "MyTools.xlam")                    #    覆盖回原名，后续操作一律用原路径
```

### 图标生成的三种方式

```python
check_icon_api()                                # 预检 API（缺配置/URL不通/key无效 → 精准诊断）
src = generate_icon("扁平风格扫帚图标，透明背景", "icon_broom.png")  # ① AI 生成（out_png 必填）
src = draw_icon("清", "icon_clean.png", bg=(31, 78, 146))          # ② 离线兜底：字符图标（out_png 必填）
# ③ 内置图标：COMMON_IMAGEMSO（实测有效清单）或 validate_imagemso() 校验
add_icon_button(dir, group, id, label, on_action, icon_png=src)  # 规整+注册+加按钮一步到位
```

图像 API 配置：`set_icon_config(base_url=..., model=..., api_key=...)`。**配置存于 `~/.vba-dev/icon_config.json`——在 skill 目录之外，发布/分享本技能永远不会泄露密钥**（key 也可用环境变量 `ICON_API_KEY` 完全不落盘）。

## 示例作品

- [examples/Excel游戏厅.xlsm](examples/Excel游戏厅.xlsm) —— 基于 vba-dev 全流程开发的示例宏工作簿，可直接下载体验；AI 想参考其实现，用 `export_vba_source` 导出源码研读（二进制工作簿本身不可直接读）。

## 系统要求

- Windows + 桌面版 Excel（2016+ 实测通过；需在信任中心开启"信任对 VBA 项目对象模型的访问"，技能可自动检测并引导开启）

- Python 3.10+，`pywin32` 必需，`Pillow`（图标）、`requests`（AI 图标）推荐
- 已在中文 Office 16 + GBK 控制台环境下全面实测（真实插件 0→1 构建 + 真实 API 联调）
- 随附回归测试 `scripts/test_xlam_toolkit.py`（89 项；其中 7 个 Excel 端到端用例默认跳过，设 `VBA_DEV_EXCEL_TESTS=1` 启用——发布或改工具代码后建议连端到端一起跑）

> ⚠️ **macOS 用户注意**：本技能的自动化链路依赖 Windows 独有的 Excel COM 接口（`pywin32`），macOS 版 Excel 不提供该接口，且 Mac 版 Excel 的 VBA 适配不全（如不支持 ActiveX、窗体/功能区能力受限）。因此在 macOS 上本技能无法正常运行；技能生成的 `.xlsm` / `.xlam` 在 Mac Excel 中打开时，部分功能同样可能不可用。请优先在 Windows 环境使用。

## 目录结构

```
vba-dev/
├── SKILL.md                    # 技能主文档：工作流、执行契约、引导规范、排错速查
├── README.md                   # 本文件
├── LICENSE                     # 开源协议（署名转载 · 商用需授权）
├── .gitignore
├── references/                 # 查阅型文档（按需读取，不常驻上下文）
│   ├── function-reference.md   #   函数签名/参数/返回值 速查
│   ├── ribbon.md               #   customUI.xml 写法 · 回调模式 · imageMso 清单 · 图标管线 · 预览
│   ├── binding-rules.md        #   后期绑定禁令、ProgID 对照、白名单口径
│   ├── vba-pitfalls.md         #   AI 生成 VBA 高频坑（PtrSafe/64 位、插件语境、状态恢复等）
│   ├── userform.md             #   窗体控件 ProgID 与事件
│   ├── interaction.md          #   交互引导细则：各决策点的选项模板
│   └── troubleshooting.md      #   排错速查表（完整版）
├── scripts/
│   ├── xlam_toolkit.py         # 全部能力实现（纯 Python，单文件，无框架依赖）
│   ├── test_xlam_toolkit.py    # 回归测试（默认纯静态不启 Excel，改 toolkit 后跑它自验）
│   └── icon_config.example.json# 图像 API 配置字段说明（不含敏感信息）
└── examples/
    └── Excel游戏厅.xlsm        # 基于 vba-dev 开发的示例作品
```

## 开发自检

改动 `scripts/xlam_toolkit.py` 后，跑回归测试确认没有碰坏既有行为：

```bash
python scripts/test_xlam_toolkit.py         # 66 项，纯静态，不启动 Excel
python -m unittest discover -s scripts -v   # 或用 unittest discover
```

覆盖续行合并、前期绑定检测、块/结构检查、格式判定、API 签名、功能区预览渲染、偏好读写、配置迁移。
另有 5 个端到端用例默认跳过——设 `VBA_DEV_EXCEL_TESTS=1` 启用（会启动隔离 Excel 实例，
不会触碰你自己打开的 Excel）。

## 设计亮点

- **AI 编辑源码而非调函数**：批量改动走「导出 → 直接编辑 .bas 文件 → 导回」，契合 AI 原生能力
- **失败原子性**：写操作中途异常不保存，文件保持原状；随时 `restore_backup()` 回滚
- **防假通过**：编译命令在自动化下的三个静默失效点（已编译状态/隐藏 VBE/宏禁用打开）均已实测攻克——通过注入进程内探针宏触发真实编译并捕获错误弹窗
- **友好引导**：多方案决策点（图标/命名/分组等）主动征询并给推荐；API 未配置时错误信息自包含（缺什么+配置示例+兜底出路），减少无效重试

## 人工安装

不依赖 AI 的手动步骤（Windows + 桌面版 Excel）：

1. 安装 Python 依赖：`pip install pywin32 Pillow`（如需 AI 生成图标，另装 `requests`）
2. 下载本仓库，把 `vba-dev/` 目录完整复制到你所用 AI 编程助手的技能（skills）目录下
3. 验证：`python <技能目录>/vba-dev/scripts/test_xlam_toolkit.py`，66 项全部通过即安装成功

首次使用时 AI 会自动运行环境自检；若 VBA 信任未开启，`enable_vba_trust()` 一键修复。

## License

开源使用（个人学习/研究/非商业用途自由使用与修改），条款如下：

- **转载/再分发**须标明出处（作者 + 仓库链接）
- **商业使用**须事先获得作者书面同意

完整条款见 [LICENSE](LICENSE)。
