# DataWorks Skill

让 AI 直接操作阿里云 **DataWorks(公共云,API 2020-05-18)** 的技能。通过一个自包含的本地 CLI(`dwcli.py`)封装官方 SDK,AI按需调用——避免把上百个 API 全量封成 MCP 导致的元数据 token 常驻消耗。

## 介绍

支持的能力:

- **新建开发节点**:可配置调度频率/时间(cron、周期、生效起止)、重跑属性(重跑模式、自动重跑次数/间隔)、依赖与调度参数;并可提交/发布上线。
- **实例运维**:查实例状态、查运行日志、**分析失败原因**、重跑 / 停止 / 置成功;`instance stat` 按业务日期输出状态分布 + 失败 Top 节点。
- **补数据**:按任务名或业务日期范围发起补数据,并跟踪补数据 DAG 进度。
- **元数据/代码**:查节点代码,节点/文件/项目列表。
- **节点血缘/影响面**:`node parents` / `node children` 查上/下游,重跑或补数前评估影响面。
- **业务流程**:`business list` / `get` / `files` 列出业务流程及其下作业。
- **文件版本**:`file versions` / `file version` 查版本历史与任意版本代码。
- **表元数据/表血缘**:`meta table`(表结构)、`meta lineage`(上下游表),表名用 `project.table`。
- **资源组**:`resource list` 列出调度/计算资源组。
- **基线保障**:`baseline list` / `baseline status` 查基线配置与当天各基线 SAFE/DANGER 状态。
- **数据质量 DQC(只读)**:`quality entity` / `rules` / `results` 查质量实体、规则与校验结果。
- **按任务名解析作业 ID**:说任务名即可解析出 nodeId/fileId 再操作(如“给 xxx 任务补数”)。

**适用场景**:DataWorks 日常数据开发与运维(早上查失败任务、看日志定位原因、重跑、补数、新建/调整节点等)。

**前置要求**:
- Python 3.8+(本机为 3.14)
- 一把具备 DataWorks 权限的阿里云 AccessKey(如授予 `AliyunDataWorksFullAccess` 或最小化策略)

## 安装教程

> 假设你收到的就是本文件夹 `dataworks/`。

**1. 放置技能目录**

- **Claude Code**:放到 `~/.claude/skills/dataworks/`(也可放项目 `.claude/skills/` 下;加载技能时会提示 Base directory,脚本按该目录动态定位)。
- **其他 Agent**:按其技能/规则机制加载 `SKILL.md`,命令中的 `<技能目录>` 一律替换为本目录实际路径。

**2. 安装依赖(SDK,装到全局 python3)**

```bash
cd skills/dataworks  # 切换到skill目录下
python3 -m pip install -r requirements.txt
# 或一键脚本(装依赖 + 生成 .env + 自检):
bash bootstrap.sh
```

> 💡 Windows(Git Bash)可能只有 `python` 无 `python3`(且 `WindowsApps/python3` 是 Microsoft Store 假入口,运行即失败)。bootstrap.sh 已自动探测回退到 `python`;手动安装时请用 `python -m pip install -r requirements.txt`。

**3. 配置你自己的凭证**

复制模板并填写:
  ```bash
  cp config.example.env .env
  # 编辑 .env,填入 DATAWORKS_ACCESS_KEY_ID / DATAWORKS_ACCESS_KEY_SECRET / DATAWORKS_REGION_ID
  ```
  也可用 `export` 写进 `~/.zshrc`:
  ```bash
  export DATAWORKS_ACCESS_KEY_ID=<你的AK>
  export DATAWORKS_ACCESS_KEY_SECRET=<你的SK>
  export DATAWORKS_REGION_ID=cn-shanghai      # 按实际区域
  # export DATAWORKS_PROJECT_ID=<工作空间数字ID>   # 可先用 project list 查,先留空也行
  ```

> ⚠️ `.env` 含密钥,切勿提交或转发给他人;每人配置自己的 AccessKey。

**4. 验证**

```bash
python3 skills/dataworks/scripts/dwcli.py doctor --check   # 自检 + 连通性，请使用完整路径
python3 skills/dataworks/scripts/dwcli.py project list     # 列出工作空间,拿 projectId
```

**5. 在 Claude Code 中使用（样例子）**

直接用自然语言描述需求即可触发,例如:
- “查一下昨天失败的 DataWorks 实例,分析下原因”
- “给 `order_ods` 这个任务补一下 7 月 18 号到 19 号的数据”
- “看下 `user_dwd` 节点的代码”

## 配置说明

| 变量 | 说明 | 默认 / 回退 |
|---|---|---|
| `DATAWORKS_ACCESS_KEY_ID` / `DATAWORKS_ACCESS_KEY_SECRET` | AccessKey | 回退 `ALIBABA_CLOUD_ACCESS_KEY_ID/SECRET` |
| `DATAWORKS_REGION_ID` | 区域 | 回退 `ALIBABA_CLOUD_REGION_ID`,默认 `cn-shanghai` |
| `DATAWORKS_ENDPOINT` | 覆盖 endpoint(VPC/金融云) | `dataworks.<region>.aliyuncs.com` |
| `DATAWORKS_PROJECT_ID` | 工作空间数字 ID | 可先留空,`project list` 查 |
| `DATAWORKS_PROJECT_ENV` | 环境 | `PROD` |

优先级:**命令行参数 > 进程环境变量 > skill 目录 `.env` > 默认值**。

> 💡 **写给 AI 助手 / 脚本调用**:Agent 通常以**非交互 shell** 执行命令,mac只读 `~/.zshenv` 而**不读 `~/.zshrc`**。若希望 AI 助手能直接调用,请把环境变量写入 `~/.zshenv`(对所有 zsh 生效);`~/.zshrc` 仅交互式终端有效。

## 文件结构

```
dataworks/
├── SKILL.md            # 给 Claude 的技能说明(工作流编排、命令文档)
├── scripts/
│   └── dwcli.py        # 自包含 CLI(封装 DataWorks SDK)
├── reference/          # 按需读取的参考文档(触发器见 SKILL.md)
│   ├── scheduling.md   # 调度配置速查(Cron/调度参数/节点类型编码)
│   └── pitfalls.md     # 常见坑与排查
├── requirements.txt    # SDK 依赖(锁定版本)
├── config.example.env  # 凭证模板 → 复制为 .env
├── bootstrap.sh        # 一键:装依赖 + 生成 .env + 自检
├── .gitignore          # 忽略 .env 与缓存
└── README.md           # 本文件
```

## 常见问题 / 排错

| 现象 | 处理 |
|---|---|
| 退出码 6 / 提示 SDK 未安装 | `python3 -m pip install -r requirements.txt` |
| 退出码 3(鉴权失败) | 检查 AccessKey;`doctor` 查看脱敏后的 AK 是否正确 |
| 退出码 5(无权限) | 为 AccessKey 授予 DataWorks 权限(如 `AliyunDataWorksFullAccess`) |
| 提示“缺少 projectId” | 先 `project list` 查数字 ID,再传 `--project-id` 或设 `DATAWORKS_PROJECT_ID` |
| 查不到数据 | 确认 `--env`(PROD/DEV)与业务日期;加 `--debug` 看实际请求 |
| 节点类型编码不对 | 用 DataWorks 控制台核对数字编码,`--type` 可直接传数字 |
| 查实例结果不完整/超时 | `instance list` 必须带 `--biz-date`(或状态/节点)限定范围,否则拉全量历史 |
| 多个工作空间易混 | 操作前用 `project list` 确认目标 `--project-id`,防跨空间误操作 |
| `baseline status` 报 bizdate 格式错误 | CLI 已内部把 `--biz-date` 转 `yyyy-MM-ddTHH:mm:ss+0800`(RFC822 时区);勿手动传字面 `Z` |
| `quality entity` 返回空 | 该工作空间未配置 DQC 质量实体,需先在「数据质量」模块为表建实体/规则 |
| 全量遍历命令限流(417 Throttling) | `instance stat` / `business files` / `meta` 为分页遍历,CLI 已自动退避重试 |

> 写操作(新建/提交/发布/重跑/停止/置成功/补数据)均需 `--yes` 确认;`--dry-run` 可先预览将执行的请求。详见 `SKILL.md`。
