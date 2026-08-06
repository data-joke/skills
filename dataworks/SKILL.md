---
name: dataworks
description: 操作阿里云 DataWorks(公共云,API 2020-05-18)的本地 CLI 技能。能力:新建/调整开发节点(调度频率、调度参数、依赖、重跑策略,可提交/发布)、查实例状态与日志并分析失败原因、重跑/停止/置成功、按任务名或业务日期补数据并跟踪进度、按任务名解析 nodeId/fileId、查节点代码与项目/节点/文件列表。当用户提到 DataWorks、新建节点/任务、调度实例、失败重跑、补数据、查节点代码、按任务名找作业 ID 时使用。
allowed-tools:
  - Bash
  - Read
---

# DataWorks 运维 Skill

## 定位与调用

通过技能目录 `scripts/` 下的 `dwcli.py` 操作 DataWorks;**技能目录 = 加载本文件时提示的 `Base directory`**。调用统一用完整前缀(Bash 不保留别名/状态):

```bash
cd <技能目录> && python3 scripts/dwcli.py <子命令> [参数]
```

若不确定技能目录:`find ~/.claude -name dwcli.py 2>/dev/null`。

> stdout=结果,stderr=诊断;命令带 `2>&1` 混流时,结果以 stdout 为准。

## 帮助与参数确认(重要)

**不确定参数时,先查 `--help`,不要凭记忆猜测。** 本文件与 README 的命令说明可能滞后于 CLI,`--help` 是唯一权威来源:

```bash
python3 scripts/dwcli.py <子命令> --help          # 子命令级,如 file / instance
python3 scripts/dwcli.py <子命令> <操作> --help   # 操作级,如 file create --help
```

取值仍不确定时,用 `--dry-run` 预览实际请求。

## 命令速查(意图 → 命令)

| 用户意图 | 命令 | 备注 |
|---|---|---|
| 找/解析某任务 | `node resolve --name <任务名>` | `node list --name` 为**精确匹配**;名称不完整时用 `file list --keyword 关键字`(模糊包含) |
| 查作业所在目录 | `file folder (--name <任务名> \| --file-id <id> \| --node-id <id>)` | 输出完整 FolderPath,如 `Business Flow/.../商创/python脚本`;`file list` 也带 file_folder_id/path |
| 查上/下游节点 | `node parents \| node children (--node-id <id> \| --task-name <名>)` | 重跑/补数前评估影响面 |
| 查节点代码 | `file get (--node-id <id> \| --file-id <id>) --format text` | `file list --keyword` 按关键字找 file_id |
| 文件版本历史 | `file versions --file-id <id>`;取某版代码 `file version --file-id <id> --file-version <N> --format text` | 排查改动何时上线 |
| 新建开发节点 | `file create ...`(见下) | 写操作需 `--yes` |
| 提交/发布节点 | `file submit` / `file deploy` | 写操作需 `--yes` |
| 查实例/状态 | `instance list` | **必须限定范围**,见约束 |
| 实例按天统计 | `instance stat --biz-date <日期>` | 状态分布 + 失败 Top 节点,早上看失败任务 |
| 查日志/找失败原因 | `instance log --instance-id <id> [--grep ERROR]` | 结合 `instance list --failed` |
| 重跑/停止/置成功 | `instance restart/stop/set-success` | 写操作需 `--yes` |
| 补数据 | `complement run --task-name <名> --start-biz <起> [--end-biz <止>] --yes` | 返回 `dag_id`,`complement status` 跟踪 |
| 业务流程 | `business list` / `business get --business-id <id>` / `business files --business-id <id>` | files 全量遍历匹配,较慢 |
| 表结构/表血缘 | `meta table --table <project.table>` / `meta lineage --table <project.table> [--direction up\|down\|all]` | 数据地图;表名用 `lyy_gz.xxx` |
| 资源组 | `resource list` | 合并调度+计算类型 |
| 基线保障 | `baseline list` / `baseline status --biz-date <日期>` | status 输出当天各基线 SAFE/DANGER |
| 数据质量 DQC | `quality entity --table <project.table>` / `rules --entity-id` / `results` | 只读;需先配置 DQC 实体 |
| 看有哪些工作空间 | `project list` | 多空间操作前先确认目标 |

### file create — 调度配置项(写,需 `--yes`)

```
--cron <CronExpress> --cycle-type DAY|NOT_DAY     # 调度频率/周期
--start-effect <起> --end-effect <止>              # 生效区间
--para <调度参数,如 bizdate=$[yyyymmdd-1]>        # 日期参数,代码里 args['xx']/${xx} 引用
--input <上游输出名> 或 --dep-type+--dep-nodes     # 依赖
--rerun-mode ALL_ALLOWED|FAILURE_ALLOWED|ALL_DENIED --auto-rerun-times <N> --auto-rerun-interval <ms>
--folder <目录> --owner <X> --resource-group <RG>  # 归属
```
`--type` 别名:`odps-sql`/`odps-mr`/`odps-script`/`di`/`shell`/`virtual`/`pyodps2`/`pyodps3`(或直接传数字编码)。详见 `file create --help`。

## 关键约束(务必遵守)

1. **`instance list` 必须限定范围**:加 `--biz-date <YYYY-MM-DD>`(或 `--status`/`--node-id`/`--dag-id`),否则拉全量历史且可能超时/失败。
2. **多工作空间防误操作**:操作前确认 `--project-id` 归属;跨空间的重跑/置成功/补数据不可逆。
3. **取代码 ID**:优先用 `node resolve`/`file list` 返回的 ID;`node get` 的 `FileId` 与 `file list` 的 file_id 可能不同,不可混用。
4. **全量遍历命令**:`instance stat`、`business files` 会分页遍历全量数据,较慢且可能触发限流(417 Throttling),已内置 0.2-0.3s/页退避;`instance stat` 仍必须带 `--biz-date`。

## 写操作安全约定(重要)

`file create/submit/deploy`、`instance restart/stop/set-success`、`complement run` 均为写操作。
**执行前必须先向用户复述将要执行的操作(对象、范围、影响),获得确认后才加 `--yes`**;可用 `--dry-run` 预览。补数据、重跑、置成功、发布尤其谨慎。

## 排错要点

| 现象 | 处理 |
|---|---|
| 鉴权失败(退出码 3) | 凭证用 `DATAWORKS_ACCESS_KEY_ID/SECRET`(或 `ALIBABA_CLOUD_ACCESS_KEY_*`)或技能目录 `.env`;`doctor` 自检 |
| 查不到数据 | 确认 `--env`(PROD/DEV)与业务日期;`--debug` 看实际请求 |
| `baseline status` 报 bizdate pattern 错误 | CLI 已内部把 `--biz-date YYYY-MM-DD` 转 `yyyy-MM-ddTHH:mm:ss+0800`(RFC822 时区,网关正则要求);勿手动传字面 `Z` |
| `quality entity` 返回空 | 该工作空间未配置 DQC 质量实体,需先在 DataWorks「数据质量」模块为表建实体/规则 |
| 限流 417 Throttling.User | 连续快速调用触发,稍后重试;全量遍历命令已内置退避 |
| 结果与预期不符 | 先 `instance list --failed` + `instance log` 分析根因,再决定是否重跑/置成功 |

## 参考文档(按需读取)

- **新建/调整节点、配置调度频率/调度参数/依赖时** → 先 `Read reference/scheduling.md`(Cron 语法、调度参数、节点类型编码、重跑模式)。
- **实例状态 / 补数结果与预期不符,或排错无思路时** → `Read reference/pitfalls.md`(常见坑与排查)。
- 位于技能目录 `reference/` 下;仅相关场景才读,不要每次调用都加载。
