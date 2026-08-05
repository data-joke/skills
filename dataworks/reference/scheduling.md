# 调度配置速查(`file create` 相关)

> 触发场景:新建/调整开发节点、配置调度频率、调度参数、依赖时读取。参数以 `scripts/dwcli.py file create --help` 为准,本文件只补语法知识。

## CronExpress 语法(DataWorks)

字段顺序:`秒 分 时 日 月 周`(可省略年)。

常用特殊字符:

| 字符 | 含义 | 示例 |
|---|---|---|
| `*` | 每个 | `*` 分 → 每分钟 |
| `?` | 无指定(仅日/周用) | 周字段 `?` = 不按周 |
| `-` | 范围 | `9-23` = 9 到 23 |
| `,` | 列表 | `1,15` = 1 和 15 |
| `/` | 步进 | `09-23/1` = 9 到 23 每小时 |

常用表达式:

| 意图 | 表达式 |
|---|---|
| 每天 00:05 | `00 05 00 * * ?` |
| 每天 09:05~23:05 每小时(共 15 次) | `00 05 09-23/1 * * ?` |
| 每小时整点 | `00 00 * * * ?` |
| 每周一 09:00 | `00 00 09 ? * 1` |
| 每月 1 号 00:00 | `00 00 00 1 * ?` |

## 调度参数

- **系统参数**:`$[yyyymmdd]`(当天)、`$[yyyymmdd-1]`(前一天)、`$[yyyymmdd-2]`;小时级 `$[yyyymmddhh24]`、`$[yyyymmddhh24-1/24]`(**上一小时**,常用)。`${bizdate}`、`${cyctime}` 为 DataWorks 内置系统参数。
- **代码内引用**:ODPS SQL / Shell 用 `${param}`;PyODPS 用 `args['param']`(如 `args['dt']`)。
- **小时分区约定**:本环境常见 `dt=yyyymmddhh` 小时分区,取数需按小时精确匹配(如 `dt='{dt}'`)。

## 节点类型编码(`--type`)

| 别名 | 数字编码 |
|---|---|
| odps-sql | 10 |
| odps-mr | 11 |
| odps-script | 24 |
| di / data-integration | 23 |
| shell | 4 |
| virtual | 99 |
| pyodps2 | 221 |
| pyodps / pyodps3 | 225 |

## 重跑模式(`--rerun-mode`)

| 值 | 含义 |
|---|---|
| ALL_ALLOWED | 允许任意重跑(默认常用) |
| FAILURE_ALLOWED | 仅失败可重跑 |
| ALL_DENIED | 禁止重跑 |

配合 `--auto-rerun-times <N>`(自动重跑次数)与 `--auto-rerun-interval <毫秒>`(间隔)。

## 依赖配置

- `--input <上游输出名>`:按上游输出表/输出名建立血缘依赖,逗号分隔多个。
- `--dep-type + --dep-nodes <node_id 列表>`:直接指定依赖的节点 ID。

## 生效区间

- `--start-effect` / `--end-effect`:调度生效起止(YYYY-MM-DD)。新建节点当天不生效,一般结束时间设 `2099-01-01`。
