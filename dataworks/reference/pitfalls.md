# 常见坑与排查

> 触发场景:实例状态异常、补数结果与预期不符、查不到数据、排错无思路时读取。

## 实例全是 NOT_RUN / 当天不跑

- **当天新创建/发布的节点,当天不触发自动调度**:实例虽已生成但处于 NOT_RUN,从**次日**首个调度点起生效。
  - 例:每天 09:05~23:05 的节点,当天下午创建,当天 15 个实例全部 NOT_RUN,属正常。
- NOT_RUN = 实例已生成、未到触发时间(等待中),不是故障。
- **需要当天立即生效** → 用 `complement run` 补数据触发,或对指定实例手动 `instance restart`。

## instance list 查不到 / 超时 / 503

- 不带 `--biz-date`(或 `--status`/`--node-id`/`--dag-id`)会**拉全量历史**,可能超时或报 `ServiceUnavailable`(503)。
- **先加 `--biz-date YYYY-MM-DD` 或 `--status FAILURE` 缩小范围**。
- 503 多为 DataWorks 服务端临时故障,稍后重试即可,与本 CLI 无关。

## 时间戳显示

- JSON 输出中 `bizdate`/`begin_running_time`/`finish_time` 为 **epoch 毫秒**(如 `1785859200000`);`--format table` 已自动转为 `YYYY-MM-DD HH:MM`。
- 需要精确到秒/特定时区时,以 JSON 原值为准。

## file_id 混用

- `file list` / `node resolve` 返回的 `file_id` 可直接用于 `file get --file-id`。
- `node get` 的 `FileId` 字段与 file_id **可能不同**(有时等于 node_id),**不可混用**。
- 取代码优先:`file get --node-id <id> --format text`。

## 多工作空间

- 同一 AccessKey 下可能有多个工作空间(如 `archive`、`lyy_gz`),多数命令需 `--project-id`。
- **操作前先确认 projectId 归属**,跨空间的重跑/置成功/补数据影响不可逆。
- 不确定时先 `project list` 与用户确认目标空间。

## 查失败实例的完整套路

1. `instance list --failed --biz-date <日期>`(或 `--task-name <任务名>`)找到失败实例。
2. `instance log --instance-id <id> --grep ERROR --lines 200 --format text` 取日志。
3. 分析根因:SQL 语法错误 / 资源不足(OOM)/ 上游未完成 / 权限不足 / 数据质量问题。
4. 修复后 `instance restart`;确认可忽略则 `instance set-success`。

## 补数据到小时级(重点)

小时任务补数最容易"补错分区",核心是 **bizdate 与调度日的映射**:

- **规律(实测)**:DataWorks 小时任务实例,业务日期 `D` 的实例,**实际调度时间 cyc 在 `D+1`**,而 `dt = cyc - 1 小时`。
  - 例:补 `--start-biz 2026-08-01` 后,实例 cyc=`2026-08-02 12:05`、`dt=2026080211`(**检查的是 08-02 11 点分区**,不是 08-01!)。日志里看 `SKYNET_BIZDATE`(业务日期)与 `SKYNET_CYCTIME`(调度时间)即可确认。
- **要补"某日 HH 点数据分区 dt=YYYYMMDDHH"** → 业务日期应传 **该日前一天**,时间范围覆盖调度 `(H+1):05`:
  - 例:补 `dt=2026080111`(08-01 11 点)→ 业务日期 `2026-07-31`、时间 `12:00:00~12:59:59`。
- **推荐直接让 CLI 换算**:`complement run --task-name <名> --data-date 2026-08-01 --hour 11 --yes`,CLI 自动算成业务日期 07-31 + 时间 12:00:00~12:59:59 并打印换算明细。
- **时间参数格式**:`--begin-time` / `--end-time` 需为 `HH:mm:ss`(如 `12:00:00`、`12:59:59`);`begin` 必须早于 `end`,否则报"时间区间为空"。
- **parallelism 是 bool 且必填**:RunCycleDagNodes 的 `Parallelism` 传数值/字符串均报 `InvalidParallelism`,CLI 已内置为 `True`,无需也不能手动传值。

## file update 与提交发布

- `file update` 只改**开发环境(DEV)**代码/调度配置;生产不受影响。
- 上线顺序:`file update` → `file submit` → `file deploy`。
- `file update` 不传 `--content`/`--content-file` 时**自动保留现有代码**,可单独改调度参数(如 `--para`),不会清空代码。
- 更新后可用 `file get --node-id <id> --format text --env PROD` 核对生产是否已生效。

## 其他

- 退出码:3 = 鉴权失败(检查 AccessKey)、5 = 无权限(需授权)、6 = SDK 未安装(`pip install -r requirements.txt`)。
- 凭证问题先跑 `doctor --check`。
- 不确定参数:`python3 scripts/dwcli.py <子命令> <操作> --help`。
