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

## 其他

- 退出码:3 = 鉴权失败(检查 AccessKey)、5 = 无权限(需授权)、6 = SDK 未安装(`pip install -r requirements.txt`)。
- 凭证问题先跑 `doctor --check`。
- 不确定参数:`python3 scripts/dwcli.py <子命令> <操作> --help`。
