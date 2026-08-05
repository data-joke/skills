# AI Skills 集合

我的 Claude Code 自定义 skills,按 skill 分目录存放,每个目录自包含(含 `SKILL.md` 指令 + 脚本 + 文档)。

## 目录

- [dataworks](dataworks/) — 阿里云 DataWorks 运维:新建/调整开发节点(调度/重跑/依赖/参数,可提交发布)、查实例状态与日志并分析失败原因、重跑/停止/置成功、按任务名或业务日期补数据并跟踪进度、按任务名解析 nodeId/fileId、查节点代码。

## 使用

以 Claude Code 为例:把某个 skill 目录放到 `~/.claude/skills/<name>/`,即可被自动识别。每个 skill 的详细用法见其目录内 `README.md`。
