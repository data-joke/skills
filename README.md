# 数据分析 Skills 集合

按 skill 分目录存放,每个目录自包含(含 `SKILL.md` 指令 + 脚本 + 文档)。

## 目录

- [dataworks](dataworks/) — 阿里云 DataWorks 运维:新建/调整开发节点(调度/重跑/依赖/参数,可提交发布)、查实例状态与日志并分析失败原因、重跑/停止/置成功、按任务名或业务日期补数据并跟踪进度、按任务名解析 nodeId/fileId、查节点代码;并支持节点上下游血缘、实例聚合统计、业务流程、文件版本、表元数据/血缘、资源组、基线保障、数据质量(DQC,只读)。

## 下载与安装

### 方式一:手动点击下载(无需 git)

1. 进入目标 skill 目录,如 [dataworks](dataworks/),逐个点击文件下载;或
2. 在仓库首页点击 **Code → Download ZIP**,解压后取出对应 skill 目录(如 `dataworks/`)。
3. 把整个 skill 目录放到 `~/.claude/skills/<name>/`(Claude Code),即可被自动识别。

### 方式二:git clone(推荐,便于后续更新)

```bash
# 克隆整个仓库
git clone https://github.com/data-joke/skills.git

# 或只拉取单个 skill(稀疏检出,省流量)
git clone --depth 1 --filter=blob:none --sparse https://github.com/data-joke/skills.git
cd skills && git sparse-checkout set dataworks

# 把 dataworks/ 目录放到 ~/.claude/skills/dataworks/
```

### 配置凭证(重要)

复制模板并填写自己的 AccessKey:

```bash
cp dataworks/config.example.env dataworks/.env
# 编辑 .env,填入 DATAWORKS_ACCESS_KEY_ID / SECRET / REGION_ID
```

> ⚠️ `.env` 含密钥,**切勿提交或转发**;`.gitignore` 已忽略 `.env`。每人配置自己的 AccessKey。

## 使用

以 Claude Code 为例:把某个 skill 目录放到 `~/.claude/skills/<name>/`,即可被自动识别。每个 skill 的详细用法见其目录内 `README.md`。
