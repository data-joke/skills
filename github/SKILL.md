---
name: github
description: 通过 gh CLI 操作 GitHub:创建/查看/列表/评论/关闭 Issue(含标签、指派),PR 全流程(建分支、提交、创建、checkout、diff、CI checks、review、合并),Actions 运维(触发 workflow、查运行状态与失败日志、失败重跑),仓库/Release/Gist 管理与代码/仓库搜索。当用户提到 GitHub、gh、issue、PR/拉取请求、合并、merge、Actions、workflow、流水线失败/重跑、release/发版、gist、fork、clone 仓库,或说"建个 PR""看下这个 PR""合并""重跑失败的 job""发布 v1.0"等时使用。深度多 agent PR 代码评审请改用 /review。
allowed-tools:
  - Bash
  - Read
---

# GitHub 操作 Skill

## 1. 定位

本 skill 把 `gh` CLI(2.97.0)文档化为你的操作手册——**直接调 `gh`,不经过任何封装层**。当前已用 OAuth 认证登录账号 `data-joke`(token 存 macOS 钥匙串,git 协议 https),scopes:`gist, read:org, repo, workflow`(无 `delete_repo`/`admin:*` 等管理员权限)。

分工:**深度、多 agent 的 PR 代码评审用内置 `/review`**;本 skill 负责 gh 侧一切操作(取 diff、看 CI、提交评审结论、合并、Actions 排障、issue/release/gist 管理等)。

## 2. 前置自检(每个任务第一步)

每个 github 任务的**首条命令**必须先跑自检:

```bash
bash ~/.claude/skills/github/scripts/preflight.sh                  # 全量自检,输出 JSON
bash ~/.claude/skills/github/scripts/preflight.sh --skip-git       # 纯 API/只读场景,跳过 git 身份检查
bash ~/.claude/skills/github/scripts/preflight.sh --need-scopes delete_repo   # 预检额外 scope
```

- 读 JSON 的 `ok` 字段;`ok=false` 时按 `problems` 数组**逐条处理**后再继续。
- 退出码(多项失败按优先级取最高):

| 码 | 含义 | 处理 |
|---|---|---|
| 0 | 就绪 | 继续 |
| 1 | gh 未安装 | 提示用户 `brew install gh` |
| 2 | 未认证 | 指引用户 `gh auth login`(需浏览器,**你不代跑**) |
| 3 | api.github.com 不通 | 检查网络/代理,稍后重试 |
| 4 | git 身份未配置 | 引导 per-repo 设置(见 §8,不擅自动全局) |
| 5 | 缺所需 scope | 指引用户 `gh auth refresh -s <scope>`(需浏览器,**你不代跑**) |

> `gh auth login` / `gh auth refresh` 都需要浏览器交互,你**只能指引用户自己运行**,不得代跑。

## 3. 通用约定

- **仓库定位**:在仓库目录内 gh 自动从 git remote 推断;否则一律显式 `-R owner/repo`。owner/repo 未知时先 `gh repo list <user> --json nameWithOwner` 或 `gh search repos` 查。
- **机器可读输出**:list/view 类一律加 `--json <fields>`,必要时配 `--jq '<expr>'` 收敛;通用端点用 `gh api <path> --jq '...'`。
- **非交互参数必须传全**:你没有 TTY。`gh issue create` / `gh pr create` 必须带 `--title --body`(PR 还要 `--base`),否则会进交互/editor 卡死。
  - 坑:`gh pr create --dry-run` 预览时**仍可能已 push 本地提交**,push 前同样要用户确认。
- **gh 没有全局 `--yes`**:少数命令自带(`gh repo delete --yes`、`gh release delete --yes`、`gh gist delete --yes`);merge/close 类没有确认 flag,**确认责任在你**(见 §8)。
- **大输出防护**:`gh run view --log`、`gh pr diff` 可能上万行,必须用 `--log-failed` / `| tail -n 200` / `--jq` 收敛后再读,**禁止整段灌入上下文**。
- **token 卫生**:**绝不**把 `gh auth token` 的输出打印进回复、日志或写进任何文件/commit。
- git 凭证由 gh 登录后自动配置的 credential helper 处理;环境无 `GITHUB_TOKEN` 变量,**不要创建**。

## 4. 命令参考 A — Issues

| 操作 | 命令 | 级别 |
|---|---|---|
| 创建 | `gh issue create -R o/r --title "T" --body "B" [--label bug] [--assignee @me]` | 写 |
| 查看 | `gh issue view 12 -R o/r --json title,state,body,labels,comments` | 读 |
| 列表 | `gh issue list -R o/r --state open [--label bug] [--assignee @me] --limit 20 --json number,title,state` | 读 |
| 评论 | `gh issue comment 12 -R o/r --body "..."`(长文用 `--body-file`) | 写 |
| 关闭 | `gh issue close 12 -R o/r --reason completed --comment "..."`(`--reason`: completed/not_planned) | 破坏 |
| 重开 | `gh issue reopen 12 -R o/r` | 写 |
| 改标签/指派 | `gh issue edit 12 -R o/r --add-label bug --add-assignee <login>` | 写 |
| 标签管理 | `gh label list -R o/r`;`gh label create <名> --color FF0000 -R o/r` | 读/写 |

## 5. 命令参考 B — PR 全流程

| 操作 | 命令 | 级别 |
|---|---|---|
| 本地分支/提交 | `git switch -c feat/x` → 改码 → `git add -A && git commit -m "..."` → `git push -u origin feat/x`(提交前先确认 git 身份,见 §2/§8) | 写(push) |
| 创建 | `gh pr create -R o/r --title "T" --body "B" --base main [--draft] [--reviewer x]`;偷懒可用 `--fill`;关联 issue 写进 body:`Fixes #12` | 写 |
| 查看/列表 | `gh pr view 34 -R o/r --json number,title,state,mergeable,mergeStateStatus,reviewDecision,statusCheckRollup,headRefName`;`gh pr status`(与我相关);`gh pr list --search "is:open review:required"` | 读 |
| checkout | `gh pr checkout 34`(本地操作,自动建跟踪分支——会改本地工作区,执行前口头说明) | 读 |
| diff / CI | `gh pr diff 34`(大 diff 收敛读);`gh pr checks 34 [--watch --fail-fast]` | 读 |
| 提交评审 | `gh pr review 34 --approve\|--request-changes\|--comment --body "..."`(深度评审先用 `/review`,本命令只负责把结论提交上去) | 写 |
| 合并 | `gh pr merge 34 --squash --delete-branch -R o/r`;`-m` merge commit / `-r` rebase;可 `--auto` 挂自动合并;**禁止主动加 `--admin` 绕过分支保护**,除非用户明确要求 | 破坏 |
| 其他 | `gh pr ready 34`;`gh pr edit 34 --add-reviewer x`;`gh pr update-branch 34`;`gh pr close 34`(破坏) | — |

## 6. 命令参考 C — Actions

| 操作 | 命令 | 级别 |
|---|---|---|
| 列 workflow | `gh workflow list -R o/r --json name,state,id` | 读 |
| 触发 | `gh workflow run <name.yml> -R o/r --ref <branch> [-f key=value ...]`(workflow 须有 `on.workflow_dispatch`) | 写 |
| 拿 run id | 触发后 `gh run list --workflow <name.yml> -R o/r --limit 1 --json databaseId,status,url` | 读 |
| 列运行 | `gh run list -R o/r [--branch feat/x] [--status failure] --json databaseId,name,conclusion,event,headBranch` | 读 |
| 看 run/job | `gh run view <run-id> -R o/r --json jobs --jq '.jobs[] | {name,conclusion,databaseId}'` | 读 |
| 看日志 | `gh run view <run-id> --log-failed \| tail -n 150`;单 job:`gh run view --job <databaseId> --log` | 读 |
| 重跑 | `gh run rerun <run-id> --failed` 或 `--job <databaseId>` | 写 |
| 跟踪 | `gh run watch <run-id> --exit-status [--compact]` | 读 |
| 取消 | `gh run cancel <run-id>` | 写 |

> **关键坑**:`--job` 要传 **`databaseId`**,不是浏览器 URL 里 `/jobs/<number>` 的那个 number,否则 API 404。databaseId 用 `gh run view <run-id> --json jobs --jq '.jobs[] | {name,databaseId}'` 取。

## 7. 命令参考 D — 仓库 / Release / Gist / 搜索

| 操作 | 命令 | 级别 |
|---|---|---|
| 建仓 | `gh repo create <name> --private --clone --add-readme [--gitignore Python]` | 写 |
| 克隆/fork | `gh repo clone o/r [<dir>]`;`gh repo fork o/r --clone` | 写(fork) |
| 查看/列表 | `gh repo view o/r --json nameWithOwner,defaultBranchRef,visibility`;`gh repo list <owner> --limit 30 --json nameWithOwner` | 读 |
| Release | `gh release create v1.0.0 -R o/r --generate-notes [--title "..." --notes-file notes.md] [asset.zip ...]`;`gh release list/view/download` | 写(create) |
| Gist | `gh gist create <file> --desc "..."`(默认 secret;`--public` 公开更需确认);`gh gist list/view <id>` | 写(create) |
| 搜索 | `gh search repos "q" [--owner x] --json fullName,description`;`gh search issues "q" --repo o/r --json number,title`;`gh search code "symbol" --repo o/r`;`gh search prs "q" --state open`;排除限定符用 `--`:`gh search issues -- "foo -label:bug"` | 读 |
| 兜底 API | `gh api repos/o/r/<...> [--method POST] [-f key=value] [--jq ...]`;写方法(POST/PUT/PATCH/DELETE)同样走 §8 确认 | 视方法 |

## 8. 工作流编排 + 安全约定 + 权限边界 + 排错

### 8.1 常用工作流

**W1. 从需求到合并 PR**
1. `preflight.sh`(含 git 身份检查;缺失先引导 per-repo 设置,需用户同意)
2. `gh issue create ...`(如需,先确认)或用用户给定的 issue 号
3. `git switch -c fix-12` → 改码 → `git commit` → **向用户复述 push 内容并确认** → `git push -u origin fix-12`
4. `gh pr create --title ... --body "Fixes #12" --base main`(确认)
5. `gh pr checks <n> --watch --fail-fast` 看 CI
6. 需深度评审 → 交 `/review`;轻量自查用 `gh pr diff` + `gh pr view`
7. CI 绿且用户确认 → `gh pr merge <n> --squash --delete-branch`

**W2. 排查失败的 Actions**
1. `gh run list --status failure --limit 5 --json databaseId,name,conclusion,headBranch` 定位
2. `gh run view <id> --json jobs --jq '.jobs[] | {name,conclusion,databaseId}'` 找失败 job
3. `gh run view --job <databaseId> --log | tail -n 200` 或 `--log-failed` 取日志
4. 分析根因(代码错 / 依赖拉取失败 / runner 资源 / 上游 secret 缺失),给修复建议
5. 代码问题 → 修复并走 W1 步骤 3–5 触发新 run,`gh run watch`;基建偶发 → 用户确认后 `gh run rerun <id> --failed`

**W3. 接手别人的 PR 并验证**
1. `gh pr view <n> --json title,body,headRefName,statusCheckRollup` 读背景
2. `gh pr checkout <n>`(口头说明会改本地工作区)→ 本地跑测试
3. `gh pr diff <n>` + `gh pr checks <n>`;深度评审转 `/review`
4. 结论提交:`gh pr review <n> --approve|--request-changes --body "..."`(确认)

### 8.2 写操作安全分级

| 级别 | 示例 | 要求 |
|---|---|---|
| 只读(免确认) | `* view/list/status/diff/checks`、`gh run view/list/watch`、`gh search *`、`gh api`(GET)、`git fetch`、`gh pr checkout` | 直接执行(checkout 口头说明一句) |
| 远程写(必须确认) | `gh pr create/merge/close/reopen/edit/review/ready`、`gh issue create/close/reopen/edit/comment`、`gh repo create/fork/sync`、`gh release create`、`gh run cancel/rerun`、`gh workflow run/enable/disable`、`gh label create/edit`、`gh gist create/edit`、`git push`、`gh api -X POST/PUT/PATCH` | 自然语言复述确认后才执行 |
| 破坏性(确认+提示不可逆) | `gh repo delete/archive/unarchive/rename`、`gh release delete`、`gh issue delete`、`gh pr close`、`gh run delete`、`gh label delete`、`gh gist delete`、`gh api -X DELETE` | 复述时明写"不可逆/影响协作者",用户明确说继续才执行 |

**确认话术模板**(操作/命令/对象/影响四行):

> 准备执行 GitHub 写操作,请确认:
> - 操作:合并 PR #34(squash,合并后删除源分支)
> - 命令:`gh pr merge 34 --squash --delete-branch -R owner/repo`
> - 对象:owner/repo
> - 影响:3 个 commit 进入 main;远程分支 feat/x 被删除;会触发 main 的 CI
>
> 确认后我才执行。

**硬性规则**:
1. 一次确认只覆盖当次破坏性操作,不得批量打包确认;用户说"全部照办"也要逐项列清单再确认。
2. `gh pr merge` 禁止主动加 `--admin` 绕过分支保护,除非用户明确要求。
3. 写操作失败后先用只读命令(`view`/`checks`/`gh api GET`)核对实际状态,禁止盲目重试(防重复 merge、重复触发 workflow)。
4. git 身份配置默认建议 per-repo(`git config user.name`),**不擅自动 `--global`**;动全局前单独确认。
5. 任何情况下不打印 token(`gh auth token`)、不把 token 写进文件或 commit。

### 8.3 权限边界(当前 scopes 能做/不能做)

**不能**(遇 403 / "Resource not accessible" **一律不重试**):
- `gh repo delete`(需 `delete_repo`)
- org 管理(需 `admin:org`)
- `gh secret set`(需 `admin:repo`)
- 部分受保护分支/仓库设置管理

遇到缺权限:向用户解释缺失的 scope,给出 `gh auth refresh -s <scope>` 指引(需浏览器,用户自己跑)。

### 8.4 排错

| 现象 | 原因 / 处理 |
|---|---|
| preflight exit 4 | git 身份缺失 → per-repo 设置(需用户同意) |
| 401 / token 过期 | `gh auth refresh`(用户自己跑) |
| 私有仓库 404 | token 无该仓库权限,确认仓库归属与 scope |
| commit 报 "Please tell me who you are" | git 身份未配 → per-repo 设置 |
| merge 报 not mergeable | 查 `gh pr view --json mergeStateStatus`,或 `gh pr update-branch` 后重试 |
| `rerun --job` 404 | 传的是 URL number 而非 databaseId(见 §6 坑) |
| `gh pr create` 卡住 | 参数没传全(缺 `--title`/`--body`/`--base`),进了交互 |

## 9. 交付规范：README + 敏感信息检查（每次任务必做）

本 skill 产出代码 / 仓库 / 工具后，push 前必须完成以下两项。

### 9.1 README（必写，精简三段式）

每个仓库 / 工具必须配套 `README.md`，尽量精简，只保留与使用直接相关的内容：

1. **一句话简介**：项目是什么、解决什么问题（1–2 句，不铺陈背景）。
2. **AI 安装提示词**：一段可直接复制发给 AI 的安装引导，含 clone / 下载指令、环境变量（用占位符，如 `<你的 api key>`）、验证步骤。
3. **人工安装方式**：放在**最下方**，给不依赖 AI 的用户的手动步骤（clone → 装依赖 → 配置 → 运行）。

### 9.2 敏感信息检查（push 前必做）

账号、密码、API key、token、数据库连接串、密钥等敏感信息**一律不进仓库**。push 前自行判断哪些属于敏感信息，并核对：

| 检查项 | 做法 |
|---|---|
| 敏感文件被 git 忽略 | 确认 `.gitignore` 已忽略各类敏感文件 |
| README 无真实凭据 | 敏感值一律用占位符 `<...>`，不写真值 |
| 已跟踪 / 本次提交无敏感文件 | `git ls-files`、`git add -n .`、`git diff --cached --name-only` 逐项核对，确认无敏感内容 |

**硬性规则**：
1. commit 前先跑上表检查，通过才 push。
2. 发现敏感文件已入库：`git rm --cached <文件>` 并补 `.gitignore`；同时提醒用户——key / 密码一旦进过远程仓库即视为泄露，改 `.gitignore` 无法撤销，需评估是否轮换。
