# github

让 Claude Code 通过 `gh` CLI 操作 GitHub 的 skill：Issue / PR 全流程 / Actions 运维 / Release / 搜索，内置环境自检（`preflight.sh`）与写操作安全确认。

## AI 安装提示词

> 复制下面整段发给 AI，让它引导你完成安装。

````markdown
使用如下指令下载并安装 github skill：

```bash
git clone https://github.com/data-joke/skills.git
mkdir -p ~/.claude/skills
cp -r skills/github ~/.claude/skills/github
```

完成前置依赖与认证（认证需用户在浏览器里自己完成，AI 不代跑）：

```bash
brew install gh        # 安装 gh CLI
gh auth login          # 浏览器 OAuth 登录（用户自己跑）
```

验证安装：

```bash
bash ~/.claude/skills/github/scripts/preflight.sh
```

输出 JSON 里 `ok=true` 即就绪；否则按 `problems` 逐条处理——缺 scope 用 `gh auth refresh -s <scope>`，git 身份缺失用 `git config user.name / user.email`。
````

## 人工安装方式

1. 装 gh CLI：`brew install gh`
2. 登录授权：`gh auth login`（浏览器 OAuth，scope 需含 `repo workflow gist read:org`）
3. 安装 skill：把本目录 `github/`（`SKILL.md` + `scripts/`）复制到 `~/.claude/skills/github/`
4. 配 git 身份：`git config user.name "X"` / `git config user.email "X@Y"`
5. 验证：`bash ~/.claude/skills/github/scripts/preflight.sh`，看 `ok=true`

> 脚本路径在 SKILL.md 里固定为 `~/.claude/skills/github/`，请勿改安装位置。
