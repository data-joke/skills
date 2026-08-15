#!/usr/bin/env bash
# github skill 环境自检脚本
# 用法: preflight.sh [--skip-git] [--need-scopes scope1,scope2]
#   --skip-git            跳过 git 身份检查(纯 gh API / 只读场景)
#   --need-scopes a,b     在基线 scope 之外追加预检的 scope
# 输出: JSON 到 stdout(诊断信息走 stderr)
# 退出码(多项失败按优先级取最高 1>2>5>4>3):
#   0 就绪 / 1 gh 未安装 / 2 未认证 / 3 网络不可达 / 4 git 身份缺失 / 5 scope 不足
set -u

SKIP_GIT=0
NEED=""
while [ $# -gt 0 ]; do
  case "$1" in
    --skip-git) SKIP_GIT=1 ;;
    --need-scopes) NEED="${2:-}"; shift ;;
    *) echo "未知参数: $1" >&2 ;;
  esac
  shift
done

# ---------- 工具函数 ----------
json_escape() {
  local s=$1
  s=${s//\\/\\\\}
  s=${s//\"/\\\"}
  s=${s//$'\n'/\\n}
  s=${s//$'\t'/\\t}
  printf '%s' "$s"
}
# 渲染 bash 数组为 JSON 字符串数组(跳过空元素,空数组输出 [])
json_str_array() {
  local out="" first=1 x
  for x in "$@"; do
    [ -z "$x" ] && continue
    if [ $first -eq 1 ]; then first=0; else out="$out,"; fi
    out="$out\"$(json_escape "$x")\""
  done
  printf '[%s]' "$out"
}

# ---------- 状态标志(true/false) ----------
HAS_GH=false
AUTHED=false
SCOPES_OK=true
GIT_OK=true
NET_OK=false
PROBLEMS=()

# ---------- 1. gh 安装检查 ----------
GH_PATH=""
GH_VER=""
if command -v gh >/dev/null 2>&1; then
  HAS_GH=true
  GH_PATH=$(command -v gh)
  GH_VER=$(gh --version 2>/dev/null | head -1 | awk '{print $3}')
else
  PROBLEMS+=("gh CLI 未安装:请运行 brew install gh")
fi

# ---------- 2. 认证与 scope 检查 ----------
LOGIN=""
SCOPES=()
MISSING=()
if [ "$HAS_GH" = true ]; then
  AUTH_OUT=$(gh auth status 2>&1 || true)
  LOGIN=$(printf '%s' "$AUTH_OUT" | grep -E 'Logged in to' | grep -oE 'account [^ ]+' | head -1 | awk '{print $2}')
  if [ -n "$LOGIN" ]; then
    AUTHED=true
    # 解析 scopes:行形如  Token scopes: 'gist', 'read:org', 'repo', 'workflow'
    SCOPE_LINE=$(printf '%s' "$AUTH_OUT" | grep -E 'Token scopes:' | head -1 | grep -oE "'[^']+'" | tr -d "'" || true)
    while IFS= read -r s; do
      [ -n "$s" ] && SCOPES+=("$s")
    done <<< "$SCOPE_LINE"

    # 基线 + 追加的必需 scope
    REQUIRED="repo workflow gist read:org"
    if [ -n "$NEED" ]; then
      REQUIRED="$REQUIRED ${NEED//,/ }"
    fi
    SCOPE_LIST=" ${SCOPES[*]:-} "
    for s in $REQUIRED; do
      case "$SCOPE_LIST" in
        *" $s "*) ;;
        *) MISSING+=("$s") ;;
      esac
    done
    if [ ${#MISSING[@]} -gt 0 ]; then
      SCOPES_OK=false
      PROBLEMS+=("缺少 scope: ${MISSING[*]}。请用户运行 gh auth refresh -s ${MISSING[*]// /,}(需浏览器)")
    fi
  else
    PROBLEMS+=("gh 未认证:请用户运行 gh auth login(需浏览器交互,Claude 不代跑)")
  fi
fi

# ---------- 3. git 身份检查 ----------
GIT_NAME=""
GIT_EMAIL=""
GIT_SOURCE="missing"
if [ "$SKIP_GIT" = 1 ]; then
  GIT_SOURCE="skipped"
else
  LOCAL_NAME=$(git config --local user.name 2>/dev/null || true)
  LOCAL_EMAIL=$(git config --local user.email 2>/dev/null || true)
  GLOBAL_NAME=$(git config --global user.name 2>/dev/null || true)
  GLOBAL_EMAIL=$(git config --global user.email 2>/dev/null || true)
  if [ -n "$LOCAL_NAME" ] && [ -n "$LOCAL_EMAIL" ]; then
    GIT_NAME=$LOCAL_NAME; GIT_EMAIL=$LOCAL_EMAIL; GIT_SOURCE="repo"
  elif [ -n "$GLOBAL_NAME" ] && [ -n "$GLOBAL_EMAIL" ]; then
    GIT_NAME=$GLOBAL_NAME; GIT_EMAIL=$GLOBAL_EMAIL; GIT_SOURCE="global"
  else
    GIT_OK=false
    PROBLEMS+=("git 身份未配置,本地 commit 会失败。建议在仓库目录内 per-repo 设置(需用户同意): git config user.name 'X' && git config user.email 'X@Y'")
  fi
fi
GIT_HINT="在仓库目录内: git config user.name 'X' && git config user.email 'X@Y' (per-repo,需用户同意)"

# ---------- 4. 网络连通性 ----------
HTTP=$(curl -s -o /dev/null -m 8 -w '%{http_code}' https://api.github.com/rate_limit 2>/dev/null || echo 000)
if [ "$HTTP" != "000" ]; then
  NET_OK=true
else
  PROBLEMS+=("api.github.com 不可达:请检查网络/代理后重试")
fi

# ---------- 计算退出码(优先级 1>2>5>4>3) ----------
EXIT=0
[ "$HAS_GH" = false ] && EXIT=1
[ $EXIT -eq 0 ] && [ "$AUTHED" = false ] && EXIT=2
[ $EXIT -eq 0 ] && [ "$SCOPES_OK" = false ] && EXIT=5
[ $EXIT -eq 0 ] && [ "$GIT_OK" = false ] && EXIT=4
[ $EXIT -eq 0 ] && [ "$NET_OK" = false ] && EXIT=3

OK=false
[ $EXIT -eq 0 ] && OK=true

# ---------- 渲染 JSON ----------
LOGIN_JSON=null
[ -n "$LOGIN" ] && LOGIN_JSON="\"$(json_escape "$LOGIN")\""
NAME_JSON=null
[ -n "$GIT_NAME" ] && NAME_JSON="\"$(json_escape "$GIT_NAME")\""
EMAIL_JSON=null
[ -n "$GIT_EMAIL" ] && EMAIL_JSON="\"$(json_escape "$GIT_EMAIL")\""

cat <<EOF
{
  "ok": $OK,
  "gh": {"installed": $HAS_GH, "version": "$(json_escape "$GH_VER")", "path": "$(json_escape "$GH_PATH")"},
  "auth": {"authenticated": $AUTHED, "login": $LOGIN_JSON, "scopes": $(json_str_array "${SCOPES[@]:-}"), "missing_scopes": $(json_str_array "${MISSING[@]:-}")},
  "git_identity": {"name": $NAME_JSON, "email": $EMAIL_JSON, "source": "$GIT_SOURCE", "hint": "$(json_escape "$GIT_HINT")"},
  "network": {"api_reachable": $NET_OK},
  "problems": $(json_str_array "${PROBLEMS[@]:-}")
}
EOF

echo "[preflight] ok=$OK exit=$EXIT" >&2
exit $EXIT
