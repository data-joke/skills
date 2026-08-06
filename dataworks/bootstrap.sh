#!/usr/bin/env bash
# DataWorks skill 一键安装:装 SDK(全局 python)+ 生成 .env + 自检
set -e
cd "$(dirname "$0")"

# 兼容 Windows(Git Bash 可能只有 python,无 python3;且 WindowsApps/python3 是
# Microsoft Store 假入口,运行即失败)。用"实际能否运行"探测,而非 `command -v`,否则会被假入口骗过。
PY="python3"
if ! "$PY" --version >/dev/null 2>&1; then PY="python"; fi
echo "== 使用解释器: $PY =="

echo "== [1/3] 安装 DataWorks SDK(全局 $PY)=="
"$PY" -m pip install -r requirements.txt

echo "== [2/3] 生成 .env 模板(若不存在)=="
if [ ! -f .env ]; then
  cp config.example.env .env
  echo "已创建 .env,请编辑填入你的 AccessKey(或复用已设置的 DATAWORKS_* / ALIBABA_CLOUD_* 环境变量)。"
else
  echo ".env 已存在,跳过。"
fi

echo "== [3/3] 自检 =="
"$PY" scripts/dwcli.py doctor || true

echo ""
echo "完成。配置好凭证后,用以下命令验证:"
echo "  $PY scripts/dwcli.py project list"
