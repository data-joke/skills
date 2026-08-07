#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""dwcli — 阿里云 DataWorks(2020-05-18,公共云)命令行工具。

供 Claude Code 的 dataworks skill 调用。设计原则:
  * JSON 优先输出(便于 Claude 解析与串联),--format table/text 供人查看
  * 写操作(restart/stop/set-success/complement run/file create/submit/deploy)
    需 --yes 确认;--dry-run 仅打印将执行的请求、不真正调用
  * 守护式 SDK import:缺依赖时给出安装指引并以退出码 6 退出
  * 凭证默认读取 DATAWORKS_ACCESS_KEY_ID/SECRET,回退 ALIBABA_CLOUD_ACCESS_KEY_ID/SECRET

退出码:
  0 成功 · 1 API/一般错误 · 2 用法错误或缺 --yes · 3 鉴权
  4 未找到/零匹配 · 5 无权限 · 6 SDK 未安装 · 7 超时
"""
import argparse
import collections
import datetime
import json
import os
import re
import sys
import time
from types import SimpleNamespace

_HERE = os.path.dirname(os.path.abspath(__file__))
# 兼容根目录与 scripts/ 两种布局:.env 与 requirements.txt 始终在技能根目录
SKILL_DIR = os.path.dirname(_HERE) if os.path.basename(_HERE) == "scripts" else _HERE

# ---- 退出码 ----
E_OK, E_ERR, E_USAGE, E_AUTH, E_NOTFOUND, E_PERM, E_SDK, E_TIMEOUT = 0, 1, 2, 3, 4, 5, 6, 7


def eprint(*a, **k):
    print(*a, file=sys.stderr, **k)


def die(msg, code=E_ERR, **extra):
    payload = {"error": msg}
    payload.update(extra)
    eprint(json.dumps(payload, ensure_ascii=False))
    sys.exit(code)


def mask(s):
    if not s:
        return None
    s = str(s)
    return s[:4] + "****" + s[-4:] if len(s) > 8 else "****"


# ---- .env 加载(不引第三方依赖;进程 env 优先于 .env)----
def load_dotenv():
    env = {}
    path = os.path.join(SKILL_DIR, ".env")
    if os.path.exists(path):
        try:
            with open(path, encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if not line or line.startswith("#") or "=" not in line:
                        continue
                    k, v = line.split("=", 1)
                    env[k.strip()] = v.strip().strip('"').strip("'")
        except Exception:
            pass
    return env


DOTENV = load_dotenv()


def envget(*names):
    for n in names:
        if os.environ.get(n):
            return os.environ[n]
    for n in names:
        if DOTENV.get(n):
            return DOTENV[n]
    return None


# ---- 守护式 SDK import ----
try:
    from alibabacloud_tea_openapi import models as oapi_models
    from alibabacloud_dataworks_public20200518.client import Client as DwClient
    from alibabacloud_dataworks_public20200518 import models as dw_models
    SDK_OK = True
    SDK_ERR = None
except Exception as _e:  # noqa: BLE001
    SDK_OK = False
    SDK_ERR = repr(_e)
    oapi_models = DwClient = dw_models = None


def require_sdk():
    if not SDK_OK:
        die(
            f"DataWorks SDK 未安装或导入失败: {SDK_ERR}。请运行: "
            f"python3 -m pip install -r {os.path.join(SKILL_DIR, 'requirements.txt')}",
            E_SDK,
        )


# ---- 区域/上下文解析 ----
def region_from_endpoint(ep):
    """从 MaxCompute endpoint(如 http://service.cn-shanghai.maxcompute.aliyun.com/api)解析 region。"""
    if not ep:
        return None
    m = re.search(r"service\.([a-z0-9-]+?)\.maxcompute", str(ep))
    if m:
        return re.sub(r"-(vpc|intranet)$", "", m.group(1))
    return None


def build_ctx(args):
    ak = args.access_key_id or envget(
        "DATAWORKS_ACCESS_KEY_ID", "ALIBABA_CLOUD_ACCESS_KEY_ID"
    )
    sk = args.access_key_secret or envget(
        "DATAWORKS_ACCESS_KEY_SECRET", "ALIBABA_CLOUD_ACCESS_KEY_SECRET"
    )
    region = (
        args.region
        or envget("DATAWORKS_REGION_ID")
        or envget("ALIBABA_CLOUD_REGION_ID")
        or "cn-shanghai"
    )
    endpoint = args.endpoint or envget("DATAWORKS_ENDPOINT") or f"dataworks.{region}.aliyuncs.com"
    project_id = args.project_id or envget("DATAWORKS_PROJECT_ID") or ""
    env = args.env or envget("DATAWORKS_PROJECT_ENV") or "PROD"
    fmt = args.format or envget("DATAWORKS_FORMAT") or "json"
    return SimpleNamespace(
        ak=ak, sk=sk, region=region, endpoint=endpoint, project_id=project_id,
        env=env, fmt=fmt, timeout=getattr(args, "timeout", 30),
        debug=getattr(args, "debug", False), dry_run=getattr(args, "dry_run", False),
        yes=getattr(args, "yes", False),
    )


def need_project(ctx):
    if not ctx.project_id:
        die("缺少 projectId:用 --project-id 或 DATAWORKS_PROJECT_ID(可先 `dw project list` 查询)。", E_USAGE)
    try:
        return int(ctx.project_id)
    except ValueError:
        die(f"projectId 需为数字,当前值: {ctx.project_id}", E_USAGE)


def make_client(ctx):
    require_sdk()
    if not ctx.ak or not ctx.sk:
        die(
            "缺少 AccessKey。请设置 DATAWORKS_ACCESS_KEY_ID/SECRET(或 "
            "ALIBABA_CLOUD_ACCESS_KEY_ID/SECRET),或在技能目录 .env 中配置。",
            E_AUTH,
        )
    cfg = oapi_models.Config(
        access_key_id=ctx.ak, access_key_secret=ctx.sk,
        region_id=ctx.region, endpoint=ctx.endpoint,
    )
    cfg.read_timeout = ctx.timeout * 1000
    cfg.connect_timeout = ctx.timeout * 1000
    return DwClient(cfg)


# ---- 大小写不敏感的嵌套取值(DataWorks 响应为 PascalCase,稳妥起见忽略大小写)----
def g(d, *keys, default=None):
    cur = d
    for k in keys:
        if not isinstance(cur, dict):
            return default
        nxt = None
        for ck, cv in cur.items():
            if str(ck).lower() == str(k).lower():
                nxt = cv
                break
        if nxt is None:
            return default
        cur = nxt
    return cur


def extract_list(body, key):
    """DataWorks 列表接口信封不一致:多数用 Data.<key>,ListProjects 用 PageResult.<key>。"""
    for path in (("Data", key), ("PageResult", key), (key,)):
        v = g(body, *path, default=None)
        if isinstance(v, list):
            return v
    return []


def handle_api_error(e):
    code = str(getattr(e, "code", "") or "")
    msg = getattr(e, "message", None) or str(e)
    lc = code.lower()
    data = getattr(e, "data", None)
    extra = {"api_code": code}
    if data:
        extra["data"] = data
    if any(x in lc for x in ("invalidaccesskey", "signaturedoesnotmatch", "missingaccesskey",
                             "noaccesskey", "accesskeyid", "unauthorized")):
        die(msg, E_AUTH, **extra)
    if any(x in lc for x in ("forbidden", "nopermission", "notauthorized", "accessdenied", "ram.")):
        die(msg, E_PERM, **extra)
    if "timeout" in lc or "timedout" in lc:
        die(msg, E_TIMEOUT, **extra)
    if any(x in lc for x in ("notfound", "notexist", "invalid.entity", "entitynotexist")):
        die(msg, E_NOTFOUND, **extra)
    die(msg, E_ERR, **extra)


def call(ctx, method_name, request, retries=2):
    """调用 API。遇 Throttling(417 限流)自动退避重试,缓解全量遍历命令的中途失败。"""
    client = make_client(ctx)
    if ctx.debug:
        try:
            eprint("[debug] " + method_name + " " + json.dumps(request.to_map(), ensure_ascii=False, default=str))
        except Exception:
            eprint("[debug] " + method_name)
    resp = None
    for attempt in range(retries + 1):
        try:
            resp = getattr(client, method_name)(request)
            break
        except SystemExit:
            raise
        except Exception as e:  # noqa: BLE001
            code = str(getattr(e, "code", "") or "").lower()
            if "throttl" in code and attempt < retries:
                time.sleep(1.5 * (attempt + 1))
                continue
            handle_api_error(e)
    body = getattr(resp, "body", resp)
    if hasattr(body, "to_map"):
        body = body.to_map()
    # 逻辑失败(HTTP 200 但 Success=false)
    if isinstance(body, dict) and g(body, "Success") is False:
        die(f"API 返回失败: {g(body, 'ErrorCode')} {g(body, 'ErrorMessage')}", E_ERR,
            code=g(body, "ErrorCode"))
    return body


# ---- 输出 ----
def ts_str(ms):
    """epoch 毫秒 → 'YYYY-MM-DD HH:MM'(空值/非时间戳原样返回),仅用于 table 展示。"""
    if ms is None or ms == "" or ms == "None":
        return ""
    try:
        v = int(ms)
    except (TypeError, ValueError):
        return str(ms)
    if v > 10 ** 11:  # 毫秒级 epoch
        return datetime.datetime.fromtimestamp(v / 1000).strftime("%Y-%m-%d %H:%M")
    return str(ms)


def print_table(rows, cols):
    rows = rows or []
    if not cols:
        cols = list(rows[0].keys()) if rows else []
    widths = {c: max([len(str(c))] + [len(str(r.get(c, ""))) for r in rows]) for c in cols}
    print("  ".join(str(c).ljust(widths[c]) for c in cols))
    print("  ".join("-" * widths[c] for c in cols))
    for r in rows:
        print("  ".join(str(r.get(c, "")).ljust(widths[c]) for c in cols))
    eprint(f"({len(rows)} rows)")


def out(ctx, obj, table_cols=None):
    if ctx.fmt == "table" and table_cols:
        print_table(obj if isinstance(obj, list) else [obj], table_cols)
    elif ctx.fmt == "text":
        print(obj if isinstance(obj, str) else json.dumps(obj, ensure_ascii=False, default=str))
    else:
        print(json.dumps(obj, ensure_ascii=False, indent=2, default=str))


# ---- 写操作门控 ----
def confirm_write(ctx, action, request_map):
    preview = {"WOULD_EXECUTE": action, "request": request_map}
    if ctx.dry_run:
        print(json.dumps(preview, ensure_ascii=False, indent=2, default=str))
        sys.exit(E_OK)
    if not ctx.yes:
        eprint(json.dumps(preview, ensure_ascii=False, indent=2, default=str))
        die("写操作需确认:加 --yes 执行,或加 --dry-run 仅预览(不调用)。", E_USAGE)


# ---- 日期工具 ----
def to_epoch_ms(v):
    if v is None or v == "":
        return None
    if str(v).isdigit():
        return int(v)
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
        try:
            return int(datetime.datetime.strptime(v, fmt).timestamp() * 1000)
        except ValueError:
            continue
    die(f"无法解析日期: {v}(请用 YYYY-MM-DD 或 epoch 毫秒)", E_USAGE)


def norm_biz_datetime(v):
    """补数据 StartBizDate/EndBizDate 规范为 'YYYY-MM-DD HH:MM:SS'。"""
    if v is None:
        return None
    v = str(v)
    if re.match(r"^\d{4}-\d{2}-\d{2}$", v):
        return v + " 00:00:00"
    return v


def to_biz_datetime(v, end_of_day=False):
    """实例查询 bizdate:服务端要 'YYYY-MM-DD HH:MM:SS' 字符串(非 epoch 毫秒)。"""
    if v is None or v == "":
        return None
    v = str(v)
    if re.match(r"^\d{4}-\d{2}-\d{2}$", v):
        return v + (" 23:59:59" if end_of_day else " 00:00:00")
    return v


# ---- 核心:名称 → 节点 解析 ----
def resolve_nodes(ctx, name, exact=False, include_files=True, limit=50):
    pid = need_project(ctx)
    results = {}
    # 1) 调度节点 by node_name
    req = dw_models.ListNodesRequest(project_id=pid, node_name=name, project_env=ctx.env,
                                     page_number=1, page_size=limit)
    body = call(ctx, "list_nodes", req)
    for n in extract_list(body, "Nodes"):
        nid = g(n, "NodeId")
        results[nid] = {
            "node_id": nid, "node_name": g(n, "NodeName"),
            "owner": g(n, "OwnerId") or g(n, "Owner"),
            "program_type": g(n, "ProgramType"), "project_id": g(n, "ProjectId"),
            "cron_express": g(n, "CronExpress"), "file_id": g(n, "FileId"),
        }
    # 2) 文件 by keyword,补 file_id
    if include_files:
        freq = dw_models.ListFilesRequest(project_id=pid, keyword=name, page_number=1, page_size=limit)
        fbody = call(ctx, "list_files", freq)
        for f in extract_list(fbody, "Files"):
            nid = g(f, "NodeId")
            if nid and nid in results:
                results[nid]["file_id"] = g(f, "FileId") or results[nid]["file_id"]
            elif nid:
                results[nid] = {
                    "node_id": nid, "node_name": g(f, "FileName"),
                    "owner": g(f, "Owner"), "program_type": g(f, "FileType"),
                    "project_id": g(f, "ProjectId"), "file_id": g(f, "FileId"),
                    "cron_express": None,
                }
    out_list = list(results.values())
    if exact:
        out_list = [r for r in out_list if str(r.get("node_name") or "").lower() == str(name).lower()]
    return out_list


def resolve_single_node(ctx, name):
    cands = resolve_nodes(ctx, name, exact=True)
    if not cands:
        cands = resolve_nodes(ctx, name, exact=False)
    if not cands:
        die(f"未找到名为 '{name}' 的节点(project={ctx.project_id}, env={ctx.env})。", E_NOTFOUND)
    if len(cands) > 1:
        eprint(json.dumps({"candidates": cands}, ensure_ascii=False, indent=2))
        die(f"'{name}' 匹配到 {len(cands)} 个节点,请用更精确的名称或直接传 --node-id。", E_USAGE)
    return cands[0]


# ---- 实例查询(复用于 list 与按任务名写操作)----
def find_instances(ctx, node_id=None, node_name=None, status=None, dag_id=None,
                   begin_bizdate=None, end_bizdate=None, page_size=100):
    pid = need_project(ctx)
    req = dw_models.ListInstancesRequest(project_id=pid, project_env=ctx.env,
                                         page_number=1, page_size=min(max(int(page_size), 1), 100))
    if node_id is not None:
        req.node_id = int(node_id)
    if node_name:
        req.node_name = node_name
    if status:
        req.status = status
    if dag_id is not None:
        req.dag_id = int(dag_id)
    if begin_bizdate is not None:
        req.begin_bizdate = begin_bizdate
    if end_bizdate is not None:
        req.end_bizdate = end_bizdate
    body = call(ctx, "list_instances", req)
    rows = []
    for it in extract_list(body, "Instances"):
        rows.append({
            "instance_id": g(it, "InstanceId"), "node_id": g(it, "NodeId"),
            "node_name": g(it, "NodeName"), "dag_id": g(it, "DagId"),
            "status": g(it, "Status"), "bizdate": g(it, "Bizdate"),
            "begin_running_time": g(it, "BeginRunningTime"), "finish_time": g(it, "FinishTime"),
            "owner": g(it, "OwnerId") or g(it, "Owner"),
        })
    return rows


def resolve_instance_for_task(ctx, task_name, status):
    node = resolve_single_node(ctx, task_name)
    insts = find_instances(ctx, node_id=node["node_id"], status=status)
    if not insts:
        die(f"节点 '{task_name}'(node_id={node['node_id']})没有状态为 {status} 的实例。", E_NOTFOUND)
    if len(insts) > 1:
        eprint(json.dumps({"candidates": insts}, ensure_ascii=False, indent=2))
        die(f"匹配到 {len(insts)} 个 {status} 实例,请直接传 --instance-id 指定。", E_USAGE)
    return insts[0], node


# ---- FileType 友好名映射(数字编码以 DataWorks 控制台/OpenAPI 为准,可直接传数字)----
FILE_TYPES = {
    "odps-sql": 10, "odps-mr": 11, "odps-script": 24,
    "di": 23, "data-integration": 23,
    "shell": 4, "virtual": 99,
    "pyodps2": 221, "pyodps3": 225, "pyodps": 225,
}


def resolve_file_type(t):
    if t is None:
        return None
    if str(t).isdigit():
        return int(t)
    return FILE_TYPES.get(str(t).lower())


# =========================================================================
# 命令实现
# =========================================================================
def cmd_doctor(ctx, args):
    info = {
        "skill_dir": SKILL_DIR,
        "python": sys.version.split()[0],
        "sdk_installed": SDK_OK,
        "access_key_id": mask(ctx.ak),
        "has_secret": bool(ctx.sk),
        "region": ctx.region,
        "endpoint": ctx.endpoint,
        "project_id": ctx.project_id or "(空 — 设置 DATAWORKS_PROJECT_ID 或 --project-id)",
        "project_env": ctx.env,
        "dotenv_loaded": bool(DOTENV),
    }
    if not SDK_OK:
        info["sdk_error"] = SDK_ERR
    if getattr(args, "check", False) and SDK_OK and ctx.ak and ctx.sk:
        try:
            body = call(ctx, "list_projects", dw_models.ListProjectsRequest(page_number=1, page_size=1))
            info["connectivity"] = "ok"
            info["project_count"] = g(body, "PageResult", "TotalCount", default=len(extract_list(body, "ProjectList")))
        except SystemExit:
            info["connectivity"] = "failed(详见 stderr)"
    out(ctx, info)


def cmd_project_list(ctx, args):
    req = dw_models.ListProjectsRequest(page_number=args.page_number, page_size=args.page_size)
    body = call(ctx, "list_projects", req)
    projects = extract_list(body, "ProjectList")
    rows = [{
        "project_id": g(p, "ProjectId"), "name": g(p, "ProjectName"),
        "identifier": g(p, "ProjectIdentifier"),
        "status": g(p, "ProjectStatusCode") or g(p, "ProjectStatus"),
    } for p in projects]
    out(ctx, rows, ["project_id", "name", "identifier", "status"])


def cmd_node_list(ctx, args):
    pid = need_project(ctx)
    req = dw_models.ListNodesRequest(project_id=pid, project_env=ctx.env,
                                     page_number=args.page_number, page_size=args.page_size)
    if args.name:
        req.node_name = args.name
    if args.owner:
        req.owner = args.owner
    if args.program_type:
        req.program_type = args.program_type
    body = call(ctx, "list_nodes", req)
    nodes = extract_list(body, "Nodes")
    rows = [{
        "node_id": g(n, "NodeId"), "node_name": g(n, "NodeName"),
        "owner": g(n, "OwnerId") or g(n, "Owner"), "program_type": g(n, "ProgramType"),
        "cron": g(n, "CronExpress"),
    } for n in nodes]
    out(ctx, rows, ["node_id", "node_name", "owner", "program_type", "cron"])


def cmd_node_get(ctx, args):
    req = dw_models.GetNodeRequest(node_id=int(args.node_id), project_env=ctx.env)
    out(ctx, call(ctx, "get_node", req))


def cmd_node_resolve(ctx, args):
    res = resolve_nodes(ctx, args.name, exact=args.exact, include_files=True, limit=args.limit)
    if not res:
        die(f"未找到匹配 '{args.name}' 的节点(project={ctx.project_id}, env={ctx.env})。", E_NOTFOUND)
    out(ctx, res, ["node_id", "node_name", "file_id", "owner", "program_type", "cron_express"])


def _node_adjacent(ctx, args, action, method, model_name):
    """node parents / node children 公共逻辑:解析节点 → 查上/下游。"""
    node_id = args.node_id
    if args.task_name and not node_id:
        node = resolve_single_node(ctx, args.task_name)
        node_id = node["node_id"]
    if not node_id:
        die(f"{action} 需要 --node-id 或 --task-name。", E_USAGE)
    req_cls = getattr(dw_models, model_name)
    body = call(ctx, method, req_cls(node_id=int(node_id), project_env=ctx.env))
    rows = [{
        "node_id": g(n, "NodeId"), "node_name": g(n, "NodeName"),
        "owner": g(n, "OwnerId") or g(n, "Owner"), "program_type": g(n, "ProgramType"),
        "cron": g(n, "CronExpress"), "scheduler_type": g(n, "SchedulerType"),
    } for n in extract_list(body, "Nodes")]
    out(ctx, rows, ["node_id", "node_name", "owner", "program_type", "cron", "scheduler_type"])


def cmd_node_parents(ctx, args):
    _node_adjacent(ctx, args, "node parents", "get_node_parents", "GetNodeParentsRequest")


def cmd_node_children(ctx, args):
    _node_adjacent(ctx, args, "node children", "get_node_children", "GetNodeChildrenRequest")


def cmd_file_list(ctx, args):
    pid = need_project(ctx)
    # need_absolute_folder_path=True 才能拿到 AbsoluteFolderPath,否则目录路径为空
    req = dw_models.ListFilesRequest(project_id=pid, page_number=args.page_number, page_size=args.page_size,
                                     need_absolute_folder_path=True)
    if args.keyword:
        req.keyword = args.keyword
    if args.exact_file_name:
        req.exact_file_name = args.exact_file_name
    if args.node_id:
        req.node_id = int(args.node_id)
    if args.owner:
        req.owner = args.owner
    body = call(ctx, "list_files", req)
    files = extract_list(body, "Files")
    rows = [{
        "file_id": g(f, "FileId"), "file_name": g(f, "FileName"), "node_id": g(f, "NodeId"),
        "file_type": g(f, "FileType"), "owner": g(f, "Owner"), "use_type": g(f, "UseType"),
        "file_folder_id": g(f, "FileFolderId"), "file_folder_path": g(f, "AbsoluteFolderPath") or "",
    } for f in files]
    out(ctx, rows, ["file_id", "file_name", "node_id", "file_type", "owner",
                    "file_folder_id", "file_folder_path"])


def cmd_file_folder(ctx, args):
    """查作业/文件所在目录。输入 --file-id / --node-id / --name 之一,输出完整 FolderPath。"""
    pid = need_project(ctx)
    file_id, node_id = args.file_id, args.node_id
    # 名称 → 解析到 node_id/file_id
    if args.name and not file_id and not node_id:
        node = resolve_single_node(ctx, args.name)
        node_id = node.get("node_id")
        file_id = node.get("file_id")
    if not file_id and not node_id:
        die("file folder 需要 --file-id / --node-id / --name 之一。", E_USAGE)

    # 1) GetFile 拿 FileFolderId(同时可兜底补 file_id/node_id)
    greq = dw_models.GetFileRequest(project_id=pid)
    if file_id:
        greq.file_id = int(file_id)
    if node_id:
        greq.node_id = int(node_id)
    gbody = call(ctx, "get_file", greq)
    file_data = g(gbody, "Data", "File", default={}) or {}
    fid = g(file_data, "FileFolderId")
    if not fid:
        die(f"文件(file_id={file_id or ''}, node_id={node_id or ''})未关联文件夹(直接挂在根目录)。", E_NOTFOUND)

    # 2) GetFolder 拿完整目录路径(单次返回即绝对路径)
    fbody = call(ctx, "get_folder", dw_models.GetFolderRequest(project_id=pid, folder_id=fid))
    folder_data = g(fbody, "Data", default={}) or {}
    out(ctx, {
        "file_id": g(file_data, "FileId") or (int(file_id) if file_id else None),
        "node_id": g(file_data, "NodeId") or (int(node_id) if node_id else None),
        "file_name": g(file_data, "FileName"),
        "file_folder_id": fid,
        "folder_id": g(folder_data, "FolderId"),
        "folder_path": g(folder_data, "FolderPath"),
    })


def cmd_file_versions(ctx, args):
    """文件版本历史(不含代码,取代码用 `file version`)。"""
    pid = need_project(ctx)
    req = dw_models.ListFileVersionsRequest(
        project_id=pid, file_id=int(args.file_id),
        page_number=args.page_number, page_size=args.page_size)
    body = call(ctx, "list_file_versions", req)
    rows = [{
        "file_version": g(v, "FileVersion"), "change_type": g(v, "ChangeType"),
        "comment": g(v, "Comment"), "commit_time": ts_str(g(v, "CommitTime")),
        "commit_user": g(v, "CommitUser"), "is_current_prod": g(v, "IsCurrentProd"),
        "status": g(v, "Status"),
    } for v in extract_list(body, "FileVersions")]
    out(ctx, rows, ["file_version", "change_type", "comment", "commit_time", "commit_user", "is_current_prod"])


def cmd_file_version(ctx, args):
    """取指定版本详情/代码。"""
    pid = need_project(ctx)
    req = dw_models.GetFileVersionRequest(
        project_id=pid, file_id=int(args.file_id), file_version=int(args.file_version))
    body = call(ctx, "get_file_version", req)
    v = g(body, "Data", default={}) or {}
    if ctx.fmt == "text":
        print(g(v, "FileContent") or g(v, "Content") or "")
        return
    out(ctx, v)


def cmd_file_get(ctx, args):
    pid = need_project(ctx)
    req = dw_models.GetFileRequest(project_id=pid)
    if args.file_id:
        req.file_id = int(args.file_id)
    if args.node_id:
        req.node_id = int(args.node_id)
    if not args.file_id and not args.node_id:
        die("file get 需要 --file-id 或 --node-id。", E_USAGE)
    body = call(ctx, "get_file", req)
    if ctx.fmt == "text":
        print(g(body, "Data", "File", "Content", default="") or g(body, "Data", "Content", default=""))
        return
    out(ctx, body)


def apply_optional_file_fields(req, args):
    """CreateFile/UpdateFile 共用的可选字段(owner/描述/连接/资源组/调度/重跑/依赖)。"""
    if args.owner:
        req.owner = args.owner
    if args.description:
        req.file_description = args.description
    if args.connection:
        req.connection_name = args.connection
    if args.resource_group:
        req.resource_group_identifier = args.resource_group
    if args.para:
        req.para_value = args.para
    if args.input:
        req.input_list = args.input
    if args.cron:
        req.cron_express = args.cron
    if args.cycle_type:
        req.cycle_type = args.cycle_type
    if args.start_effect:
        req.start_effect_date = to_epoch_ms(args.start_effect)
    if args.end_effect:
        req.end_effect_date = to_epoch_ms(args.end_effect)
    if args.rerun_mode:
        req.rerun_mode = args.rerun_mode
    if args.auto_rerun_times is not None:
        req.auto_rerun_times = args.auto_rerun_times
    if args.auto_rerun_interval is not None:
        req.auto_rerun_interval_millis = args.auto_rerun_interval
    if args.dep_type:
        req.dependent_type = args.dep_type
    if args.dep_nodes:
        req.dependent_node_id_list = args.dep_nodes


def cmd_file_update(ctx, args):
    """更新已有开发节点(代码/调度配置)。通过 --file-id/--node-id/--name 定位。仅改 DEV,不发布。"""
    pid = need_project(ctx)
    file_id, node_id = args.file_id, args.node_id
    if args.name and not file_id and not node_id:
        node = resolve_single_node(ctx, args.name)
        node_id = node.get("node_id")
        file_id = node.get("file_id")
    if not file_id and node_id:
        greq = dw_models.GetFileRequest(project_id=pid, node_id=int(node_id))
        body = call(ctx, "get_file", greq)
        file_id = g(body, "Data", "File", "FileId")
    if not file_id:
        die("file update 需要 --file-id、--node-id 或 --name。", E_USAGE)
    file_id = int(file_id)

    content = args.content
    if args.content_file:
        with open(args.content_file, encoding="utf-8") as f:
            content = f.read()
    if content is None:
        # 未提供代码时读取现有内容,避免更新调度配置时把代码清空
        greq = dw_models.GetFileRequest(project_id=pid, file_id=file_id)
        body = call(ctx, "get_file", greq)
        content = g(body, "Data", "File", "Content")
    if content is None:
        die("无法取得节点代码:请提供 --content-file 或 --content。", E_USAGE)

    req = dw_models.UpdateFileRequest(project_id=pid, file_id=file_id, content=content)
    apply_optional_file_fields(req, args)

    confirm_write(ctx, "UpdateFile", req.to_map())
    body = call(ctx, "update_file", req)
    eprint("[info] 已更新开发环境(DEV);生产未变更,如需上线请继续 `file submit` + `file deploy`。")
    out(ctx, {"file_id": file_id, "node_id": node_id, "result": body})


def cmd_file_create(ctx, args):
    pid = need_project(ctx)
    ftype = resolve_file_type(args.type)
    if ftype is None:
        die(f"未知 --type '{args.type}'。可用别名: {', '.join(sorted(FILE_TYPES))};或直接传数字编码。", E_USAGE)
    content = args.content
    if args.content_file:
        with open(args.content_file, encoding="utf-8") as f:
            content = f.read()
    if content is None:
        die("需要节点代码:用 --content-file <路径> 或 --content <字符串>。", E_USAGE)

    req = dw_models.CreateFileRequest(
        project_id=pid, file_name=args.name, file_type=ftype,
        file_folder_path=args.folder, content=content,
    )
    apply_optional_file_fields(req, args)
    if args.create_folder:
        req.create_folder_if_not_exists = True

    confirm_write(ctx, "CreateFile", req.to_map())
    body = call(ctx, "create_file", req)
    out(ctx, {"file_id": g(body, "Data"), "raw": body})


def cmd_file_submit(ctx, args):
    pid = need_project(ctx)
    req = dw_models.SubmitFileRequest(project_id=pid, file_id=int(args.file_id))
    if args.comment:
        req.comment = args.comment
    confirm_write(ctx, "SubmitFile", req.to_map())
    out(ctx, {"result": call(ctx, "submit_file", req)})


def cmd_file_deploy(ctx, args):
    pid = need_project(ctx)
    req = dw_models.DeployFileRequest(project_id=pid)
    if args.file_id:
        req.file_id = int(args.file_id)
    if args.node_id:
        req.node_id = int(args.node_id)
    if args.comment:
        req.comment = args.comment
    if not args.file_id and not args.node_id:
        die("file deploy 需要 --file-id 或 --node-id。", E_USAGE)
    confirm_write(ctx, "DeployFile", req.to_map())
    out(ctx, {"result": call(ctx, "deploy_file", req)})


def cmd_business_list(ctx, args):
    pid = need_project(ctx)
    req = dw_models.ListBusinessRequest(project_id=pid, page_number=args.page_number, page_size=args.page_size)
    if args.keyword:
        req.keyword = args.keyword
    body = call(ctx, "list_business", req)
    rows = [{
        "business_id": g(b, "BusinessId"), "business_name": g(b, "BusinessName"),
        "description": g(b, "Description"), "owner": g(b, "Owner"), "use_type": g(b, "UseType"),
    } for b in extract_list(body, "Business")]
    out(ctx, rows, ["business_id", "business_name", "description", "owner"])


def cmd_business_get(ctx, args):
    pid = need_project(ctx)
    body = call(ctx, "get_business",
                dw_models.GetBusinessRequest(business_id=int(args.business_id), project_id=pid))
    b = g(body, "Data", default={}) or {}
    out(ctx, {
        "business_id": g(b, "BusinessId"), "business_name": g(b, "BusinessName"),
        "description": g(b, "Description"), "owner": g(b, "Owner"), "use_type": g(b, "UseType"),
    })


def cmd_business_files(ctx, args):
    """某业务流程下的文件:ListFiles 全量分页筛选 BusinessId(ListFiles/ListNodes 均无按 business_id 过滤的 API)。
    文件带 BusinessId,节点不带;故遍历文件匹配。全量遍历较慢,内置 0.3s/页退避防限流。"""
    pid = need_project(ctx)
    biz_id = str(args.business_id)
    page, total, rows = 1, 0, []
    while page <= 100:
        req = dw_models.ListFilesRequest(project_id=pid, page_number=page, page_size=100,
                                         need_absolute_folder_path=True)
        if args.keyword:
            req.keyword = args.keyword
        body = call(ctx, "list_files", req)
        files = extract_list(body, "Files")
        total = g(body, "Data", "TotalCount", default=total) or total
        for f in files:
            if str(g(f, "BusinessId") or "") != biz_id:
                continue
            rows.append({
                "file_id": g(f, "FileId"), "file_name": g(f, "FileName"),
                "node_id": g(f, "NodeId"), "file_type": g(f, "FileType"),
                "owner": g(f, "Owner"), "file_folder_path": g(f, "AbsoluteFolderPath") or "",
            })
        if not files or page * 100 >= total:
            break
        page += 1
        time.sleep(0.3)  # 全量遍历防限流
    if not rows:
        die(f"业务流程 {biz_id} 下未找到文件。", E_NOTFOUND)
    out(ctx, rows, ["file_id", "file_name", "node_id", "file_type", "file_folder_path"])


def _meta_table_parts(s):
    """把 project.table / odps.project.table 解析为 (project, table);缺 project 报错。"""
    s = (s or "").strip()
    if s.startswith("odps."):
        s = s[5:]
    parts = s.split(".")
    if len(parts) == 2:
        return parts[0], parts[1]
    die(f"无法解析表名: {s!r}(支持 project.table 或 odps.project.table,如 lyy_gz.ods_cpa_activity_order_hi)", E_USAGE)


def cmd_meta_table(ctx, args):
    """表结构/元数据(数据地图)。表名用 project.table,如 lyy_gz.ods_cpa_activity_order_hi。"""
    proj, table = _meta_table_parts(args.table)
    guid = f"odps.{proj}.{table}"
    page, cols, total = 1, [], 0
    while page <= 50:
        req = dw_models.GetMetaTableFullInfoRequest(table_guid=guid, page_num=page, page_size=100)
        body = call(ctx, "get_meta_table_full_info", req)
        d = g(body, "Data", default={}) or {}
        total = d.get("TotalColumnCount") or total
        batch = d.get("ColumnList") or []
        for c in batch:
            cols.append({
                "column_name": g(c, "ColumnName"), "column_type": g(c, "ColumnType"),
                "comment": g(c, "Comment"), "is_partition": g(c, "IsPartitionColumn"),
                "primary_key": g(c, "IsPrimaryKey"),
            })
        if not batch or len(cols) >= total:
            break
        page += 1
        time.sleep(0.2)
    out(ctx, {
        "table": f"{proj}.{table}", "table_guid": guid,
        "comment": g(body, "Data", "Comment"), "life_cycle": g(body, "Data", "LifeCycle"),
        "owner": g(body, "Data", "OwnerId"), "project": g(body, "Data", "ProjectName"),
        "create_time": ts_str(g(body, "Data", "CreateTime")), "last_modify_time": ts_str(g(body, "Data", "LastModifyTime")),
        "total_columns": total or len(cols), "columns": cols,
    })


def cmd_meta_lineage(ctx, args):
    """表血缘。direction: up 上游 / down 下游 / all 全部。"""
    proj, table = _meta_table_parts(args.table)
    guid = f"odps.{proj}.{table}"
    direction = (args.direction or "all").upper()
    entities, next_key = [], None
    while True:
        req = dw_models.GetMetaTableLineageRequest(
            table_guid=guid, direction=direction, page_size=100)
        if next_key:
            req.next_primary_key = next_key
        body = call(ctx, "get_meta_table_lineage", req)
        d = g(body, "Data", default={}) or {}
        batch = d.get("DataEntityList") or []
        for e in batch:
            entities.append({
                "table_guid": g(e, "TableGuid"), "table_name": g(e, "TableName"),
                "database": g(e, "DatabaseName"),
            })
        if not d.get("HasNext") or not batch:
            break
        next_key = d.get("NextPrimaryKey") or (batch[-1].get("TableGuid") if batch else None)
        if next_key is None:
            break
        time.sleep(0.2)
    out(ctx, {
        "table": f"{proj}.{table}", "table_guid": guid,
        "direction": direction, "count": len(entities), "entities": entities,
    })


def cmd_resource_list(ctx, args):
    """列出资源组。默认合并类型 1(调度)+2(计算)去重;--type 指定单个。"""
    types = [int(args.type)] if args.type else [1, 2]
    seen, rows = set(), []
    for rt in types:
        body = call(ctx, "list_resource_groups", dw_models.ListResourceGroupsRequest(resource_group_type=rt))
        items = g(body, "Data") or []
        if not isinstance(items, list):
            continue
        for it in items:
            rid = g(it, "Id")
            if rid in seen:
                continue
            seen.add(rid)
            rows.append({
                "id": rid, "name": g(it, "Name"), "identifier": g(it, "Identifier"),
                "type": g(it, "ResourceGroupType"), "mode": g(it, "Mode"),
                "is_default": g(it, "IsDefault"), "status": g(it, "Status"),
                "cluster": g(it, "Cluster"),
            })
    if not rows:
        die("未找到资源组。", E_NOTFOUND)
    out(ctx, rows, ["id", "name", "type", "mode", "is_default", "status", "cluster"])


def cmd_baseline_list(ctx, args):
    pid = need_project(ctx)
    req = dw_models.ListBaselineConfigsRequest(project_id=pid, page_number=args.page_number, page_size=args.page_size)
    if args.search_text:
        req.search_text = args.search_text
    body = call(ctx, "list_baseline_configs", req)
    rows = [{
        "baseline_id": g(b, "BaselineId"), "baseline_name": g(b, "BaselineName"),
        "baseline_type": g(b, "BaselineType"),
        "sla": f"{g(b, 'SlaHour')}:{str(g(b, 'SlaMinu') or 0).zfill(2)}",
        "exp": f"{g(b, 'ExpHour')}:{str(g(b, 'ExpMinu') or 0).zfill(2)}",
        "owner": g(b, "Owner"), "priority": g(b, "Priority"), "use_flag": g(b, "UseFlag"),
    } for b in extract_list(body, "Baselines")]
    out(ctx, rows, ["baseline_id", "baseline_name", "baseline_type", "sla", "exp", "owner", "priority", "use_flag"])


def cmd_baseline_status(ctx, args):
    """某天各基线保障状态。ListBaselineStatuses 的 bizdate 须为 yyyy-MM-ddTHH:mm:ss+0800(RFC822 时区,网关正则要求)。"""
    if not args.biz_date or not re.match(r"^\d{4}-\d{2}-\d{2}$", str(args.biz_date)):
        die("baseline status 需要 --biz-date YYYY-MM-DD。", E_USAGE)
    bizdate = f"{args.biz_date}T00:00:00+0800"  # +0800=中国时区
    req = dw_models.ListBaselineStatusesRequest(
        bizdate=bizdate, page_number=args.page_number, page_size=args.page_size)
    body = call(ctx, "list_baseline_statuses", req)
    rows = [{
        "baseline_id": g(b, "BaselineId"), "baseline_name": g(b, "BaselineName"),
        "status": g(b, "Status"), "finish_status": g(b, "FinishStatus"),
        "buffer_sec": g(b, "Buffer"), "sla_time": ts_str(g(b, "SlaTime")),
        "finish_time": ts_str(g(b, "FinishTime")), "priority": g(b, "Priority"),
    } for b in extract_list(body, "BaselineStatuses")]
    out(ctx, {"biz_date": args.biz_date, "count": len(rows), "baselines": rows})


def _dqc_list(body, keys):
    """DQC 列表响应信封不统一:Data 可能直接是 list,也可能嵌套在指定 key 下。"""
    d = g(body, "Data") or {}
    if isinstance(d, list):
        return d
    if isinstance(d, dict):
        for k in keys:
            v = d.get(k) or d.get(k.lower()) or d.get(k.title())
            if isinstance(v, list):
                return v
    return []


def cmd_quality_entity(ctx, args):
    """查表的质量实体(DQC)。表名用 project.table。"""
    pid = need_project(ctx)
    proj, table = _meta_table_parts(args.table)
    body = call(ctx, "get_quality_entity", dw_models.GetQualityEntityRequest(
        project_id=pid, project_name=proj, table_name=table, env_type=ctx.env))
    rows = [{
        "entity_id": g(e, "EntityId"), "table_name": g(e, "TableName"),
        "match_expression": g(e, "MatchExpression"), "on_duty": g(e, "OnDuty"),
        "env_type": g(e, "EnvType"),
    } for e in _dqc_list(body, ("EntityList", "Entities"))]
    out(ctx, rows, ["entity_id", "table_name", "match_expression", "on_duty", "env_type"])


def cmd_quality_rules(ctx, args):
    """查实体的质量规则。需 --entity-id(先用 `quality entity` 拿)。"""
    pid = need_project(ctx)
    if not args.entity_id:
        die("quality rules 需要 --entity-id(先用 `quality entity --table project.table` 查)。", E_USAGE)
    req = dw_models.ListQualityRulesRequest(
        entity_id=int(args.entity_id), project_id=pid, project_name=args.project_name or "",
        page_number=1, page_size=args.page_size)
    body = call(ctx, "list_quality_rules", req)
    rows = [{
        "rule_id": g(r, "RuleId"), "rule_name": g(r, "RuleName"),
        "checker_type": g(r, "CheckerType") or g(r, "CheckerName"),
        "operator": g(r, "Operator"), "threshold": g(r, "Threshold"),
        "match_expression": g(r, "MatchExpression"), "expect_value": g(r, "ExpectValue"),
        "block_type": g(r, "BlockType"), "template_id": g(r, "TemplateId"),
    } for r in _dqc_list(body, ("Rules", "RuleList", "Data"))]
    out(ctx, rows, ["rule_id", "rule_name", "checker_type", "operator", "threshold", "expect_value"])


def cmd_quality_results(ctx, args):
    """查质量校验结果。--rule-id 或 --entity-id 二选一。日期 YYYY-MM-DD。"""
    pid = need_project(ctx)
    pname = args.project_name or ""
    if not args.rule_id and not args.entity_id:
        die("quality results 需要 --rule-id 或 --entity-id。", E_USAGE)
    if args.rule_id:
        req = dw_models.ListQualityResultsByRuleRequest(
            rule_id=int(args.rule_id), project_id=pid, project_name=pname,
            page_number=args.page_number, page_size=args.page_size)
        if args.start_date:
            req.start_date = args.start_date
        if args.end_date:
            req.end_date = args.end_date
        body = call(ctx, "list_quality_results_by_rule", req)
        items = _dqc_list(body, ("RuleCheckResultList", "Results", "Data"))
    else:
        req = dw_models.ListQualityResultsByEntityRequest(
            entity_id=int(args.entity_id), project_id=pid, project_name=pname,
            page_number=args.page_number, page_size=args.page_size)
        if args.start_date:
            req.start_date = args.start_date
        if args.end_date:
            req.end_date = args.end_date
        body = call(ctx, "list_quality_results_by_entity", req)
        items = _dqc_list(body, ("EntityCheckResultList", "Results", "Data"))
    rows = [{
        "rule_id": g(r, "RuleId"), "rule_name": g(r, "RuleName"),
        "check_time": ts_str(g(r, "CheckTime")), "check_status": g(r, "CheckStatus") or g(r, "Status"),
        "actual_value": g(r, "ActualValue") or g(r, "ActualExpression"),
        "expect_value": g(r, "ExpectValue"), "operator": g(r, "Operator"),
    } for r in items]
    out(ctx, {"count": len(rows), "results": rows})


def cmd_instance_list(ctx, args):
    node_id = args.node_id
    if args.task_name and not node_id:
        node = resolve_single_node(ctx, args.task_name)
        node_id = node["node_id"]
    status = args.status
    if args.failed:
        status = "FAILURE"
    if not (args.biz_date or args.begin_bizdate or args.end_bizdate or status or args.dag_id
            or node_id or args.node_name):
        eprint("[warn] 未限定业务日期/状态/节点,将查询全量历史实例;建议加 --biz-date 或 --status")
    rows = find_instances(
        ctx, node_id=node_id, node_name=args.node_name, status=status, dag_id=args.dag_id,
        begin_bizdate=to_biz_datetime(args.begin_bizdate) if args.begin_bizdate else (
            to_biz_datetime(args.biz_date) if args.biz_date else None),
        end_bizdate=to_biz_datetime(args.end_bizdate, end_of_day=True) if args.end_bizdate else (
            to_biz_datetime(args.biz_date, end_of_day=True) if args.biz_date else None),
        page_size=args.page_size,
    )
    if ctx.fmt == "table":
        rows = [
            {**r, "bizdate": ts_str(r.get("bizdate")),
             "begin_running_time": ts_str(r.get("begin_running_time")),
             "finish_time": ts_str(r.get("finish_time"))}
            for r in rows
        ]
    out(ctx, rows, ["instance_id", "node_id", "node_name", "status", "bizdate",
                    "begin_running_time", "finish_time", "owner"])


def cmd_instance_get(ctx, args):
    req = dw_models.GetInstanceRequest(instance_id=int(args.instance_id), project_env=ctx.env)
    out(ctx, call(ctx, "get_instance", req))


def cmd_instance_log(ctx, args):
    req = dw_models.GetInstanceLogRequest(instance_id=int(args.instance_id), project_env=ctx.env)
    body = call(ctx, "get_instance_log", req)
    log = g(body, "Data", default="")
    if not isinstance(log, str):
        log = g(body, "Data", "Log", default="") or json.dumps(log, ensure_ascii=False, default=str)
    lines = str(log).splitlines()
    if args.grep:
        lines = [l for l in lines if args.grep.lower() in l.lower()]
    if args.lines:
        lines = lines[-args.lines:]
    filtered = "\n".join(lines)
    if ctx.fmt == "text":
        print(filtered)
    else:
        out(ctx, {"instance_id": int(args.instance_id), "matched_lines": len(lines), "log": filtered})


def _summarize_status(s):
    """DataWorks 实例状态 → 汇总桶(参考 reference/pitfalls.md 的状态口径)。"""
    s = str(s or "").upper()
    if "SUCCESS" in s:
        return "success"
    if "FAIL" in s:
        return "failure"
    if "RUN" in s or "CHECK" in s:
        return "running"
    if "WAIT" in s:
        return "waiting"
    if "NOT_RUN" in s or "NOTRUN" in s:
        return "not_run"
    if "SKIP" in s or "PAUSED" in s or "DONE" in s:
        return "skipped"
    return "other"


def cmd_instance_stat(ctx, args):
    """按业务日期统计实例:状态分布 + 失败 Top 节点(--biz-date 必填,限定范围)。"""
    pid = need_project(ctx)
    begin = to_biz_datetime(args.begin_bizdate or args.biz_date)
    end = to_biz_datetime(args.end_bizdate or args.biz_date, end_of_day=True)
    if not begin:
        die("instance stat 需要 --biz-date YYYY-MM-DD(或 --begin-bizdate)。", E_USAGE)
    buckets = collections.Counter()
    by_status = collections.Counter()
    failed_by_node = {}
    page, total = 1, 0
    while page <= 100:  # 分页遍历,ListInstances PageSize 上限 100
        req = dw_models.ListInstancesRequest(
            project_id=pid, project_env=ctx.env,
            begin_bizdate=begin, end_bizdate=end,
            page_number=page, page_size=100,
        )
        body = call(ctx, "list_instances", req)
        insts = extract_list(body, "Instances")
        total = g(body, "Data", "TotalCount", default=total) or total
        for it in insts:
            raw = str(g(it, "Status") or "").upper() or "UNKNOWN"
            by_status[raw] += 1
            buckets[_summarize_status(raw)] += 1
            if "FAIL" in raw:
                nid = g(it, "NodeId")
                nm = g(it, "NodeName") or "?"
                f = failed_by_node.setdefault(nid, {"node_id": nid, "node_name": nm, "count": 0})
                f["count"] += 1
        if not insts or sum(by_status.values()) >= total:
            break
        page += 1
    total_found = sum(by_status.values())
    failed_top = sorted(failed_by_node.values(), key=lambda x: x["count"], reverse=True)[:args.top]
    out(ctx, {
        "biz_date": args.biz_date or args.begin_bizdate or "",
        "range": {"begin": begin, "end": end},
        "total": total or total_found,
        "summary": dict(buckets),
        "by_status": dict(by_status),
        "failed_top": failed_top,
    })


def _write_instance_op(ctx, args, status_target, action, model_name, method):
    instance_id = args.instance_id
    if args.task_name and not instance_id:
        inst, node = resolve_instance_for_task(ctx, args.task_name, status_target)
        instance_id = inst["instance_id"]
    if not instance_id:
        die(f"{action} 需要 --instance-id 或 --task-name。", E_USAGE)
    req_cls = getattr(dw_models, model_name)
    req = req_cls(instance_id=int(instance_id), project_env=ctx.env)
    confirm_write(ctx, action, req.to_map())
    out(ctx, {"result": call(ctx, method, req), "instance_id": int(instance_id)})


def cmd_instance_restart(ctx, args):
    _write_instance_op(ctx, args, "FAILURE", "RestartInstance", "RestartInstanceRequest", "restart_instance")


def cmd_instance_stop(ctx, args):
    _write_instance_op(ctx, args, None, "StopInstance", "StopInstanceRequest", "stop_instance")


def cmd_instance_set_success(ctx, args):
    _write_instance_op(ctx, args, "FAILURE", "SetSuccessInstance", "SetSuccessInstanceRequest", "set_success_instance")


def norm_time_range(begin_t, end_t):
    """小时级补数时间范围 → (begin, end)。格式 HH:mm:ss;也兼容 yyyy-MM-dd HH:mm:ss(自动取时间部分)。"""
    outs = []
    for v in (begin_t, end_t):
        if not re.match(r"^\d{2}:\d{2}:\d{2}$", v):
            m2 = re.match(r"^\d{4}-\d{2}-\d{2}[T ](\d{2}:\d{2}:\d{2})$", v)
            if not m2:
                die(f"时间格式需为 HH:mm:ss(收到: {v})。", E_USAGE)
            v = m2.group(1)
        outs.append(v)
    begin, end = outs
    if begin >= end:
        die(f"--begin-time({begin}) 需早于 --end-time({end}),否则时间区间为空、无法匹配任何实例。", E_USAGE)
    return begin, end


def cmd_complement_run(ctx, args):
    # 解析目标节点
    if args.task_name:
        node = resolve_single_node(ctx, args.task_name)
        node_id = node["node_id"]
        root = node_id
        include = args.include_node_ids or str(node_id)
    elif args.root_node_id:
        root = int(args.root_node_id)
        include = args.include_node_ids or str(root)
        node_id = root
    else:
        die("complement run 需要 --task-name 或 --root-node-id。", E_USAGE)

    start = end = None
    begin_time = end_time = None

    # 语义补数:--data-date + --hour = 想检查"某日某小时"的数据分区
    if args.data_date or args.hour is not None:
        if args.start_biz:
            die("--data-date/--hour 与 --start-biz 互斥,请选一种语义。", E_USAGE)
        if not args.data_date or args.hour is None:
            die("--data-date 与 --hour 需同时提供。", E_USAGE)
        if not re.match(r"^\d{4}-\d{2}-\d{2}$", str(args.data_date)):
            die(f"--data-date 需为 YYYY-MM-DD: {args.data_date}", E_USAGE)
        if not 0 <= args.hour <= 23:
            die(f"--hour 需为 0-23: {args.hour}", E_USAGE)
        dd = datetime.datetime.strptime(args.data_date, "%Y-%m-%d")
        # 该任务规律:实例 cyc = bizdate + 1,dt = cyc - 1 小时;
        # 要检查数据分区 dt=YYYYMMDDHH → 调度 (H+1):05 的实例 → bizdate = 数据日期 - 1。
        biz = dd - datetime.timedelta(days=1)
        check_h = args.hour + 1
        start = end = biz.strftime("%Y-%m-%d 00:00:00")
        begin_time = f"{check_h:02d}:00:00"
        end_time = f"{check_h:02d}:59:59"
        eprint(
            f"[info] 换算(适用\"每小时 HH:05 调度、检查前一小时分区\"的任务): "
            f"数据分区 dt={args.data_date.replace('-', '')}{args.hour:02d} → "
            f"业务日期 {biz.strftime('%Y-%m-%d')} + 时间 {begin_time}~{end_time}(调度 {check_h:02d}:05)"
        )
    else:
        start = norm_biz_datetime(args.start_biz)
        end = norm_biz_datetime(args.end_biz) or start
        if not start:
            die("complement run 需要 --start-biz(YYYY-MM-DD)或 --data-date+--hour。", E_USAGE)
        if args.begin_time or args.end_time:
            if not (args.begin_time and args.end_time):
                die("--begin-time 与 --end-time 需同时提供。", E_USAGE)
            begin_time, end_time = norm_time_range(args.begin_time, args.end_time)

    name = args.name or f"dwcli-backfill-{node_id}-{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}"

    req = dw_models.RunCycleDagNodesRequest(
        project_env=ctx.env, name=name, root_node_id=int(root),
        include_node_ids=str(include), start_biz_date=start, end_biz_date=end,
    )
    if args.exclude_node_ids:
        req.exclude_node_ids = args.exclude_node_ids
    if begin_time:
        req.biz_begin_time = begin_time
        req.biz_end_time = end_time
    # RunCycleDagNodes 的 parallelism 为 bool 且必填(开启并行补数);见实测
    req.parallelism = True
    if args.node_params:
        req.node_params = args.node_params

    confirm_write(ctx, "RunCycleDagNodes", req.to_map())
    body = call(ctx, "run_cycle_dag_nodes", req)
    dag_id = g(body, "Data")
    eprint(f"[info] 补数已发起(dag_id={dag_id}),跟踪进度: complement status --dag-id {dag_id}")
    out(ctx, {"dag_id": dag_id, "name": name, "raw": body})


def cmd_complement_status(ctx, args):
    dag = call(ctx, "get_dag", dw_models.GetDagRequest(dag_id=int(args.dag_id), project_env=ctx.env))
    dag_data = g(dag, "Data", default={}) or {}
    pid = need_project(ctx)
    counts = {"total": 0, "success": 0, "failure": 0, "running": 0, "waiting": 0, "not_run": 0, "other": 0}
    page, total = 1, 0
    while page <= 100:  # 分页拉取该 DAG 的实例(ListInstances PageSize 上限 100)
        req = dw_models.ListInstancesRequest(project_id=pid, project_env=ctx.env,
                                             dag_id=int(args.dag_id), page_number=page, page_size=100)
        body = call(ctx, "list_instances", req)
        insts = extract_list(body, "Instances")
        total = g(body, "Data", "TotalCount", default=total) or total
        for it in insts:
            s = str(g(it, "Status") or "").upper()
            counts["total"] += 1
            if "SUCCESS" in s:
                counts["success"] += 1
            elif "FAIL" in s:
                counts["failure"] += 1
            elif "RUN" in s or "CHECK" in s:
                counts["running"] += 1
            elif "WAIT" in s:
                counts["waiting"] += 1
            elif "NOT_RUN" in s or "NOTRUN" in s:
                counts["not_run"] += 1
            else:
                counts["other"] += 1
        if not insts or counts["total"] >= total:
            break
        page += 1
    counts["total"] = total or counts["total"]
    done = counts["success"] + counts["failure"]
    progress = round(done / counts["total"] * 100, 1) if counts["total"] else None
    out(ctx, {
        "dag_id": int(args.dag_id),
        "dag_status": g(dag_data, "Status") or g(dag_data, "DagStatus"),
        "dag_type": g(dag_data, "Type") or g(dag_data, "DagType"),
        "instances": counts,
        "progress_pct": progress,
        "dag_raw": dag_data,
    })


# =========================================================================
# 参数解析
# =========================================================================
def add_global_flags(parser, suppress=False):
    """全局参数。suppress=True 时默认值用 SUPPRESS,挂在子命令上也不会覆盖顶层已给的值。"""
    sd = argparse.SUPPRESS
    d = (lambda real: sd if suppress else real)
    parser.add_argument("--access-key-id", default=d(None), help="默认取 DATAWORKS_ACCESS_KEY_ID / ALIBABA_CLOUD_ACCESS_KEY_ID")
    parser.add_argument("--access-key-secret", default=d(None), help="默认取 DATAWORKS_ACCESS_KEY_SECRET / ALIBABA_CLOUD_ACCESS_KEY_SECRET")
    parser.add_argument("-r", "--region", default=d(None), help="默认 DATAWORKS_REGION_ID / ALIBABA_CLOUD_REGION_ID,否则 cn-shanghai")
    parser.add_argument("--endpoint", default=d(None), help="默认 dataworks.<region>.aliyuncs.com")
    parser.add_argument("--project-id", default=d(None), help="DataWorks 工作空间数字 ID(默认 DATAWORKS_PROJECT_ID)")
    parser.add_argument("--env", default=d(None), help="ProjectEnv,默认 PROD(可选 DEV)")
    parser.add_argument("--format", choices=["json", "table", "text"], default=d(None), help="默认 json")
    parser.add_argument("--timeout", type=int, default=d(30), help="超时秒数,默认 30")
    parser.add_argument("--debug", action="store_true", default=d(False), help="打印解析后的请求到 stderr")
    parser.add_argument("--dry-run", action="store_true", default=d(False), help="写操作仅打印请求、不调用")
    parser.add_argument("--yes", action="store_true", default=d(False), help="确认执行写操作")


def build_parser():
    p = argparse.ArgumentParser(prog="dw", description="DataWorks 2020-05-18 CLI(供 dataworks skill 调用)")
    add_global_flags(p, suppress=False)
    # 同名全局参数挂到每个子命令(SUPPRESS 默认):放子命令前后都能识别,且不覆盖顶层值
    parent = argparse.ArgumentParser(add_help=False)
    add_global_flags(parent, suppress=True)
    sub = p.add_subparsers(dest="cmd", required=True)

    # doctor
    sp = sub.add_parser("doctor", parents=[parent], help="自检:SDK/凭证/region/endpoint(加 --check 测连通)")
    sp.add_argument("--check", action="store_true")
    sp.set_defaults(func=cmd_doctor)

    # project
    proj = sub.add_parser("project", help="工作空间")
    projs = proj.add_subparsers(dest="sub", required=True)
    pl = projs.add_parser("list", parents=[parent], help="列出工作空间")
    pl.add_argument("--page-number", type=int, default=1)
    pl.add_argument("--page-size", type=int, default=50)
    pl.set_defaults(func=cmd_project_list)

    # node
    node = sub.add_parser("node", help="调度节点")
    nodes = node.add_subparsers(dest="sub", required=True)
    nl = nodes.add_parser("list", parents=[parent], help="节点列表")
    nl.add_argument("--name"); nl.add_argument("--owner"); nl.add_argument("--program-type")
    nl.add_argument("--page-number", type=int, default=1); nl.add_argument("--page-size", type=int, default=50)
    nl.set_defaults(func=cmd_node_list)
    ng = nodes.add_parser("get", parents=[parent], help="节点详情")
    ng.add_argument("--node-id", required=True)
    ng.set_defaults(func=cmd_node_get)
    nr = nodes.add_parser("resolve", parents=[parent], help="名称→nodeId/fileId")
    nr.add_argument("--name", required=True); nr.add_argument("--exact", action="store_true")
    nr.add_argument("--limit", type=int, default=50)
    nr.set_defaults(func=cmd_node_resolve)
    np = nodes.add_parser("parents", parents=[parent], help="上游(父)节点")
    np.add_argument("--node-id"); np.add_argument("--task-name", help="按任务名解析")
    np.set_defaults(func=cmd_node_parents)
    nc = nodes.add_parser("children", parents=[parent], help="下游(子)节点")
    nc.add_argument("--node-id"); nc.add_argument("--task-name", help="按任务名解析")
    nc.set_defaults(func=cmd_node_children)

    # file
    filep = sub.add_parser("file", help="文件/节点代码")
    files = filep.add_subparsers(dest="sub", required=True)
    fl = files.add_parser("list", parents=[parent], help="文件列表")
    fl.add_argument("--keyword"); fl.add_argument("--exact-file-name"); fl.add_argument("--node-id")
    fl.add_argument("--owner"); fl.add_argument("--page-number", type=int, default=1)
    fl.add_argument("--page-size", type=int, default=50)
    fl.set_defaults(func=cmd_file_list)
    ff = files.add_parser("folder", parents=[parent], help="查作业所在目录(FolderPath)")
    ff.add_argument("--file-id"); ff.add_argument("--node-id"); ff.add_argument("--name", help="按任务名解析")
    ff.set_defaults(func=cmd_file_folder)
    fg = files.add_parser("get", parents=[parent], help="取文件/节点代码(--format text 仅输出代码)")
    fg.add_argument("--file-id"); fg.add_argument("--node-id")
    fg.set_defaults(func=cmd_file_get)
    fv = files.add_parser("versions", parents=[parent], help="文件版本历史")
    fv.add_argument("--file-id", required=True)
    fv.add_argument("--page-number", type=int, default=1); fv.add_argument("--page-size", type=int, default=20)
    fv.set_defaults(func=cmd_file_versions)
    fver = files.add_parser("version", parents=[parent], help="取某版本详情/代码(--format text 仅代码)")
    fver.add_argument("--file-id", required=True); fver.add_argument("--file-version", required=True, type=int)
    fver.set_defaults(func=cmd_file_version)
    fc = files.add_parser("create", parents=[parent], help="新建开发节点(写,需 --yes)")
    fc.add_argument("--name", required=True)
    fc.add_argument("--type", required=True, help=f"节点类型,别名: {', '.join(sorted(FILE_TYPES))} 或数字编码")
    fc.add_argument("--folder", help="FileFolderPath,如 业务流程/xx/MaxCompute/数据开发")
    fc.add_argument("--content-file", help="节点代码文件路径")
    fc.add_argument("--content", help="节点代码字符串(与 --content-file 二选一)")
    fc.add_argument("--owner"); fc.add_argument("--description"); fc.add_argument("--connection")
    fc.add_argument("--resource-group"); fc.add_argument("--para", help="调度参数,如 bizdate=$[yyyymmdd-1]")
    fc.add_argument("--input", help="输入依赖 InputList,逗号分隔")
    fc.add_argument("--cron", help="定时 CronExpress,如 '00 05 00 * * ?'")
    fc.add_argument("--cycle-type", help="DAY / NOT_DAY")
    fc.add_argument("--start-effect", help="调度生效起(YYYY-MM-DD 或 epoch 毫秒)")
    fc.add_argument("--end-effect", help="调度生效止")
    fc.add_argument("--rerun-mode", help="ALL_ALLOWED / FAILURE_ALLOWED / ALL_DENIED")
    fc.add_argument("--auto-rerun-times", type=int, help="自动重跑次数")
    fc.add_argument("--auto-rerun-interval", type=int, help="自动重跑间隔(毫秒)")
    fc.add_argument("--dep-type", help="依赖类型")
    fc.add_argument("--dep-nodes", help="依赖节点 ID 列表")
    fc.add_argument("--create-folder", action="store_true", help="目录不存在时自动创建")
    fc.set_defaults(func=cmd_file_create)
    fu = files.add_parser("update", parents=[parent], help="更新开发节点代码/调度配置(写,需 --yes)")
    fu.add_argument("--file-id"); fu.add_argument("--node-id"); fu.add_argument("--name", help="按任务名解析定位")
    fu.add_argument("--content-file", help="节点代码文件路径(不传则保留现有代码)")
    fu.add_argument("--content", help="节点代码字符串(与 --content-file 二选一;不传则保留现有代码)")
    fu.add_argument("--owner"); fu.add_argument("--description"); fu.add_argument("--connection")
    fu.add_argument("--resource-group"); fu.add_argument("--para", help="调度参数,如 bizdate=$[yyyymmdd-1]")
    fu.add_argument("--input", help="输入依赖 InputList,逗号分隔")
    fu.add_argument("--cron", help="定时 CronExpress,如 '00 05 00 * * ?'")
    fu.add_argument("--cycle-type", help="DAY / NOT_DAY")
    fu.add_argument("--start-effect", help="调度生效起(YYYY-MM-DD 或 epoch 毫秒)")
    fu.add_argument("--end-effect", help="调度生效止")
    fu.add_argument("--rerun-mode", help="ALL_ALLOWED / FAILURE_ALLOWED / ALL_DENIED")
    fu.add_argument("--auto-rerun-times", type=int, help="自动重跑次数")
    fu.add_argument("--auto-rerun-interval", type=int, help="自动重跑间隔(毫秒)")
    fu.add_argument("--dep-type", help="依赖类型")
    fu.add_argument("--dep-nodes", help="依赖节点 ID 列表")
    fu.set_defaults(func=cmd_file_update)
    fs = files.add_parser("submit", parents=[parent], help="提交文件(写,需 --yes)")
    fs.add_argument("--file-id", required=True); fs.add_argument("--comment")
    fs.set_defaults(func=cmd_file_submit)
    fd = files.add_parser("deploy", parents=[parent], help="发布文件到生产(写,需 --yes)")
    fd.add_argument("--file-id"); fd.add_argument("--node-id"); fd.add_argument("--comment")
    fd.set_defaults(func=cmd_file_deploy)

    # business 业务流程
    biz = sub.add_parser("business", help="业务流程")
    bizs = biz.add_subparsers(dest="sub", required=True)
    bl = bizs.add_parser("list", parents=[parent], help="列出业务流程")
    bl.add_argument("--keyword"); bl.add_argument("--page-number", type=int, default=1)
    bl.add_argument("--page-size", type=int, default=50)
    bl.set_defaults(func=cmd_business_list)
    bg = bizs.add_parser("get", parents=[parent], help="业务流程详情")
    bg.add_argument("--business-id", required=True)
    bg.set_defaults(func=cmd_business_get)
    bf = bizs.add_parser("files", parents=[parent], help="某业务流程下的文件(全量遍历匹配,较慢)")
    bf.add_argument("--business-id", required=True); bf.add_argument("--keyword")
    bf.set_defaults(func=cmd_business_files)

    # meta 表元数据/血缘
    meta = sub.add_parser("meta", help="表元数据/血缘(数据地图)")
    metas = meta.add_subparsers(dest="sub", required=True)
    mt = metas.add_parser("table", parents=[parent], help="表结构/元数据(表名用 project.table)")
    mt.add_argument("--table", required=True, help="如 lyy_gz.ods_cpa_activity_order_hi")
    mt.set_defaults(func=cmd_meta_table)
    ml = metas.add_parser("lineage", parents=[parent], help="表血缘(上游/下游)")
    ml.add_argument("--table", required=True, help="如 lyy_gz.ods_cpa_activity_order_hi")
    ml.add_argument("--direction", choices=["up", "down", "all"], default="all", help="up 上游 / down 下游 / all 全部")
    ml.set_defaults(func=cmd_meta_lineage)

    # resource 资源组
    res = sub.add_parser("resource", help="资源组")
    ress = res.add_subparsers(dest="sub", required=True)
    rl = ress.add_parser("list", parents=[parent], help="列出资源组(默认合并调度+计算)")
    rl.add_argument("--type", type=int, help="仅指定类型:1=调度,2=计算")
    rl.set_defaults(func=cmd_resource_list)

    # baseline 基线保障
    bln = sub.add_parser("baseline", help="基线保障")
    blns = bln.add_subparsers(dest="sub", required=True)
    bll = blns.add_parser("list", parents=[parent], help="列出基线配置")
    bll.add_argument("--search-text")
    bll.add_argument("--page-number", type=int, default=1); bll.add_argument("--page-size", type=int, default=50)
    bll.set_defaults(func=cmd_baseline_list)
    bls = blns.add_parser("status", parents=[parent], help="某天各基线保障状态(--biz-date 必填)")
    bls.add_argument("--biz-date", required=True, help="业务日期 YYYY-MM-DD")
    bls.add_argument("--page-number", type=int, default=1); bls.add_argument("--page-size", type=int, default=50)
    bls.set_defaults(func=cmd_baseline_status)

    # quality 数据质量 DQC(只读)
    qual = sub.add_parser("quality", help="数据质量 DQC(只读)")
    quals = qual.add_subparsers(dest="sub", required=True)
    qe = quals.add_parser("entity", parents=[parent], help="查表的质量实体")
    qe.add_argument("--table", required=True, help="project.table,如 lyy_gz.ods_cpa_activity_order_hi")
    qe.set_defaults(func=cmd_quality_entity)
    qr = quals.add_parser("rules", parents=[parent], help="查实体的质量规则")
    qr.add_argument("--entity-id", required=True, help="质量实体 ID(先 quality entity 查)")
    qr.add_argument("--project-name", help="MaxCompute 项目名,如 lyy_gz")
    qr.add_argument("--page-size", type=int, default=50)
    qr.set_defaults(func=cmd_quality_rules)
    qres = quals.add_parser("results", parents=[parent], help="查质量校验结果")
    qres.add_argument("--rule-id", type=int, help="规则 ID")
    qres.add_argument("--entity-id", type=int, help="实体 ID(与 --rule-id 二选一)")
    qres.add_argument("--project-name", help="MaxCompute 项目名,如 lyy_gz")
    qres.add_argument("--start-date", help="起始日期 YYYY-MM-DD")
    qres.add_argument("--end-date", help="结束日期 YYYY-MM-DD")
    qres.add_argument("--page-number", type=int, default=1); qres.add_argument("--page-size", type=int, default=50)
    qres.set_defaults(func=cmd_quality_results)

    # instance
    inst = sub.add_parser("instance", help="实例运维")
    insts = inst.add_subparsers(dest="sub", required=True)
    il = insts.add_parser("list", parents=[parent], help="实例列表")
    il.add_argument("--node-id"); il.add_argument("--node-name"); il.add_argument("--task-name", help="按任务名解析后过滤")
    il.add_argument("--dag-id"); il.add_argument("--status"); il.add_argument("--failed", action="store_true")
    il.add_argument("--biz-date", help="业务日期 YYYY-MM-DD(=begin=end)")
    il.add_argument("--begin-bizdate"); il.add_argument("--end-bizdate")
    il.add_argument("--page-size", type=int, default=100)
    il.set_defaults(func=cmd_instance_list)
    istat = insts.add_parser("stat", parents=[parent], help="按业务日期统计实例状态分布与失败Top(--biz-date 必填)")
    istat.add_argument("--biz-date", help="业务日期 YYYY-MM-DD(=begin=end)")
    istat.add_argument("--begin-bizdate"); istat.add_argument("--end-bizdate")
    istat.add_argument("--top", type=int, default=10, help="失败Top节点数,默认10")
    istat.set_defaults(func=cmd_instance_stat)
    ig = insts.add_parser("get", parents=[parent], help="实例详情")
    ig.add_argument("--instance-id", required=True)
    ig.set_defaults(func=cmd_instance_get)
    ilg = insts.add_parser("log", parents=[parent], help="实例运行日志")
    ilg.add_argument("--instance-id", required=True)
    ilg.add_argument("--lines", type=int, help="仅取尾部 N 行")
    ilg.add_argument("--grep", help="按关键字过滤行")
    ilg.set_defaults(func=cmd_instance_log)
    ir = insts.add_parser("restart", parents=[parent], help="重跑实例(写,需 --yes)")
    ir.add_argument("--instance-id"); ir.add_argument("--task-name")
    ir.set_defaults(func=cmd_instance_restart)
    ist = insts.add_parser("stop", parents=[parent], help="停止实例(写,需 --yes)")
    ist.add_argument("--instance-id"); ist.add_argument("--task-name")
    ist.set_defaults(func=cmd_instance_stop)
    iss = insts.add_parser("set-success", parents=[parent], help="置成功(写,需 --yes)")
    iss.add_argument("--instance-id"); iss.add_argument("--task-name")
    iss.set_defaults(func=cmd_instance_set_success)

    # complement
    comp = sub.add_parser("complement", help="补数据")
    comps = comp.add_subparsers(dest="sub", required=True)
    cr = comps.add_parser("run", parents=[parent], help="发起补数据(写,需 --yes)")
    cr.add_argument("--task-name", help="按任务名解析节点")
    cr.add_argument("--root-node-id"); cr.add_argument("--include-node-ids"); cr.add_argument("--exclude-node-ids")
    cr.add_argument("--start-biz", help="起始业务日期 YYYY-MM-DD(--data-date 未给时必填)")
    cr.add_argument("--end-biz", help="结束业务日期 YYYY-MM-DD(默认=start)")
    cr.add_argument("--name", help="补数据任务名")
    cr.add_argument("--begin-time", help="小时级补数:业务时间起 HH:mm:ss(如 12:00:00)")
    cr.add_argument("--end-time", help="小时级补数:业务时间止 HH:mm:ss(如 12:59:59)")
    cr.add_argument("--data-date", help="语义补数:要检查的数据日期 YYYY-MM-DD(自动换算业务日期=前一天)")
    cr.add_argument("--hour", type=int, help="语义补数:要检查的数据小时 HH(0-23),配合 --data-date")
    cr.add_argument("--node-params", help="节点参数 JSON 串")
    cr.set_defaults(func=cmd_complement_run)
    cs = comps.add_parser("status", parents=[parent], help="补数据 DAG 进度")
    cs.add_argument("--dag-id", required=True)
    cs.set_defaults(func=cmd_complement_status)

    return p


def main(argv=None):
    parser = build_parser()
    args = parser.parse_args(argv)
    ctx = build_ctx(args)
    try:
        args.func(ctx, args)
    except SystemExit:
        raise
    except KeyboardInterrupt:
        die("已中断。", E_ERR)
    except Exception as e:  # noqa: BLE001
        handle_api_error(e)


if __name__ == "__main__":
    main()
