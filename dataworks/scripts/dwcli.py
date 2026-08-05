#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""dwcli — 阿里云 DataWorks(2020-05-18,公共云)命令行工具。

供 Claude Code 的 dataworks skill 调用。设计原则:
  * JSON 优先输出(便于 Claude 解析与串联),--format table/text 供人查看
  * 写操作(restart/stop/set-success/complement run/file create/submit/deploy)
    需 --yes 确认;--dry-run 仅打印将执行的请求、不真正调用
  * 守护式 SDK import:缺依赖时给出安装指引并以退出码 6 退出
  * 凭证默认复用 odps MCP 的环境变量(ODPS_ACCESS_ID/ODPS_ACCESS_KEY)

退出码:
  0 成功 · 1 API/一般错误 · 2 用法错误或缺 --yes · 3 鉴权
  4 未找到/零匹配 · 5 无权限 · 6 SDK 未安装 · 7 超时
"""
import argparse
import datetime
import json
import os
import re
import sys
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
        "DATAWORKS_ACCESS_KEY_ID", "ODPS_ACCESS_ID", "ALIBABA_CLOUD_ACCESS_KEY_ID"
    )
    sk = args.access_key_secret or envget(
        "DATAWORKS_ACCESS_KEY_SECRET", "ODPS_ACCESS_KEY", "ALIBABA_CLOUD_ACCESS_KEY_SECRET"
    )
    region = (
        args.region
        or envget("DATAWORKS_REGION_ID")
        or region_from_endpoint(envget("ODPS_ENDPOINT"))
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
            "缺少 AccessKey。请设置 ODPS_ACCESS_ID/ODPS_ACCESS_KEY(复用 odps)或 "
            "DATAWORKS_ACCESS_KEY_ID/SECRET,或在 skill 目录 .env 中配置。",
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


def call(ctx, method_name, request):
    client = make_client(ctx)
    if ctx.debug:
        try:
            eprint("[debug] " + method_name + " " + json.dumps(request.to_map(), ensure_ascii=False, default=str))
        except Exception:
            eprint("[debug] " + method_name)
    try:
        resp = getattr(client, method_name)(request)
    except SystemExit:
        raise
    except Exception as e:  # noqa: BLE001
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


def cmd_file_list(ctx, args):
    pid = need_project(ctx)
    req = dw_models.ListFilesRequest(project_id=pid, page_number=args.page_number, page_size=args.page_size)
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
    } for f in files]
    out(ctx, rows, ["file_id", "file_name", "node_id", "file_type", "owner"])


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
    # 调度
    if args.cron:
        req.cron_express = args.cron
    if args.cycle_type:
        req.cycle_type = args.cycle_type
    if args.start_effect:
        req.start_effect_date = to_epoch_ms(args.start_effect)
    if args.end_effect:
        req.end_effect_date = to_epoch_ms(args.end_effect)
    # 重跑
    if args.rerun_mode:
        req.rerun_mode = args.rerun_mode
    if args.auto_rerun_times is not None:
        req.auto_rerun_times = args.auto_rerun_times
    if args.auto_rerun_interval is not None:
        req.auto_rerun_interval_millis = args.auto_rerun_interval
    # 依赖
    if args.dep_type:
        req.dependent_type = args.dep_type
    if args.dep_nodes:
        req.dependent_node_id_list = args.dep_nodes
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

    start = norm_biz_datetime(args.start_biz)
    end = norm_biz_datetime(args.end_biz) or start
    if not start:
        die("complement run 需要 --start-biz(YYYY-MM-DD)。", E_USAGE)
    name = args.name or f"dwcli-backfill-{node_id}-{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}"

    req = dw_models.RunCycleDagNodesRequest(
        project_env=ctx.env, name=name, root_node_id=int(root),
        include_node_ids=str(include), start_biz_date=start, end_biz_date=end,
    )
    if args.exclude_node_ids:
        req.exclude_node_ids = args.exclude_node_ids
    if args.parallelism is not None:
        req.parallelism = args.parallelism
    if args.node_params:
        req.node_params = args.node_params

    confirm_write(ctx, "RunCycleDagNodes", req.to_map())
    body = call(ctx, "run_cycle_dag_nodes", req)
    out(ctx, {"dag_id": g(body, "Data"), "name": name, "raw": body})


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
    parser.add_argument("--access-key-id", default=d(None), help="默认取 ODPS_ACCESS_ID / DATAWORKS_ACCESS_KEY_ID")
    parser.add_argument("--access-key-secret", default=d(None), help="默认取 ODPS_ACCESS_KEY / DATAWORKS_ACCESS_KEY_SECRET")
    parser.add_argument("-r", "--region", default=d(None), help="默认 DATAWORKS_REGION_ID 或从 ODPS_ENDPOINT 解析,否则 cn-shanghai")
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

    # file
    filep = sub.add_parser("file", help="文件/节点代码")
    files = filep.add_subparsers(dest="sub", required=True)
    fl = files.add_parser("list", parents=[parent], help="文件列表")
    fl.add_argument("--keyword"); fl.add_argument("--exact-file-name"); fl.add_argument("--node-id")
    fl.add_argument("--owner"); fl.add_argument("--page-number", type=int, default=1)
    fl.add_argument("--page-size", type=int, default=50)
    fl.set_defaults(func=cmd_file_list)
    fg = files.add_parser("get", parents=[parent], help="取文件/节点代码(--format text 仅输出代码)")
    fg.add_argument("--file-id"); fg.add_argument("--node-id")
    fg.set_defaults(func=cmd_file_get)
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
    fs = files.add_parser("submit", parents=[parent], help="提交文件(写,需 --yes)")
    fs.add_argument("--file-id", required=True); fs.add_argument("--comment")
    fs.set_defaults(func=cmd_file_submit)
    fd = files.add_parser("deploy", parents=[parent], help="发布文件到生产(写,需 --yes)")
    fd.add_argument("--file-id"); fd.add_argument("--node-id"); fd.add_argument("--comment")
    fd.set_defaults(func=cmd_file_deploy)

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
    cr.add_argument("--start-biz", help="起始业务日期 YYYY-MM-DD(必填)")
    cr.add_argument("--end-biz", help="结束业务日期 YYYY-MM-DD(默认=start)")
    cr.add_argument("--name", help="补数据任务名")
    cr.add_argument("--parallelism", type=int)
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
