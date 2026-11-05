import copy
import json
import os
import re
import shutil
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
import uuid
import zipfile
from datetime import datetime
from io import BytesIO
from pathlib import Path
from typing import Dict, List, Optional

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

APP_DIR = Path(os.getenv("APP_DIR", "/app"))
DATA_DIR = Path(os.getenv("DATA_DIR", str(APP_DIR / "data")))
UPLOAD_DIR = DATA_DIR / "uploads"
OUTPUT_DIR = DATA_DIR / "output"
TEMP_DIR = DATA_DIR / "temp"
PUBLIC_DIR = APP_DIR / "public"
BLANK_TEMPLATE_DIR = Path(os.getenv("BLANK_TEMPLATE_DIR", str(APP_DIR / "blank-templates")))
STANDARDS_DIR = Path(os.getenv("STANDARDS_DIR", str(DATA_DIR / "standards")))
MAX_UPLOAD_MB = int(os.getenv("MAX_UPLOAD_MB", "80"))
MAX_TDS_FILES = int(os.getenv("MAX_TDS_FILES", "10"))
FILE_TTL_HOURS = int(os.getenv("FILE_TTL_HOURS", "24"))
DOC_CONVERSION_TIMEOUT_SECONDS = int(os.getenv("DOC_CONVERSION_TIMEOUT_SECONDS", "150"))
DOC_CONVERTER_URL = os.getenv("DOC_CONVERTER_URL", "http://libreoffice-converter:8080/convert")
DOC_CONVERTER_HEALTH_URL = os.getenv("DOC_CONVERTER_HEALTH_URL", "http://libreoffice-converter:8080/health")
DOC_CONVERTER_TOKEN = os.getenv("DOC_CONVERTER_TOKEN", "")
BLANK_TEMPLATE_FILE = os.getenv("BLANK_TEMPLATE_FILE", "").strip()

for directory in (UPLOAD_DIR, OUTPUT_DIR, TEMP_DIR, BLANK_TEMPLATE_DIR):
    directory.mkdir(parents=True, exist_ok=True)

app = FastAPI(title="TDS to Report Importer", version="2.0.0-multi-blank")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_credentials=False, allow_methods=["*"], allow_headers=["*"])
JOBS: Dict[str, dict] = {}
JOBS_LOCK = threading.Lock()
CLAUSE_RE = re.compile(r"^(?:TABLE\s*:\s*)?(?:Clause\s+)?(\d+(?:\.\d+){1,4}[A-Za-z]?)(?=\s|:|$)", re.I)
END_RE = re.compile(r"^(?:Supplementary\s+information|Note\s*1)\s*:", re.I)
SPECIAL_894_RE = re.compile(r"INSULATION\s+DIAGRAM", re.I)
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
FORMAT_TAGS = ("rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline", "shadow", "emboss", "imprint", "color", "spacing", "w", "kern", "position", "sz", "szCs", "highlight", "u", "vertAlign", "rtl", "cs", "lang", "eastAsianLayout")


def set_job(job_id: str, **changes):
    with JOBS_LOCK:
        JOBS.setdefault(job_id, {}).update(changes)


def safe_name(name: str) -> str:
    name = re.sub(r'[\\/:*?"<>|]+', "_", Path(name or "file.docx").name)
    return name[:180] or "file.docx"


def ensure_word_file(upload: UploadFile) -> str:
    suffix = Path(upload.filename or "").suffix.lower()
    if suffix not in {".doc", ".docx"}:
        raise HTTPException(400, "仅支持 DOC 或 DOCX 文件")
    return suffix


def read_upload(upload: UploadFile) -> bytes:
    content = upload.file.read()
    if not content:
        raise HTTPException(400, f"{upload.filename} 是空文件")
    if len(content) > MAX_UPLOAD_MB * 1024 * 1024:
        raise HTTPException(413, f"单个文件不能超过 {MAX_UPLOAD_MB} MB")
    return content


def validate_docx_bytes(content: bytes, filename: str):
    try:
        with zipfile.ZipFile(BytesIO(content), "r") as zf:
            if "word/document.xml" not in zf.namelist() or zf.testzip():
                raise ValueError("invalid docx")
        Document(BytesIO(content))
    except Exception as exc:
        raise HTTPException(400, f"{filename} 不是有效的 DOCX 文件") from exc


def converter_health() -> dict:
    try:
        with urllib.request.urlopen(urllib.request.Request(DOC_CONVERTER_HEALTH_URL), timeout=5) as response:
            return json.loads(response.read().decode("utf-8"))
    except Exception as exc:
        return {"status": "unavailable", "message": str(exc), "active": False, "waiting": 0, "concurrency": 1}


def convert_doc_bytes_to_docx(content: bytes, original_name: str) -> bytes:
    headers = {"Content-Type": "application/msword", "X-Filename": urllib.parse.quote(Path(original_name).name)}
    if DOC_CONVERTER_TOKEN:
        headers["X-Converter-Token"] = DOC_CONVERTER_TOKEN
    request = urllib.request.Request(DOC_CONVERTER_URL, data=content, headers=headers, method="POST")
    try:
        with urllib.request.urlopen(request, timeout=DOC_CONVERSION_TIMEOUT_SECONDS) as response:
            converted = response.read()
    except urllib.error.HTTPError as exc:
        try:
            message = json.loads(exc.read().decode("utf-8", errors="replace")).get("message", str(exc))
        except Exception:
            message = str(exc)
        raise HTTPException(422, f"DOC 转 DOCX 失败：{message}") from exc
    except urllib.error.URLError as exc:
        raise HTTPException(503, f"无法连接 LibreOffice 转换服务：{exc.reason}") from exc
    except TimeoutError as exc:
        raise HTTPException(408, "DOC 转 DOCX 超时，请稍后重试") from exc
    validate_docx_bytes(converted, original_name)
    return converted


def prepare_word_bytes(upload: UploadFile) -> bytes:
    suffix = ensure_word_file(upload)
    content = read_upload(upload)
    if suffix == ".doc":
        return convert_doc_bytes_to_docx(content, upload.filename or "input.doc")
    validate_docx_bytes(content, upload.filename or "input.docx")
    return content


def normalize_clause_text(text: str) -> str:
    text = re.sub(r"\s+", " ", str(text or "")).strip()
    return re.sub(r"(?<=\d)\s*\.\s*(?=\d)", ".", text)


def cell_text(tc) -> str:
    return " ".join(t.text for t in tc.findall(f".//{qn('w:t')}") if t.text).strip()


def table_first_row_text(tbl) -> str:
    tr = tbl.find(qn("w:tr"))
    return "" if tr is None else " ".join(cell_text(tc) for tc in tr.findall(qn("w:tc"))).strip()


def table_clause(tbl) -> Optional[str]:
    match = CLAUSE_RE.search(normalize_clause_text(table_first_row_text(tbl)))
    return match.group(1) if match else None


def table_special_clause(tbl) -> Optional[str]:
    return "8.9.4" if SPECIAL_894_RE.search(normalize_clause_text(table_first_row_text(tbl))) else None


def table_has_end(tbl) -> bool:
    return any(END_RE.search(cell_text(tc) or "") for tc in tbl.findall(f".//{qn('w:tc')}") )


def body_tables(doc: Document):
    return list(doc.element.body.findall(qn("w:tbl")))


def scan_groups(doc: Document) -> Dict[str, List]:
    groups: Dict[str, List] = {}
    active: Optional[str] = None
    for tbl in body_tables(doc):
        clause = table_clause(tbl) or table_special_clause(tbl)
        if clause:
            active = clause
            groups.setdefault(clause, []).append(tbl)
        elif active:
            groups.setdefault(active, []).append(tbl)
        if active and table_has_end(tbl):
            active = None
    return groups


def build_style_context(doc: Document) -> dict:
    root = doc.styles.element
    styles = {}
    for style in root.findall(qn("w:style")):
        sid = style.get(qn("w:styleId"))
        if sid:
            based = style.find(qn("w:basedOn"))
            styles[sid] = {"based": based.get(qn("w:val")) if based is not None else None, "rPr": style.find(qn("w:rPr"))}
    defaults = root.find(f".//{qn('w:docDefaults')}/{qn('w:rPrDefault')}/{qn('w:rPr')}")
    return {"styles": styles, "defaults": defaults, "cache": {}}


def merge_rpr(target: dict, rpr):
    if rpr is None:
        return
    for name in FORMAT_TAGS:
        node = rpr.find(qn(f"w:{name}"))
        if node is not None:
            target[name] = copy.deepcopy(node)


def style_properties(style_id: Optional[str], context: dict, trail=None) -> dict:
    if not style_id:
        return {}
    if style_id in context["cache"]:
        return {k: copy.deepcopy(v) for k, v in context["cache"][style_id].items()}
    trail = set(trail or ())
    if style_id in trail or style_id not in context["styles"]:
        return {}
    trail.add(style_id)
    style = context["styles"][style_id]
    props = style_properties(style["based"], context, trail)
    merge_rpr(props, style["rPr"])
    context["cache"][style_id] = {k: copy.deepcopy(v) for k, v in props.items()}
    return props


def materialize_table_format(tbl, context: dict) -> int:
    default_props = {}
    merge_rpr(default_props, context.get("defaults"))
    changed = 0
    for paragraph in tbl.findall(f".//{qn('w:p')}"):
        ppr = paragraph.find(qn("w:pPr"))
        pstyle_id = None
        if ppr is not None:
            pstyle = ppr.find(qn("w:pStyle"))
            pstyle_id = pstyle.get(qn("w:val")) if pstyle is not None else None
        paragraph_props = dict(default_props)
        paragraph_props.update(style_properties(pstyle_id, context))
        for run in paragraph.findall(qn("w:r")):
            rpr = run.find(qn("w:rPr"))
            if rpr is None:
                rpr = OxmlElement("w:rPr")
                run.insert(0, rpr)
            rstyle = rpr.find(qn("w:rStyle"))
            effective = dict(paragraph_props)
            effective.update(style_properties(rstyle.get(qn("w:val")) if rstyle is not None else None, context))
            for name, source in effective.items():
                if rpr.find(qn(f"w:{name}")) is None:
                    rpr.append(copy.deepcopy(source))
                    changed += 1
    return changed


def reset_numbering(tbl) -> int:
    changed = 0
    for paragraph in tbl.findall(f".//{qn('w:p')}"):
        ppr = paragraph.find(qn("w:pPr"))
        if ppr is not None:
            for tag in (qn("w:numPr"), qn("w:outlineLvl")):
                for node in list(ppr.findall(tag)):
                    ppr.remove(node)
                    changed += 1
    return changed


def sanitize_relationships(tbl) -> dict:
    removed_objects = removed_attrs = unwrapped_links = 0
    for hyperlink in list(tbl.findall(f".//{qn('w:hyperlink')}")):
        parent = hyperlink.getparent()
        if parent is not None:
            position = parent.index(hyperlink)
            for child in list(hyperlink):
                hyperlink.remove(child); parent.insert(position, child); position += 1
            parent.remove(hyperlink); unwrapped_links += 1
    removable = {"drawing", "pict", "object", "OLEObject", "AlternateContent", "control", "movie", "imagedata", "chart", "diagram", "anchor", "inline"}
    for element in list(tbl.iter()):
        local = element.tag.split("}")[-1] if "}" in element.tag else element.tag
        if local in removable:
            parent = element.getparent()
            if parent is not None:
                parent.remove(element); removed_objects += 1
            continue
        for attr in list(element.attrib):
            if attr.startswith("{" + REL_NS + "}"):
                del element.attrib[attr]; removed_attrs += 1
    return {"objects": removed_objects, "attrs": removed_attrs, "links": unwrapped_links}


def find_blank_template() -> Optional[Path]:
    if BLANK_TEMPLATE_FILE:
        path = BLANK_TEMPLATE_DIR / BLANK_TEMPLATE_FILE
        return path if path.exists() else None
    candidates = sorted(BLANK_TEMPLATE_DIR.glob("*.docx"))
    return candidates[0] if candidates else None


def load_blank_source():
    path = find_blank_template()
    if not path:
        return None, {}, None
    doc = Document(str(path))
    return path, scan_groups(doc), build_style_context(doc)


def merge_tds_sources(tds_docs: List[Document], names: List[str]):
    merged = {}
    contexts = {}
    candidates = {}
    for index, (doc, name) in enumerate(zip(tds_docs, names)):
        context = build_style_context(doc)
        for clause, group in scan_groups(doc).items():
            candidates.setdefault(clause, []).append(name)
            # Last uploaded TDS wins for duplicate clauses.
            merged[clause] = group
            contexts[clause] = context
    return merged, contexts, candidates


def replace_group(old_tables: List, new_tables: List, style_context: dict) -> dict:
    if not old_tables:
        raise ValueError("目标报告表格组为空")
    parent = old_tables[0].getparent()
    insert_at = parent.index(old_tables[0])
    for tbl in old_tables:
        if tbl.getparent() is parent:
            parent.remove(tbl)
    stats = {"inserted": 0, "numbering_removed": 0, "format_properties_materialized": 0, "removed_objects": 0, "removed_rel_attrs": 0, "unwrapped_links": 0}
    for offset, source in enumerate(new_tables):
        clone = copy.deepcopy(source)
        stats["format_properties_materialized"] += materialize_table_format(clone, style_context)
        stats["numbering_removed"] += reset_numbering(clone)
        rel = sanitize_relationships(clone)
        stats["removed_objects"] += rel["objects"]
        stats["removed_rel_attrs"] += rel["attrs"]
        stats["unwrapped_links"] += rel["links"]
        parent.insert(insert_at + offset, clone)
        stats["inserted"] += 1
    return stats


def validate_docx(path: Path):
    with zipfile.ZipFile(path, "r") as zf:
        damaged = zf.testzip()
        if damaged:
            raise RuntimeError(f"生成的 DOCX 压缩包损坏：{damaged}")
    Document(str(path))


def prepare_preview(tds_docs: List[Document], tds_names: List[str], report_doc: Document):
    tds_groups, _, candidates = merge_tds_sources(tds_docs, tds_names)
    report_groups = scan_groups(report_doc)
    blank_path, blank_groups, _ = load_blank_source()
    rows = []
    for clause in sorted(report_groups, key=lambda x: [int(p) if p.isdigit() else p for p in re.split(r"(\d+)", x)]):
        if clause in tds_groups:
            source = "tds"
            source_name = candidates[clause][-1]
            status = "tds"
        elif clause in blank_groups:
            source = "blank"
            source_name = blank_path.name if blank_path else ""
            status = "blank"
        else:
            source = "none"
            source_name = ""
            status = "missing"
        rows.append({
            "clause": clause,
            "report_tables": len(report_groups[clause]),
            "replacement_tables": len(tds_groups.get(clause, blank_groups.get(clause, []))),
            "source": source,
            "source_name": source_name,
            "status": status,
            "selected": source != "none",
            "conflict": len(candidates.get(clause, [])) > 1,
            "candidate_tds": candidates.get(clause, []),
        })
    return {
        "report_clause_count": len(report_groups),
        "tds_clause_count": len(tds_groups),
        "tds_file_count": len(tds_docs),
        "blank_template": blank_path.name if blank_path else "",
        "tds_count": sum(1 for row in rows if row["source"] == "tds"),
        "blank_count": sum(1 for row in rows if row["source"] == "blank"),
        "missing_count": sum(1 for row in rows if row["source"] == "none"),
        "conflict_count": sum(1 for row in rows if row["conflict"]),
        "rows": rows,
    }


def process_job(job_id: str, selected_clauses: List[str]):
    job = JOBS[job_id]
    report_path = Path(job["report_path"])
    tds_paths = [Path(path) for path in job["tds_paths"]]
    try:
        set_job(job_id, status="processing", progress=6, message="正在读取多份 TDS 和报告")
        tds_docs = [Document(str(path)) for path in tds_paths]
        report_doc = Document(str(report_path))
        tds_groups, tds_contexts, _ = merge_tds_sources(tds_docs, job["tds_names"])
        blank_path, blank_groups, blank_context = load_blank_source()
        report_groups = scan_groups(report_doc)
        total = max(1, len(selected_clauses))
        details = []
        for index, clause in enumerate(selected_clauses, start=1):
            if clause not in report_groups:
                details.append({"clause": clause, "status": "skipped_no_report_table"}); continue
            if clause in tds_groups:
                source_group, source_context, source = tds_groups[clause], tds_contexts[clause], "tds"
            elif clause in blank_groups and blank_context:
                source_group, source_context, source = blank_groups[clause], blank_context, "blank"
            else:
                details.append({"clause": clause, "status": "skipped_no_source"}); continue
            stats = replace_group(report_groups[clause], source_group, source_context)
            details.append({"clause": clause, "status": "replaced", "source": source, **stats})
            set_job(job_id, progress=8 + int(index / total * 80), message=f"正在替换 Clause {clause} · {source.upper()}")
        original_stem = Path(job.get("report_name") or report_path.name).stem
        output_name = f"{original_stem}_Imported_from_TDS.docx"
        output_path = OUTPUT_DIR / f"{job_id}_{safe_name(output_name)}"
        report_doc.save(output_path)
        set_job(job_id, progress=94, message="正在校验新报告")
        validate_docx(output_path)
        set_job(job_id, status="done", progress=100, message="导入完成", output_path=str(output_path), output_name=output_name, details=details)
    except Exception as exc:
        set_job(job_id, status="failed", progress=100, message=str(exc), error=repr(exc))
    finally:
        for path in [report_path, *tds_paths]:
            try: path.unlink(missing_ok=True)
            except Exception: pass


def cleanup_old_files():
    cutoff = time.time() - FILE_TTL_HOURS * 3600
    for directory in (UPLOAD_DIR, OUTPUT_DIR, TEMP_DIR):
        for path in directory.glob("*"):
            try:
                if path.is_file() and path.stat().st_mtime < cutoff: path.unlink()
            except Exception: pass


@app.get("/", response_class=HTMLResponse)
def home():
    path = PUBLIC_DIR / "index.html"
    if not path.exists(): raise HTTPException(500, "public/index.html 不存在")
    return path.read_text(encoding="utf-8")


@app.get("/api/health")
def health():
    converter = converter_health()
    blank = find_blank_template()
    return {
        "status": "ok", "time": datetime.now().isoformat(timespec="seconds"),
        "doc_conversion": converter.get("status") == "ok",
        "converter_active": converter.get("active", False),
        "converter_waiting": converter.get("waiting", 0),
        "blank_template_ready": bool(blank),
        "blank_template": blank.name if blank else "",
    }


@app.post("/api/preview")
def preview(tds: List[UploadFile] = File(...), report: UploadFile = File(...)):
    cleanup_old_files()
    if not tds: raise HTTPException(400, "请至少上传一份 TDS")
    if len(tds) > MAX_TDS_FILES: raise HTTPException(400, f"最多上传 {MAX_TDS_FILES} 份 TDS")
    tds_bytes = [prepare_word_bytes(item) for item in tds]
    report_bytes = prepare_word_bytes(report)
    tds_docs = [Document(BytesIO(content)) for content in tds_bytes]
    report_doc = Document(BytesIO(report_bytes))
    result = prepare_preview(tds_docs, [safe_name(item.filename) for item in tds], report_doc)
    job_id = uuid.uuid4().hex
    tds_paths = []
    for index, content in enumerate(tds_bytes):
        path = UPLOAD_DIR / f"{job_id}_tds_{index}.docx"; path.write_bytes(content); tds_paths.append(str(path))
    report_path = UPLOAD_DIR / f"{job_id}_report.docx"; report_path.write_bytes(report_bytes)
    with JOBS_LOCK:
        JOBS[job_id] = {
            "status": "preview", "progress": 0, "message": "等待确认",
            "tds_paths": tds_paths, "report_path": str(report_path),
            "tds_names": [safe_name(item.filename) for item in tds],
            "report_name": safe_name(report.filename), "created_at": time.time(),
        }
    return {"job_id": job_id, **result}


@app.post("/api/start/{job_id}")
def start(job_id: str, clauses_json: str = Form("[]")):
    with JOBS_LOCK: job = JOBS.get(job_id)
    if not job: raise HTTPException(404, "任务不存在或已过期")
    if job.get("status") not in {"preview", "failed"}: raise HTTPException(409, "任务已开始")
    try: clauses = json.loads(clauses_json)
    except Exception: raise HTTPException(400, "条款列表格式错误")
    if not isinstance(clauses, list) or not clauses: raise HTTPException(400, "请至少选择一个可替换条款")
    threading.Thread(target=process_job, args=(job_id, [str(x) for x in clauses]), daemon=True).start()
    return {"status": "started", "job_id": job_id}


@app.get("/api/jobs/{job_id}")
def job_status(job_id: str):
    with JOBS_LOCK: job = dict(JOBS.get(job_id, {}))
    if not job: raise HTTPException(404, "任务不存在")
    for key in ("tds_paths", "report_path", "output_path"): job.pop(key, None)
    return job


@app.get("/api/download/{job_id}")
def download(job_id: str):
    with JOBS_LOCK: job = JOBS.get(job_id)
    if not job or job.get("status") != "done": raise HTTPException(404, "文件尚未生成")
    path = Path(job["output_path"])
    if not path.exists(): raise HTTPException(404, "输出文件已过期")
    return FileResponse(path, media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document", filename=job["output_name"])
