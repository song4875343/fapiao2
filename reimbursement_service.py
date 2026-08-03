"""Integration layer for the advanced reimbursement workflow."""

from copy import deepcopy
from datetime import datetime, time
from pathlib import Path
from tempfile import TemporaryDirectory
import shutil
import sys


BASE_DIR = Path(__file__).resolve().parent
BAOXIAO_DIR = BASE_DIR / "baoxiao"
BAOXIAO_SITE_PACKAGES = BAOXIAO_DIR / ".venv" / "Lib" / "site-packages"

# The reimbursement module ships with its already-tested dependency environment.
# Prefer it when present so the main app and standalone module use identical parsers.
if BAOXIAO_SITE_PACKAGES.is_dir():
    dependency_path = str(BAOXIAO_SITE_PACKAGES)
    if dependency_path not in sys.path:
        sys.path.insert(0, dependency_path)
    for package_name in ("openpyxl", "dotenv", "et_xmlfile"):
        loaded = sys.modules.get(package_name)
        if loaded is not None and not getattr(loaded, "__file__", None):
            del sys.modules[package_name]

from baoxiao.invoice_processor import CATEGORIES, categorize_pdfs, generate_excel
from baoxiao.taxi_planner import export_fixed_templates, plan_taxis
from baoxiao.taxi_reimbursement import generate_taxi_reimbursement


DEFAULT_TEMPLATE = BAOXIAO_DIR / "出租车报销单模板.xlsx"
DOUBLE_TEMPLATE = BAOXIAO_DIR / "出租车报销单双页模板.xlsx"


def _write_outputs(rows, output_dir, unit_city="郑州", home_to_station=(20.0, 30.0),
                   station_to_project_min=30.0):
    output = Path(output_dir)
    output.mkdir(parents=True, exist_ok=True)
    generate_excel(rows, output / "发票.xlsx")
    taxis = [row for row in rows if row["category"] == "出租票"]
    trains = [row for row in rows if row["category"] == "高铁票"]
    if not taxis:
        return

    original_taxis = deepcopy(taxis)
    planned = plan_taxis(
        taxis, trains, unit_city, tuple(home_to_station), float(station_to_project_min)
    )
    for taxi in taxis:
        trips = taxi.get("trips") or []
        if trips:
            taxi["date"] = "、".join(dict.fromkeys(
                trip.get("date", "") for trip in trips if trip.get("date")
            ))
            taxi["route"] = "；".join(
                f"{trip.get('start', '')} - {trip.get('end', '')}" for trip in trips
            )
    generate_excel(rows, output / "发票.xlsx")

    planned_by_file = {row.get("filename"): row for row in taxis}
    for original in original_taxis:
        planned_row = planned_by_file.get(original.get("filename"), {})
        original["status"] = planned_row.get("status", "")
        original["note"] = planned_row.get("note", "")
        for index, trip in enumerate(original.get("trips") or []):
            planned_trips = planned_row.get("trips") or []
            if index < len(planned_trips):
                trip["status"] = planned_trips[index].get("status", "")
                trip["note"] = planned_trips[index].get("note", "")

    generate_taxi_reimbursement(original_taxis, DEFAULT_TEMPLATE, output / "出租车报销单.xlsx")
    export_fixed_templates(
        planned, DEFAULT_TEMPLATE, DOUBLE_TEMPLATE,
        double_path=output / "出租车报销单双页.xlsx",
    )


def _copy_and_process(source_dir, output_dir):
    source = Path(source_dir).expanduser().resolve()
    output = Path(output_dir).expanduser().resolve()
    if not source.is_dir():
        raise ValueError("本地发票文件夹不存在")
    pdfs = [path for path in source.rglob("*.pdf") if output not in path.resolve().parents]
    if not pdfs:
        raise ValueError("本地文件夹中没有 PDF")
    with TemporaryDirectory() as temp_dir:
        staging = Path(temp_dir)
        for path in pdfs:
            target = staging / path.relative_to(source)
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(path, target)
        return categorize_pdfs(staging.rglob("*.pdf"), output)


def _json_row(row):
    return {
        "category": row.get("category", ""),
        "date": row.get("date", ""),
        "content": row.get("route") or row.get("hotel") or row.get("kind") or "",
        "amount": float(row.get("amount") or 0),
        "invoiceNo": row.get("invoice_no", "未识别"),
        "filename": row.get("filename", ""),
        "path": row.get("path", ""),
    }


def _result_payload(results, rows, output_dir):
    output = Path(output_dir).expanduser().resolve()
    detail = []
    for category in CATEGORIES:
        category_rows = sorted(
            (row for row in rows if row.get("category") == category),
            key=lambda row: (
                row.get("date", ""), row.get("time", ""), row.get("filename", "")
            ),
        )
        detail.extend(_json_row(row) for row in category_rows)
    counts = {category: len(results.get(category, [])) for category in results}
    amounts = {
        category: round(sum(row["amount"] for row in detail if row["category"] == category), 2)
        for category in CATEGORIES
    }
    output_files = []
    for name in ("发票.xlsx", "出租车报销单.xlsx", "出租车报销单双页.xlsx"):
        path = output / name
        if path.exists():
            output_files.append({"name": name, "path": str(path)})
    print_files = list(dict.fromkeys(row["path"] for row in detail if row["path"]))
    return {
        "success": True,
        "outputDir": str(output),
        "counts": counts,
        "amounts": amounts,
        "totalCount": len(detail),
        "totalAmount": round(sum(row["amount"] for row in detail), 2),
        "rows": detail,
        "outputFiles": output_files,
        "printFiles": print_files,
    }


def process(options):
    mode = options.get("mode", "local")
    output_dir = options.get("outputDir", "").strip()
    if not output_dir:
        raise ValueError("请选择分类输出目录")
    if mode == "email":
        from baoxiao.invoice_downloader import download_emails

        start = datetime.combine(datetime.fromisoformat(options["startDate"]).date(), time.min)
        end = datetime.combine(datetime.fromisoformat(options["endDate"]).date(), time.max)
        if start > end:
            raise ValueError("开始日期不能晚于结束日期")
        raw_dir = options.get("rawDir", "").strip()
        if not raw_dir:
            raise ValueError("请选择附件下载目录")
        results, rows = download_emails(start, end, raw_dir, output_dir)
    else:
        results, rows = _copy_and_process(options.get("sourceDir", ""), output_dir)
    _write_outputs(
        rows, output_dir, options.get("unitCity", "郑州"),
        (float(options.get("homeMin", 20)), float(options.get("homeMax", 30))),
        float(options.get("projectMin", 30)),
    )
    return _result_payload(results, rows, output_dir)
