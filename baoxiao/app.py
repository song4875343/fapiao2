"""Streamlit 控制页：邮箱下载处理或本地文件夹处理。"""

from datetime import datetime, time
from copy import deepcopy
from pathlib import Path
from tempfile import TemporaryDirectory
import shutil

import streamlit as st

from invoice_downloader import download_emails
from invoice_processor import CATEGORIES, categorize_pdfs, generate_excel
from taxi_reimbursement import generate_taxi_reimbursement
from taxi_planner import export_fixed_templates, plan_taxis


DEFAULT_TEMPLATE = Path(__file__).with_name("出租车报销单模板.xlsx")


def write_outputs(rows, output_dir, template, unit_city="郑州", home_to_station=(20.0, 30.0), station_to_project_min=30.0):
    output = Path(output_dir)
    output.mkdir(parents=True, exist_ok=True)
    generate_excel(rows, output / "发票.xlsx")
    taxis = [row for row in rows if row["category"] == "出租票"]
    trains = [row for row in rows if row["category"] == "高铁票"]
    if taxis:
        original_taxis = deepcopy(taxis)
        planned = plan_taxis(taxis, trains, unit_city, home_to_station, station_to_project_min)
        # Reflect redesigned trips in the detailed taxi sheet used by the summary.
        for taxi in taxis:
            trips = taxi.get("trips") or []
            if trips:
                taxi["date"] = "、".join(dict.fromkeys(t.get("date", "") for t in trips if t.get("date")))
                taxi["route"] = "；".join(f"{t.get('start', '')} - {t.get('end', '')}" for t in trips)
        generate_excel(rows, output / "发票.xlsx")
        # Intermediate single-page output keeps the previous dynamic layout,
        # split coloring and expandable rows. Only the print output is fixed.
        # Restore original date/route/amount values for the intermediate sheet,
        # while carrying over only the planning status annotations.
        planned_by_file = {row.get("filename"): row for row in taxis}
        for original in original_taxis:
            planned_row = planned_by_file.get(original.get("filename"), {})
            original["status"] = planned_row.get("status", "")
            original["note"] = planned_row.get("note", "")
            original_trips = original.get("trips") or []
            planned_trips = planned_row.get("trips") or []
            for index, trip in enumerate(original_trips):
                if index < len(planned_trips):
                    trip["status"] = planned_trips[index].get("status", "")
                    trip["note"] = planned_trips[index].get("note", "")
        generate_taxi_reimbursement(original_taxis, template, output / "出租车报销单.xlsx")
        export_fixed_templates(
            planned, template, Path(__file__).with_name("出租车报销单双页模板.xlsx"),
            double_path=output / "出租车报销单双页.xlsx",
        )


def process_local_folder(source_dir, output_dir, template=DEFAULT_TEMPLATE, unit_city="郑州",
                         home_to_station=(20.0, 30.0), station_to_project_min=30.0):
    source, output = Path(source_dir).resolve(), Path(output_dir).resolve()
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
        results, rows = categorize_pdfs(staging.rglob("*.pdf"), output)
    write_outputs(rows, output, template, unit_city, home_to_station, station_to_project_min)
    return results


def show_result(results, output_dir):
    st.success(f"处理完成，结果保存在：{Path(output_dir).resolve()}")
    st.dataframe({"分类": list(results), "数量": [len(files) for files in results.values()]}, hide_index=True, width="stretch")


st.set_page_config(page_title="发票整理", page_icon="🧾", layout="centered")
st.title("发票下载与整理")
mode = st.radio("处理模式", ["邮箱下载并处理", "本地文件夹处理"], horizontal=True)

if mode == "邮箱下载并处理":
    left, right = st.columns(2)
    start = left.date_input("开始日期")
    end = right.date_input("结束日期")
    raw_dir = st.text_input("附件下载目录", "./tem")
    output_dir = st.text_input("分类输出目录", "./发票")
    template = st.text_input("出租车报销单模板", str(DEFAULT_TEMPLATE))
    unit_city = st.text_input("单位所在城市", "郑州")
    home_range = st.slider("家到高铁站费用范围", 0, 200, (20, 30))
    station_min = st.number_input("高铁站到项目所在地最低费用", min_value=0.0, value=30.0, step=1.0)
    if st.button("开始下载并处理", type="primary", width="stretch"):
        try:
            if start > end:
                raise ValueError("开始日期不能晚于结束日期")
            with st.spinner("正在下载并处理..."):
                results, rows = download_emails(datetime.combine(start, time.min), datetime.combine(end, time.max), raw_dir, output_dir)
                write_outputs(rows, output_dir, template, unit_city, home_range, station_min)
            show_result(results, output_dir)
        except Exception as exc:
            st.error(str(exc))
else:
    source_dir = st.text_input("本地发票文件夹", "./exm")
    output_dir = st.text_input("分类输出目录", "./本地发票结果")
    template = st.text_input("出租车报销单模板", str(DEFAULT_TEMPLATE))
    unit_city = st.text_input("单位所在城市", "郑州")
    home_range = st.slider("家到高铁站费用范围", 0, 200, (20, 30))
    station_min = st.number_input("高铁站到项目所在地最低费用", min_value=0.0, value=30.0, step=1.0)
    if st.button("开始处理本地文件", type="primary", width="stretch"):
        try:
            with st.spinner("正在处理..."):
                results = process_local_folder(source_dir, output_dir, template, unit_city, home_range, station_min)
            show_result(results, output_dir)
        except Exception as exc:
            st.error(str(exc))
