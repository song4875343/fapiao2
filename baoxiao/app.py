"""Streamlit 控制页：邮箱下载处理或本地文件夹处理。"""

from datetime import datetime, time
from pathlib import Path
from tempfile import TemporaryDirectory
import shutil

import streamlit as st

from invoice_downloader import download_emails
from invoice_processor import CATEGORIES, categorize_pdfs, generate_excel
from taxi_reimbursement import generate_taxi_reimbursement


DEFAULT_TEMPLATE = Path(__file__).with_name("出租车报销单模板.xlsx")


def write_outputs(rows, output_dir, template):
    output = Path(output_dir)
    output.mkdir(parents=True, exist_ok=True)
    generate_excel(rows, output / "发票.xlsx")
    taxis = [row for row in rows if row["category"] == "出租票"]
    if taxis:
        generate_taxi_reimbursement(taxis, template, output / "出租车报销单.xlsx")


def process_local_folder(source_dir, output_dir, template=DEFAULT_TEMPLATE):
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
    write_outputs(rows, output, template)
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
    if st.button("开始下载并处理", type="primary", width="stretch"):
        try:
            if start > end:
                raise ValueError("开始日期不能晚于结束日期")
            with st.spinner("正在下载并处理..."):
                results, rows = download_emails(datetime.combine(start, time.min), datetime.combine(end, time.max), raw_dir, output_dir)
                write_outputs(rows, output_dir, template)
            show_result(results, output_dir)
        except Exception as exc:
            st.error(str(exc))
else:
    source_dir = st.text_input("本地发票文件夹", "./exm")
    output_dir = st.text_input("分类输出目录", "./本地发票结果")
    template = st.text_input("出租车报销单模板", str(DEFAULT_TEMPLATE))
    if st.button("开始处理本地文件", type="primary", width="stretch"):
        try:
            with st.spinner("正在处理..."):
                results = process_local_folder(source_dir, output_dir, template)
            show_result(results, output_dir)
        except Exception as exc:
            st.error(str(exc))
