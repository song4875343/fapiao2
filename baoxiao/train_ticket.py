import os
import re

import fitz


DATE_PATTERN = re.compile(r"(20\d{2})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日")
TIME_PATTERN = re.compile(r"([01]?\d|2[0-3])\s*:\s*([0-5]\d)\s*开")
TRAVEL_DATETIME_PATTERN = re.compile(
    r"(20\d{2})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日\s+"
    r"([01]?\d|2[0-3])\s*:\s*([0-5]\d)\s*开"
)
# Some railway PDFs store the currency symbol and amount in a later text block,
# even though they are visually adjacent to the "票价" label on the ticket.
AMOUNT_PATTERN = re.compile(r"票价\s*[：:]?.{0,800}?[¥￥]\s*([\d,]+(?:\.\d{1,2})?)")
REFUND_FEE_PATTERN = re.compile(r"退票费\s*[：:]?.{0,800}?[¥￥]\s*([\d,]+(?:\.\d{1,2})?)")
INVOICE_PATTERN = re.compile(r"发票号码\s*[：:]?\s*(\d{18,20})")
TRAIN_PATTERN = re.compile(r"\b([GDKCTZ]\d{1,4})\b", re.IGNORECASE)
STATION_PATTERN = re.compile(r"([\u4e00-\u9fff]{2,12})\s*站")


def _compact(text):
    return re.sub(r"\s+", " ", text or "").strip()


def is_train_ticket_text(text):
    """Return whether extracted PDF text belongs to a railway e-ticket."""
    return "铁路电子客票" in _compact(text)


def _extract_station_lines(page):
    candidates = []
    page_dict = page.get_text("dict")
    for block in page_dict.get("blocks", []):
        for line in block.get("lines", []):
            spans = line.get("spans", [])
            if not spans:
                continue
            line_text = "".join(span.get("text", "") for span in spans)
            for match in STATION_PATTERN.finditer(line_text):
                station = match.group(1)
                if station in {"车站", "到站", "发站"}:
                    continue
                candidates.append({
                    "name": station,
                    "x": min(span["bbox"][0] for span in spans),
                    "y": min(span["bbox"][1] for span in spans),
                    "size": max(span.get("size", 0) for span in spans),
                })

    if len(candidates) < 2:
        return None, None

    # Station names are the two largest labels on the same horizontal band.
    candidates.sort(key=lambda item: (-item["size"], item["y"], item["x"]))
    for first in candidates:
        same_band = [
            item for item in candidates
            if item["name"] != first["name"] and abs(item["y"] - first["y"]) <= 25
        ]
        if same_band:
            second = max(same_band, key=lambda item: item["size"])
            ordered = sorted((first, second), key=lambda item: item["x"])
            return ordered[0]["name"], ordered[1]["name"]
    return None, None


def extract_ticket(pdf_path, page_index=0, source=None):
    """Extract one railway e-ticket page into a normalized row."""
    doc = fitz.open(pdf_path)
    try:
        if page_index < 0 or page_index >= doc.page_count:
            raise ValueError("页码超出范围")
        page = doc[page_index]
        text = page.get_text()
        compact_text = _compact(text)
        if not is_train_ticket_text(text):
            raise ValueError("不是铁路电子客票")
        is_refund = "退票" in compact_text

        travel_datetime_match = TRAVEL_DATETIME_PATTERN.search(compact_text)
        date_match = travel_datetime_match or DATE_PATTERN.search(compact_text)
        time_match = travel_datetime_match or TIME_PATTERN.search(compact_text)
        amount_match = AMOUNT_PATTERN.search(compact_text)
        refund_fee_match = REFUND_FEE_PATTERN.search(compact_text)
        if not date_match or not time_match or (not is_refund and not amount_match) or (is_refund and not refund_fee_match):
            missing = []
            if not date_match:
                missing.append("日期")
            if not time_match:
                missing.append("发车时间")
            if not amount_match and not is_refund:
                missing.append("票价")
            if is_refund and not refund_fee_match:
                missing.append("退票费")
            raise ValueError("未识别到" + "、".join(missing))

        start, end = _extract_station_lines(page)
        if not start or not end:
            raise ValueError("未识别到出发站或到达站")

        if travel_datetime_match:
            year, month, day, hour, minute = (int(value) for value in travel_datetime_match.groups())
        else:
            year, month, day = (int(value) for value in date_match.groups())
            hour, minute = (int(value) for value in time_match.groups())
        amount = float(amount_match.group(1).replace(",", "")) if amount_match else 0.0
        refund_fee = float(refund_fee_match.group(1).replace(",", "")) if refund_fee_match else 0.0
        invoice_match = INVOICE_PATTERN.search(compact_text)
        train_match = TRAIN_PATTERN.search(compact_text)

        return {
            "date": f"{year:04d}-{month:02d}-{day:02d}",
            "time": f"{hour:02d}:{minute:02d}",
            "start": start,
            "end": end,
            "route": f"{start} - {end}",
            "amount": f"{amount:.2f}",
            "refundFee": f"{refund_fee:.2f}",
            "isRefund": is_refund,
            "trainNo": train_match.group(1).upper() if train_match else "",
            "invoiceNo": invoice_match.group(1) if invoice_match else "未识别",
            "source": source or f"{os.path.basename(pdf_path)}-P{page_index + 1}",
        }
    finally:
        doc.close()


def generate_form(pages):
    rows = []
    errors = []
    for page in pages:
        path = page["path"]
        page_index = int(page.get("pageIndex", page.get("page_index", 0)))
        source = page.get("source") or f"{page.get('fileName', os.path.basename(path))}-P{page_index + 1}"
        try:
            rows.append(extract_ticket(path, page_index, source))
        except Exception as exc:
            errors.append({"source": source, "error": str(exc)})

    rows.sort(key=lambda row: (row["date"], row["time"], row["source"]))
    if not rows:
        detail = "；".join(f"{item['source']}：{item['error']}" for item in errors)
        return {"success": False, "error": detail or "未识别到高铁票", "rows": [], "errors": errors}

    return {"success": True, "rows": rows, "errors": errors}
