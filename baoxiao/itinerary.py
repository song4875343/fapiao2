"""网约车行程单识别及逐段行程提取。"""

from pathlib import Path
import re
import unicodedata

import fitz


HEADERS = ("序号", "服务商", "车型", "上车时间", "城市", "起点", "终点", "金额")


def _compact(text):
    return re.sub(r"\s+", "", unicodedata.normalize("NFKC", text or ""))


def is_itinerary_text(text):
    compact = _compact(text)
    has_title = "行程单" in compact or "行程明细" in compact
    has_table = all(header in compact for header in ("上车时间", "起点", "终点", "金额"))
    return has_title and has_table


def _column_boundaries(header_words):
    centers = [(word[0] + word[2]) / 2 for word in header_words]
    return [(centers[index] + centers[index + 1]) / 2 for index in range(len(centers) - 1)]


def _column_index(word, boundaries):
    center = (word[0] + word[2]) / 2
    for index, boundary in enumerate(boundaries):
        if center < boundary:
            return index
    return len(boundaries)


def _join_words(words, spaced=False):
    ordered = sorted(words, key=lambda word: (round(word[1], 1), word[0]))
    return (" " if spaced else "").join(word[4].strip() for word in ordered).strip()


def _extract_page_trips(page):
    words = page.get_text("words")
    header_words = []
    for header in HEADERS:
        matches = [word for word in words if word[4].strip() == header]
        if not matches:
            return []
        header_words.append(min(matches, key=lambda word: word[1]))
    header_words.sort(key=lambda word: word[0])
    boundaries = _column_boundaries(header_words)
    header_bottom = max(word[3] for word in header_words)

    sequence_words = [
        word for word in words
        if word[1] > header_bottom and _column_index(word, boundaries) == 0 and re.fullmatch(r"\d+", word[4].strip())
    ]
    sequence_words.sort(key=lambda word: word[1])
    trips = []
    for index, sequence in enumerate(sequence_words):
        top = sequence[1] - 8
        bottom = sequence_words[index + 1][1] - 8 if index + 1 < len(sequence_words) else sequence[3] + 25
        row_words = [word for word in words if top <= (word[1] + word[3]) / 2 < bottom]
        columns = [[] for _ in HEADERS]
        for word in row_words:
            columns[_column_index(word, boundaries)].append(word)
        datetime_text = _join_words(columns[3], spaced=True)
        date_match = re.search(r"20\d{2}-\d{2}-\d{2}", datetime_text)
        time_match = re.search(r"\d{1,2}:\d{2}", datetime_text)
        amount_match = re.search(r"([\d,]+(?:\.\d{1,2})?)", _join_words(columns[7]))
        start, end = _join_words(columns[5]), _join_words(columns[6])
        if date_match and amount_match and start and end:
            trips.append({
                "sequence": int(sequence[4]),
                "provider": _join_words(columns[1]),
                "vehicle": _join_words(columns[2]),
                "date": date_match.group(0),
                "time": time_match.group(0) if time_match else "",
                "city": _join_words(columns[4]),
                "start": start,
                "end": end,
                "route": f"{start} - {end}",
                "amount": f"{float(amount_match.group(1).replace(',', '')):.2f}",
            })
    return trips


def parse_itinerary(pdf_path):
    path = Path(pdf_path)
    document = fitz.open(path)
    try:
        text = "\n".join(page.get_text() for page in document)
        if not is_itinerary_text(text):
            raise ValueError("不是网约车行程单")
        trips = []
        for page in document:
            trips.extend(_extract_page_trips(page))
    finally:
        document.close()

    compact = _compact(text)
    total_match = re.search(r"合计([\d,]+(?:\.\d{1,2})?)元", compact)
    application_match = re.search(r"申请时间[::]?(20\d{2}-\d{2}-\d{2})", compact)
    amount = f"{float(total_match.group(1).replace(',', '')):.2f}" if total_match else ""
    if not amount and trips:
        amount = f"{sum(float(trip['amount']) for trip in trips):.2f}"
    dates = list(dict.fromkeys(trip["date"] for trip in trips))
    routes = [trip["route"] for trip in trips]
    return {
        "category": "行程单",
        "filename": path.name,
        "date": "、".join(dates),
        "application_date": application_match.group(1) if application_match else "",
        "route": "；".join(routes),
        "amount": amount,
        "invoice_no": "未识别",
        "kind": "行程单",
        "hotel": "未识别",
        "trips": trips,
    }
