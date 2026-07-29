"""普通电子发票的票面识别与字段提取。"""

import re
import unicodedata


INVOICE_TITLE_WORDS = ("普通发票", "增值税专用发票", "增值税电子专用发票")
HOTEL_WORDS = ("住宿", "酒店", "宾馆", "旅馆", "客房", "房费", "入住", "代订住宿")
TAXI_WORDS = ("出租", "的士", "网约车", "打车", "客运服务", "旅客运输服务", "滴滴", "曹操出行", "高德打车")


def normalize_text(text):
    """统一兼容字形和空白，但保留换行供组织名称提取。"""
    text = unicodedata.normalize("NFKC", text or "")
    text = text.replace("\u00a0", " ")
    return "\n".join(re.sub(r"[ \t]+", " ", line).strip() for line in text.splitlines())


def compact_text(text):
    return re.sub(r"\s+", "", normalize_text(text))


def is_ordinary_invoice_text(text):
    compact = compact_text(text)
    has_title = "电子发票" in compact and any(word in compact for word in INVOICE_TITLE_WORDS)
    has_structure = all(word in compact for word in ("发票号码", "开票日期", "购买方", "销售方", "价税合计"))
    return has_title and has_structure


def extract_invoice_no(text):
    normalized = normalize_text(text)
    compact = re.sub(r"\s+", "", normalized)
    direct = re.search(r"发票号码[::]?([0-9]{20})", compact)
    if direct:
        return direct.group(1)
    candidates = re.findall(r"(?<![0-9])[0-9]{20}(?![0-9])", normalized)
    return candidates[0] if candidates else "未识别"


def extract_date(text):
    compact = compact_text(text)
    labeled = re.search(r"开票日期[::]?(20\d{2})年?(\d{1,2})月?(\d{1,2})日?", compact)
    match = labeled or re.search(r"(20\d{2})年(\d{1,2})月(\d{1,2})日", compact)
    if not match:
        return ""
    year, month, day = map(int, match.groups())
    return f"{year:04d}-{month:02d}-{day:02d}"


def extract_amount(text):
    normalized = normalize_text(text)
    # 电子发票文本层常把“小写”标签排在税前合计之前；价税合计是票面
    # 所有非负人民币金额中的最大值，比依赖视觉顺序更稳定。
    values = []
    currency_values = re.findall(r"[¥￥]\s*([\d,]+(?:\.\d{1,2})?)", normalized)
    # 部分数电发票的 PDF 文本层与视觉顺序相反，会抽取成“79.99\n¥”。
    currency_values += re.findall(
        r"(?<![0-9A-Za-z])([\d,]+(?:\.\d{1,2})?)\s*[¥￥]", normalized
    )
    for value in currency_values:
        digits = value.replace(",", "").split(".", 1)[0]
        # 文本层可能把发票号码排在前一个税额的货币符号后面。
        if len(digits) >= 16:
            continue
        values.append(float(value.replace(",", "")))
    return f"{max(values):.2f}" if values else ""


def extract_kinds(text):
    normalized = normalize_text(text)
    kinds = []
    for value in re.findall(r"\*([^*\n]{1,30})\*", normalized):
        value = value.strip()
        if value and value not in kinds:
            kinds.append(value)
    return "、".join(kinds) if kinds else "其他"


def _organization_candidates(text):
    candidates = []
    suffixes = r"(?:有限责任公司|股份有限公司|有限公司|分公司|分店|酒店|宾馆|旅馆|大厦)"
    for raw_line in normalize_text(text).splitlines():
        line = re.sub(r"^(?:名称|销售方名称|购买方名称)[::]?", "", raw_line).strip(" ::")
        for match in re.finditer(rf"[\u4e00-\u9fffA-Za-z0-9()()·-]{{2,80}}?{suffixes}", line):
            value = match.group(0).strip()
            if value not in candidates:
                candidates.append(value)
    return candidates


def extract_seller(text):
    candidates = _organization_candidates(text)
    return candidates[1] if len(candidates) >= 2 else (candidates[0] if candidates else "未识别")


def extract_hotel(text):
    normalized = normalize_text(text)
    match = re.search(r"酒店名称[::]\s*([^;；\n]{2,80})", normalized)
    return match.group(1).strip() if match else extract_seller(text)


def classify_invoice(text):
    compact = compact_text(text)
    if any(word in compact for word in HOTEL_WORDS):
        return "住宿票"
    if any(word in compact for word in TAXI_WORDS):
        return "出租票"
    return "其他发票"


def parse_invoice(text):
    if not is_ordinary_invoice_text(text):
        raise ValueError("不是普通电子发票票面")
    category = classify_invoice(text)
    return {
        "category": category,
        "date": extract_date(text),
        "amount": extract_amount(text),
        "invoice_no": extract_invoice_no(text),
        "kind": extract_kinds(text),
        "hotel": extract_hotel(text) if category == "住宿票" else "未识别",
        "seller": extract_seller(text),
        "route": "",
    }
