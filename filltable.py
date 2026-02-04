import fitz  # PyMuPDF，项目中已有的库
import re
import json
import os
import random
import datetime
import calendar

CONFIG_FILE = 'route_config.json'

# 默认配置
DEFAULT_ROUTES = [
    {"start": "公司", "end": "市建委", "price": 30.0},
    {"start": "公司", "end": "省建设厅", "price": 60.0},
    {"start": "公司", "end": "西三环", "price": 10.0}
]

class ReimbursementLogic:
    def __init__(self):
        self.routes = self.load_config()

    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except:
                pass
        return DEFAULT_ROUTES

    def save_config(self, routes):
        self.routes = routes
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(routes, f, ensure_ascii=False, indent=2)

    def extract_invoice_info(self, pdf_path):
        """
        提取发票信息（金额和发票号）
        使用 fitz (PyMuPDF)
        """
        try:
            doc = fitz.open(pdf_path)
            
            if doc.page_count == 0:
                doc.close()
                return None, None
            
            page = doc[0]
            text = page.get_text()
            
            amount = None
            invoice_no = None
            
            try:
                blocks = page.get_text("dict")["blocks"]
                
                # 提取金额（查找"小写"后面的金额）
                keyword_rect = None
                for block in blocks:
                    if "lines" in block:
                        for line in block["lines"]:
                            for span in line["spans"]:
                                if "小写" in span["text"]:
                                    keyword_rect = span["bbox"]
                                    break
                            if keyword_rect:
                                break
                    if keyword_rect:
                        break
                
                if keyword_rect:
                    kw_y0, kw_y1 = keyword_rect[1], keyword_rect[3]
                    kw_x1 = keyword_rect[2]
                    
                    candidates = []
                    for block in blocks:
                        if "lines" in block:
                            for line in block["lines"]:
                                for span in line["spans"]:
                                    span_bbox = span["bbox"]
                                    span_y0, span_y1 = span_bbox[1], span_bbox[3]
                                    span_x0 = span_bbox[0]
                                    span_text = span["text"].strip()
                                    
                                    y_overlap = not (span_y1 < kw_y0 or span_y0 > kw_y1)
                                    is_right = span_x0 >= kw_x1 - 10
                                    
                                    if y_overlap and is_right and ('¥' in span_text or '￥' in span_text):
                                        amount_match = re.search(r'[¥￥]\s*([\d,]+\.?\d*)', span_text)
                                        if amount_match:
                                            amount_str = amount_match.group(1).replace(',', '')
                                            try:
                                                amount_val = float(amount_str)
                                                distance = span_x0 - kw_x1
                                                candidates.append((distance, amount_val))
                                            except:
                                                pass
                    
                    if candidates:
                        candidates.sort(key=lambda x: x[0])
                        amount = candidates[0][1]
                
                # 提取发票号（查找"发票号码"）
                if '发票号码' in text:
                    idx = text.find('发票号码')
                    after_text = text[idx:idx+200]
                    match = re.search(r'(\d{18,20})', after_text)
                    if match:
                        invoice_no = match.group(1)
                
                # 备用：查找20位数字
                if not invoice_no:
                    matches_20 = re.findall(r'\b(\d{20})\b', text)
                    if matches_20:
                        invoice_no = matches_20[0]
            
            except Exception as e:
                print(f"坐标定位失败: {e}")
            
            # 备用方案
            if amount is None:
                m = re.search(r'小写[：:\s]*[¥￥]\s*([\d,]+\.?\d*)', text)
                if m:
                    amount = float(m.group(1).replace(',', ''))
            
            doc.close()
            return amount, invoice_no
            
        except Exception as e:
            print(f"解析发票失败 {pdf_path}: {e}")
            return None, None
        """
        提取发票金额（使用坐标定位，查找"小写"后面的金额）
        使用 fitz (PyMuPDF)
        """
        try:
            doc = fitz.open(pdf_path)
            
            # 只处理第一页（发票通常是单页）
            if doc.page_count == 0:
                doc.close()
                return None
            
            page = doc[0]
            text = page.get_text()
            
            amount = None
            
            try:
                # 获取页面上所有文本块及其坐标
                blocks = page.get_text("dict")["blocks"]
                
                # 查找"小写"关键字的位置
                keyword_rect = None
                for block in blocks:
                    if "lines" in block:
                        for line in block["lines"]:
                            for span in line["spans"]:
                                text_content = span["text"]
                                if "小写" in text_content:
                                    keyword_rect = span["bbox"]  # (x0, y0, x1, y1)
                                    break
                            if keyword_rect:
                                break
                    if keyword_rect:
                        break
                
                # 如果找到了"小写"，查找同一行右侧的金额
                if keyword_rect:
                    kw_y0, kw_y1 = keyword_rect[1], keyword_rect[3]
                    kw_x1 = keyword_rect[2]  # 关键字右边界
                    
                    # 查找与"小写"在同一水平线上的金额
                    candidates = []
                    for block in blocks:
                        if "lines" in block:
                            for line in block["lines"]:
                                for span in line["spans"]:
                                    span_bbox = span["bbox"]
                                    span_y0, span_y1 = span_bbox[1], span_bbox[3]
                                    span_x0 = span_bbox[0]
                                    span_text = span["text"].strip()
                                    
                                    # 判断是否在同一行（Y坐标接近）
                                    y_overlap = not (span_y1 < kw_y0 or span_y0 > kw_y1)
                                    # 在关键字右侧
                                    is_right = span_x0 >= kw_x1 - 10  # 允许10像素误差
                                    
                                    # 检查是否包含金额（¥或￥符号）
                                    if y_overlap and is_right and ('¥' in span_text or '￥' in span_text):
                                        # 提取数字部分
                                        amount_match = re.search(r'[¥￥]\s*([\d,]+\.?\d*)', span_text)
                                        if amount_match:
                                            amount_str = amount_match.group(1).replace(',', '')
                                            try:
                                                amount_val = float(amount_str)
                                                distance = span_x0 - kw_x1
                                                candidates.append((distance, amount_val))
                                            except:
                                                pass
                    
                    # 选择距离最近的金额
                    if candidates:
                        candidates.sort(key=lambda x: x[0])
                        amount = candidates[0][1]
            
            except Exception as e:
                print(f"坐标定位失败: {e}")
            
            # 备用方案：如果坐标定位失败，使用正则匹配
            if amount is None:
                # 优先匹配"小写"后面的金额
                m = re.search(r'小写[：:\s]*[¥￥]\s*([\d,]+\.?\d*)', text)
                if m:
                    amount = float(m.group(1).replace(',', ''))
                else:
                    # 兜底：价税合计
                    m2 = re.search(r'价税合计.*?[¥￥]\s*([\d,]+\.?\d*)', text, re.S)
                    if m2:
                        amount = float(m2.group(1).replace(',', ''))
            
            doc.close()
            return amount
            
        except Exception as e:
            print(f"解析金额失败 {pdf_path}: {e}")
            return None

    def parse_date_range(self, date_str):
        """解析 '2025年7-12月' 格式"""
        try:
            # 尝试匹配 2025年7-12月
            m = re.match(r'(\d{4})[年.-](\d{1,2})[-至到](\d{1,2})月?', date_str.strip())
            if m:
                year = int(m.group(1))
                start_month = int(m.group(2))
                end_month = int(m.group(3))
            else:
                # 尝试匹配单月 2025年7月
                m2 = re.match(r'(\d{4})[年.-](\d{1,2})月?', date_str.strip())
                if m2:
                    year = int(m2.group(1))
                    start_month = int(m2.group(2))
                    end_month = start_month
                else:
                    return None, None

            start_date = datetime.date(year, start_month, 1)
            
            # 获取结束月份的最后一天
            _, last_day = calendar.monthrange(year, end_month)
            end_date = datetime.date(year, end_month, last_day)
            
            return start_date, end_date
        except:
            return None, None

    def get_workdays(self, start_date, end_date):
        """获取范围内所有工作日"""
        workdays = []
        curr = start_date
        while curr <= end_date:
            # 0-4 是周一到周五
            if curr.weekday() < 5:
                workdays.append(curr)
            curr += datetime.timedelta(days=1)
        return workdays

    def match_route(self, amount):
        """根据金额找最接近的路线"""
        if not self.routes:
            return "公司", "未知目的地"
        
        # 按差值排序，取绝对值最小的
        sorted_routes = sorted(self.routes, key=lambda x: abs(float(x['price']) - amount))
        best = sorted_routes[0]
        return best['start'], best['end']

    def calculate_people(self, amount):
        """根据金额计算人数逻辑"""
        # 逻辑：<30: 1人, 30-100: 2人, >100: 3人
        if amount <= 30:
            return 1
        elif amount <= 100:
            return 2
        else:
            return 3

    def generate_data(self, file_paths, date_range_str):
        start_date, end_date = self.parse_date_range(date_range_str)
        if not start_date:
            return {'success': False, 'error': '日期格式错误，请使用如 "2025年7-12月" 或 "2025年7月" 的格式'}

        workdays = self.get_workdays(start_date, end_date)
        if not workdays:
            return {'success': False, 'error': '该时间段内没有工作日'}

        results = []
        invoice_dict = {}  # 用于去重：{invoice_no: [file_paths]}
        
        for path in file_paths:
            if not os.path.exists(path): continue
            
            amount, invoice_no = self.extract_invoice_info(path)
            if amount is None:
                continue 
            
            # 检查发票号重复
            is_duplicate = False
            if invoice_no:
                if invoice_no in invoice_dict:
                    invoice_dict[invoice_no].append(os.path.basename(path))
                    is_duplicate = True
                    continue  # 跳过重复的发票
                else:
                    invoice_dict[invoice_no] = [os.path.basename(path)]
            
            start_loc, end_loc = self.match_route(amount)
            people = self.calculate_people(amount)
            
            results.append({
                'amount': amount,
                'start': start_loc,
                'end': end_loc,
                'people': people,
                'path': path,
                'invoiceNo': invoice_no or '未识别'
            })

        if not results:
            return {'success': False, 'error': '未提取到有效发票金额，请检查发票清晰度或类型'}

        # 简单的随机分配逻辑
        results.sort(key=lambda x: x['amount'])
        
        final_rows = []
        for i, item in enumerate(results):
            # 随机选一个工作日
            r_date = random.choice(workdays)
            item['date_obj'] = r_date
            item['date_str'] = r_date.strftime('%Y-%m-%d')
            final_rows.append(item)

        # 按日期由早到晚排序
        final_rows.sort(key=lambda x: x['date_obj'])

        # 构建返回给前端的表格数据
        table_rows = []
        for idx, row in enumerate(final_rows):
            table_rows.append({
                'id': idx + 1,
                'people': row['people'],
                'date': row['date_str'],
                'start': row['start'],
                'end': row['end'],
                'amount': row['amount'],
                'invoiceNo': row['invoiceNo']
            })
        
        # 检查重复详情
        duplicate_details = []
        for inv_no, files in invoice_dict.items():
            if len(files) > 1:
                duplicate_details.append({
                    'invoiceNo': inv_no,
                    'count': len(files),
                    'files': files
                })

        return {
            'success': True, 
            'rows': table_rows,
            'duplicateDetails': duplicate_details
        }

# 单例
logic = ReimbursementLogic()
