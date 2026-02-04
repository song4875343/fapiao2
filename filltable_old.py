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

    def extract_amount(self, pdf_path):
        """
        提取发票金额
        使用 fitz (PyMuPDF) 替代 pdfplumber
        """
        try:
            doc = fitz.open(pdf_path)
            # 提取所有页面的文本
            text = "".join([page.get_text() for page in doc])
            doc.close()
            
            # 简单的关键词过滤，确保是交通类发票 (可选，如果不需要过滤可注释掉)
            # if "客运" not in text and "运输" not in text and "通行费" not in text and "车" not in text:
            #     return None

            # 1. 精确匹配：价税合计...小写...¥100.00
            # fitz 提取的文本有时会有换行符，使用 re.S (DOTALL) 让 . 匹配换行符
            # 兼容全角符号 '￥' 和半角 '¥'
            m = re.search(r'价税合计.*?小写.*?[￥¥]\s*([\d,]+\.\d{1,2})', text, re.S)
            if m:
                return float(m.group(1).replace(',', ''))

            # 2. 兜底匹配：找文中出现的所有金额，取最后一个
            # 这是一个常见的发票金额提取“土办法”，通常最后一个金额就是总额
            all_money = re.findall(r'[￥¥]\s*([\d,]+\.\d{1,2})', text)
            if all_money:
                return float(all_money[-1].replace(',', ''))
            
            return None
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
        
        for path in file_paths:
            if not os.path.exists(path): continue
            
            amount = self.extract_amount(path)
            if amount is None:
                # 可以在这里记录日志，告诉用户哪些文件没识别出来
                continue 
            
            start_loc, end_loc = self.match_route(amount)
            people = self.calculate_people(amount)
            
            results.append({
                'amount': amount,
                'start': start_loc,
                'end': end_loc,
                'people': people,
                'path': path
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
                'amount': row['amount']
            })

        return {'success': True, 'rows': table_rows}

# 单例
logic = ReimbursementLogic()