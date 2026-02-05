import os
import json
import base64
import uuid
import time
from io import BytesIO
import PyPDF2
from PyPDF2 import PdfWriter, PdfReader
import webview
import traceback
import sys
import filltable
import csv
# HTML界面
import ui

# 增加最大递归深度
sys.setrecursionlimit(2000)

class PDFMergerAPI:
    def __init__(self):
        self.source_files = {} 

    def select_pdfs(self):
        """选择PDF文件"""
        file_types = ('PDF Files (*.pdf)', 'All Files (*.*)')
        files = window.create_file_dialog(
            webview.FileDialog.OPEN,
            allow_multiple=True,
            file_types=file_types
        )
        
        result = []
        if files:
            for file_path in files:
                file_path = str(file_path)
                if not os.path.exists(file_path): continue
                try:
                    import fitz # PyMuPDF
                    doc = fitz.open(file_path)
                    page_count = doc.page_count
                    doc.close()
                    
                    result.append({
                        'path': file_path,
                        'name': os.path.basename(file_path),
                        'page_count': page_count
                    })
                except Exception as e:
                    print(f"解析文件失败 {file_path}: {e}")
        return result

    def get_file_info(self, file_path):
        """获取单个文件的信息"""
        try:
            file_path = str(file_path)
            if not os.path.exists(file_path):
                return {'success': False, 'error': '文件不存在'}
            
            import fitz
            doc = fitz.open(file_path)
            count = doc.page_count
            doc.close()
            return {'success': True, 'page_count': count, 'path': file_path}
        except Exception as e:
            return {'success': False, 'error': str(e)}

    def get_page_image(self, file_path, page_index, quality=1.0):
        """获取页面图片，增加智能清晰度控制"""
        try:
            file_path = str(file_path)
            if not os.path.exists(file_path):
                return {'success': False, 'error': '文件不存在'}

            import fitz
            doc = fitz.open(file_path)
            if page_index >= doc.page_count:
                return {'success': False, 'error': '页码越界'}

            page = doc[page_index]
            requested_zoom = float(quality)
            
            # --- 智能清晰度控制 ---
            rect = page.rect
            origin_w, origin_h = rect.width, rect.height
            
            target_long_edge = max(origin_w, origin_h) * requested_zoom
            
            # 放宽限制到 4096 (4K)
            MAX_LONG_EDGE = 4096 
            
            final_zoom = requested_zoom
            if target_long_edge > MAX_LONG_EDGE:
                final_zoom = MAX_LONG_EDGE / max(origin_w, origin_h)
            
            # 极端保护
            if (origin_w * final_zoom > 15000) or (origin_h * final_zoom > 15000):
                 scale_limit = 15000 / max(origin_w, origin_h)
                 final_zoom = min(final_zoom, scale_limit)
            
            matrix = fitz.Matrix(final_zoom, final_zoom)
            pix = page.get_pixmap(matrix=matrix, alpha=False)
            img_data = pix.tobytes("png")
            doc.close()

            img_base64 = base64.b64encode(img_data).decode('utf-8')
            return {'success': True, 'image': f'data:image/png;base64,{img_base64}'}
            
        except Exception as e:
            return {'success': False, 'error': str(e)}

    def merge_pages(self, page_list, output_path, mode='normal'):
        """合并逻辑"""
        try:
            if not page_list:
                return {'success': False, 'error': '没有可合并的页面'}

            if isinstance(output_path, (list, tuple)):
                output_path = str(output_path[0])
            else:
                output_path = str(output_path)
            
            output_path = output_path.strip().strip('"').strip("'")
            output_dir = os.path.dirname(output_path)
            if output_dir and not os.path.exists(output_dir):
                os.makedirs(output_dir)

            success_msg = '合并成功'
            
            if mode == 'normal':
                writer = PdfWriter()
                open_files = {} 
                try:
                    for item in page_list:
                        path = str(item['path'])
                        page_idx = int(item['page_index'])
                        rotation = int(item.get('rotation', 0))
                        
                        if path not in open_files:
                            if not os.path.exists(path): continue
                            open_files[path] = PdfReader(path)
                        
                        reader = open_files[path]
                        if 0 <= page_idx < len(reader.pages):
                            page = reader.pages[page_idx]
                            if rotation != 0:
                                page.rotate(rotation)
                            writer.add_page(page)

                    with open(output_path, 'wb') as f:
                        writer.write(f)
                except Exception as e:
                    raise e
            
            elif mode == 'invoice':
                self._merge_invoice_by_pages(page_list, output_path)
                success_msg = '发票合并成功'

            # --- 关键修改：合并完成后，立即生成第一页缩略图 ---
            # 这样前端就能实现“秒开”合并记录，无需等待
            first_page_thumb = None
            try:
                import fitz
                doc = fitz.open(output_path)
                if doc.page_count > 0:
                    # 生成0.5倍率的缩略图
                    page = doc[0]
                    matrix = fitz.Matrix(0.5, 0.5)
                    pix = page.get_pixmap(matrix=matrix, alpha=False)
                    img_data = pix.tobytes("png")
                    b64 = base64.b64encode(img_data).decode('utf-8')
                    first_page_thumb = f'data:image/png;base64,{b64}'
                doc.close()
            except Exception as e:
                print(f"缩略图生成失败: {e}")

            return {
                'success': True, 
                'message': success_msg, 
                'output_path': output_path,
                'thumbnail': first_page_thumb # 返回缩略图
            }

        except Exception as e:
            return {'success': False, 'error': str(e)}

    def _merge_invoice_by_pages(self, page_list, output_path):
        """发票合并逻辑"""
        import fitz
        output_doc = fitz.open()
        a4_width, a4_height = fitz.paper_size("a4")
        
        for i in range(0, len(page_list), 2):
            new_page = output_doc.new_page(width=a4_width, height=a4_height)
            if i < len(page_list):
                self._place_page_on_canvas(output_doc, new_page, page_list[i], a4_width, a4_height, True)
            if i + 1 < len(page_list):
                self._place_page_on_canvas(output_doc, new_page, page_list[i+1], a4_width, a4_height, False)

        output_doc.save(output_path)
        output_doc.close()

    def _place_page_on_canvas(self, output_doc, target_page, page_info, a4_width, a4_height, is_top):
        import fitz
        path = str(page_info['path'])
        page_idx = int(page_info['page_index'])
        rotation = int(page_info.get('rotation', 0))

        if not os.path.exists(path): return

        try:
            src_doc = fitz.open(path)
            src_page = src_doc[page_idx]
            
            rect = src_page.rect
            if rotation % 180 != 0:
                pdf_width = rect.height
                pdf_height = rect.width
            else:
                pdf_width = rect.width
                pdf_height = rect.height

            scale_x = a4_width / pdf_width
            scale_y = (a4_height / 2) / pdf_height
            scale = min(scale_x, scale_y) * 0.95

            scaled_width = pdf_width * scale
            scaled_height = pdf_height * scale
            x_offset = (a4_width - scaled_width) / 2
            
            if is_top:
                y_offset = (a4_height / 2 - scaled_height) / 2 
            else:
                y_offset = (a4_height / 2) + (a4_height / 2 - scaled_height) / 2

            target_page.show_pdf_page(
                fitz.Rect(x_offset, y_offset, x_offset + scaled_width, y_offset + scaled_height),
                src_doc, 
                page_idx,
                rotate=rotation
            )
            src_doc.close()
        except Exception as e:
            print(f"处理页面出错: {e}")

    def save_file_dialog(self):
        file_types = ('PDF Files (*.pdf)', 'All Files (*.*)')
        result = window.create_file_dialog(webview.FileDialog.SAVE, file_types=file_types, save_filename='merged.pdf')
        if result:
            if isinstance(result, (list, tuple)):
                if len(result) > 0: return str(result[0])
                return None
            return str(result)
        return None
    
    def clear_files(self):
        return True

# --- 新增：报销单相关 API ---
    def get_routes(self):
        """获取路线配置"""
        return filltable.logic.load_config()

    def save_routes(self, routes):
        """保存路线配置"""
        filltable.logic.save_config(routes)
        return {'success': True}

    def generate_reimbursement_form(self, file_paths, date_range):
        """生成报销单数据"""
        # file_paths 可能是字典列表(前端传来的pages)或者纯路径列表
        # 这里我们需要纯路径列表，且去重
        paths = []
        if file_paths and len(file_paths) > 0:
            if isinstance(file_paths[0], dict):
                seen = set()
                for p in file_paths:
                    if p['path'] not in seen:
                        paths.append(p['path'])
                        seen.add(p['path'])
            else:
                paths = list(set(file_paths)) # 也是为了去重
        
        return filltable.logic.generate_data(paths, date_range)

    def save_csv_dialog(self):
        """打开保存CSV对话框"""
        file_types = ('CSV Files (*.csv)', 'All Files (*.*)')
        result = window.create_file_dialog(webview.FileDialog.SAVE, file_types=file_types, save_filename='报销单.csv')
        if result:
            if isinstance(result, (list, tuple)):
                if len(result) > 0: return str(result[0])
                return None
            return str(result)
        return None

    def save_csv_data(self, path, rows):
        """保存数据到CSV文件（使用标准库）"""
        try:
            # 使用 utf-8-sig 编码，这样Excel打开中文才不会乱码
            with open(path, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f)
                # 写入表头
                writer.writerow(['序号', '人数', '日期', '起点', '终点', '票额'])
                # 写入数据
                total = 0.0
                for r in rows:
                    writer.writerow([r['id'], r['people'], r['date'], r['start'], r['end'], r['amount']])
                    total += float(r['amount'])
                
                # 写入合计行
                writer.writerow(['', '', '', '', '合计', f'{total:.2f}'])
            return {'success': True}
        except Exception as e:
            return {'success': False, 'error': str(e)}
    
    def save_reimbursement_csv(self, path, rows):
        """保存报销单到CSV文件"""
        try:
            with open(path, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f)
                # 写入表头（增加来源列）
                writer.writerow(['序号', '发票号', '来源', '人数', '日期', '起点', '终点', '票额'])
                # 写入数据
                total = 0.0
                for r in rows:
                    writer.writerow([r['id'], r.get('invoiceNo', '未识别'), r.get('source', ''), r['people'], r['date'], r['start'], r['end'], r['amount']])
                    total += float(r['amount'])
                
                # 写入合计行
                writer.writerow(['', '', '', '', '', '', '合计', f'{total:.2f}'])
            return {'success': True}
        except Exception as e:
            return {'success': False, 'error': str(e)}
    
    def save_statistics_csv(self, path, amounts):
        """保存统计数据到CSV文件"""
        try:
            with open(path, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f)
                # 写入表头
                writer.writerow(['序号', '发票号', '来源页面', '金额'])
                # 写入数据
                total = 0.0
                for idx, item in enumerate(amounts):
                    pages_str = ', '.join(item.get('pages', []))
                    writer.writerow([idx + 1, item.get('invoiceNo', '未识别'), pages_str, item['amount']])
                    total += float(item['amount'])
                
                # 写入合计行
                writer.writerow(['', '', '合计（已去重）', f'{total:.2f}'])
            return {'success': True}
        except Exception as e:
            return {'success': False, 'error': str(e)}
    
    def print_pdf(self, file_path):
        """打印PDF文件 - 读取文件内容并转换为Base64返回"""
        try:
            file_path = str(file_path)
            if not os.path.exists(file_path):
                return {'success': False, 'error': '文件不存在'}
            
            # 读取文件并转换为Base64
            # 这种方式比 file:// 协议更稳定，不会被浏览器安全策略拦截
            with open(file_path, "rb") as pdf_file:
                encoded_string = base64.b64encode(pdf_file.read()).decode('utf-8')
            
            return {
                'success': True, 
                'data': f'data:application/pdf;base64,{encoded_string}',  # 返回完整的数据URI
                'message': '准备打印'
            }
        
        except Exception as e:
            return {'success': False, 'error': str(e)}

    def calculate_invoice_amounts(self, pages_info):
        """从电子发票PDF中提取金额信息（按发票号去重，检测重复）"""
        try:
            import re
            import fitz
            
            # 用于去重的字典：{invoice_no: {'amount': xxx, 'pages': [...], 'fileName': xxx, 'pageIndex': xxx}}
            invoice_dict = {}
            unrecognized_invoices = []  # 未识别到发票号的
            
            # 校验：统计选中的文件数量
            selected_files = set()
            for page_info in pages_info:
                selected_files.add(page_info['path'])
            
            for idx, page_info in enumerate(pages_info):
                path = page_info['path']
                page_index = page_info['pageIndex']
                file_name = page_info['fileName']
                page_label = f"{file_name}-P{page_index+1}"
                
                if not os.path.exists(path):
                    unrecognized_invoices.append({
                        'amount': '0.00',
                        'invoiceNo': '文件不存在',
                        'pages': [page_label],
                        'isDuplicate': False,
                        'fileName': file_name,
                        'pageIndex': page_index
                    })
                    continue
                
                try:
                    doc = fitz.open(path)
                    
                    if page_index >= doc.page_count:
                        unrecognized_invoices.append({
                            'amount': '0.00',
                            'invoiceNo': '页码越界',
                            'pages': [page_label],
                            'isDuplicate': False,
                            'fileName': file_name,
                            'pageIndex': page_index
                        })
                        doc.close()
                        continue
                    
                    page = doc[page_index]
                    text = page.get_text()
                    
                    # 提取发票号码（使用坐标定位，更准确）
                    invoice_no = None
                    
                    try:
                        # 获取页面上所有文本块及其坐标
                        blocks = page.get_text("dict")["blocks"]
                        
                        # 查找"发票号码"关键字的位置
                        keyword_rect = None
                        for block in blocks:
                            if "lines" in block:
                                for line in block["lines"]:
                                    for span in line["spans"]:
                                        text_content = span["text"]
                                        if "发票号码" in text_content:
                                            keyword_rect = span["bbox"]  # (x0, y0, x1, y1)
                                            break
                                    if keyword_rect:
                                        break
                            if keyword_rect:
                                break
                        
                        # 如果找到了"发票号码"，查找同一行右侧的数字
                        if keyword_rect:
                            kw_y0, kw_y1 = keyword_rect[1], keyword_rect[3]
                            kw_x1 = keyword_rect[2]  # 关键字右边界
                            
                            # 查找与"发票号码"在同一水平线上的文本
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
                                            # 是数字
                                            is_number = span_text.isdigit() and len(span_text) >= 8
                                            
                                            if y_overlap and is_right and is_number:
                                                # 记录候选项：(距离, 文本)
                                                distance = span_x0 - kw_x1
                                                candidates.append((distance, span_text))
                            
                            # 选择距离最近的数字
                            if candidates:
                                candidates.sort(key=lambda x: x[0])
                                invoice_no = candidates[0][1]
                    
                    except Exception as e:
                        print(f"坐标定位失败: {e}")
                    
                    # 备用方案：如果坐标定位失败，使用简单的数字查找
                    if not invoice_no:
                        # 查找20位数字
                        matches_20 = re.findall(r'\b(\d{20})\b', text)
                        if matches_20:
                            invoice_no = matches_20[0]
                        # 查找18-19位数字
                        elif not invoice_no:
                            matches_18_19 = re.findall(r'\b(\d{18,19})\b', text)
                            if matches_18_19:
                                invoice_no = matches_18_19[0]
                    
                    # 提取金额（使用坐标定位，查找"小写"后面的金额）
                    amount = '0.00'
                    try:
                        # 查找"小写"关键字的位置
                        amount_keyword_rect = None
                        for block in blocks:
                            if "lines" in block:
                                for line in block["lines"]:
                                    for span in line["spans"]:
                                        text_content = span["text"]
                                        if "小写" in text_content:
                                            amount_keyword_rect = span["bbox"]
                                            break
                                    if amount_keyword_rect:
                                        break
                            if amount_keyword_rect:
                                break
                        
                        # 如果找到了"小写"，查找同一行右侧的金额
                        if amount_keyword_rect:
                            kw_y0, kw_y1 = amount_keyword_rect[1], amount_keyword_rect[3]
                            kw_x1 = amount_keyword_rect[2]
                            
                            # 收集同一行的所有文本块
                            same_line_spans = []
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
                                            
                                            if y_overlap and is_right and span_text:
                                                same_line_spans.append((span_x0, span_text))
                            
                            # 按X坐标排序，拼接文本
                            same_line_spans.sort(key=lambda x: x[0])
                            combined_text = ' '.join([text for _, text in same_line_spans])
                            
                            # 从拼接后的文本中提取金额
                            amount_match = re.search(r'[¥￥]\s*([\d,]+\.?\d*)', combined_text)
                            if amount_match:
                                try:
                                    amount = float(amount_match.group(1).replace(',', ''))
                                except:
                                    pass
                    
                    except Exception as e:
                        print(f"金额坐标定位失败: {e}")
                    
                    # 备用方案：如果坐标定位失败，使用正则匹配
                    if amount == '0.00':
                        amount_patterns = [
                            r'小写[：:\s]*[¥￥]\s*([\d,]+\.?\d*)',
                            r'价税合计[：:\s]*[¥￥]?\s*([\d,]+\.?\d*)',
                            r'合计[：:\s]*[¥￥]?\s*([\d,]+\.?\d*)',
                        ]
                        
                        for pattern in amount_patterns:
                            matches = re.findall(pattern, text)
                            if matches:
                                amounts = [float(m.replace(',', '')) for m in matches if m]
                                if amounts:
                                    amount = f'{max(amounts):.2f}'
                                    break
                    
                    # 根据发票号去重
                    if invoice_no:
                        if invoice_no in invoice_dict:
                            # 重复发票，只记录页面位置
                            invoice_dict[invoice_no]['pages'].append(page_label)
                            invoice_dict[invoice_no]['isDuplicate'] = True
                        else:
                            # 首次出现
                            invoice_dict[invoice_no] = {
                                'amount': amount,
                                'invoiceNo': invoice_no,
                                'pages': [page_label],
                                'isDuplicate': False,
                                'fileName': file_name,
                                'pageIndex': page_index
                            }
                    else:
                        # 未识别到发票号
                        unrecognized_invoices.append({
                            'amount': amount,
                            'invoiceNo': '未识别',
                            'pages': [page_label],
                            'isDuplicate': False,
                            'fileName': file_name,
                            'pageIndex': page_index
                        })
                    
                    doc.close()
                    
                except Exception as e:
                    unrecognized_invoices.append({
                        'amount': '0.00',
                        'invoiceNo': f'错误: {str(e)}',
                        'pages': [page_label],
                        'isDuplicate': False,
                        'fileName': file_name,
                        'pageIndex': page_index
                    })
            
            # 合并结果（去重后的发票 + 未识别的）
            results = list(invoice_dict.values()) + unrecognized_invoices
            
            # 计算去重后的总金额
            total_amount = sum(float(item['amount']) for item in results)
            
            # 统计重复详情
            duplicate_details = []
            for item in invoice_dict.values():
                if item['isDuplicate']:
                    duplicate_details.append({
                        'invoiceNo': item['invoiceNo'],
                        'amount': item['amount'],
                        'count': len(item['pages']),
                        'pages': item['pages']
                    })
            
            # 校验：检查识别数量与选中文件数量
            recognized_count = len(results)
            selected_count = len(pages_info)
            validation_warning = None
            unrecognized_files = []
            
            if recognized_count < selected_count:
                # 找出哪些文件/页面未被正确识别
                for page_info in pages_info:
                    page_label = f"{page_info['fileName']}-P{page_info['pageIndex']+1}"
                    found = False
                    for item in results:
                        if page_label in item['pages']:
                            found = True
                            break
                    if not found:
                        unrecognized_files.append(page_label)
                
                validation_warning = f"警告：选中了{selected_count}个页面，但只识别出{recognized_count}个发票。"
            
            return {
                'success': True,
                'amounts': results,
                'totalAmount': f'{total_amount:.2f}',
                'duplicateDetails': duplicate_details,
                'validationWarning': validation_warning,
                'unrecognizedFiles': unrecognized_files,
                'selectedCount': selected_count,
                'recognizedCount': recognized_count
            }
            
        except Exception as e:
            return {'success': False, 'error': str(e)}


if __name__ == '__main__':
    api = PDFMergerAPI()
    
    # --- 关键修改：将 HTML 写入临时文件，赋予页面合法的 file:// Origin ---
    # 1. 获取当前运行目录
    base_dir = os.path.dirname(os.path.abspath(__file__))
    
    # 2. 定义临时HTML文件路径
    temp_html_path = os.path.join(base_dir, 'app_gui_temp.html')
    
    # 3. 将 ui.py 中的 HTML 字符串写入文件
    # 这样浏览器就会以 file:// 协议加载它，从而拥有合法的 Origin
    with open(temp_html_path, 'w', encoding='utf-8') as f:
        f.write(ui.html_content)
    
    # 4. 创建窗口时，使用 url 参数加载本地文件，而不是 html 参数
    window = webview.create_window(
        '发票打印工具 - 专业版 v3.3',
        url=temp_html_path,  # 改为 url，加载本地文件
        width=1200,
        height=800,
        resizable=True,
        js_api=api
    )
    
    # 退出时尝试清理临时文件
    def on_closed():
        try:
            if os.path.exists(temp_html_path):
                os.remove(temp_html_path)
        except:
            pass
    
    window.events.closed += on_closed
    
    webview.start(debug=False)  # 开启调试模式，可以右键检查元素查看控制台