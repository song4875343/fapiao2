import os
import json
import base64
import uuid
import time
from io import BytesIO
import PyPDF2
from PyPDF2 import PdfWriter, PdfReader
try:
    import webview
except ImportError:
    webview = None
import traceback
import sys
import filltable
import train_ticket
import csv
# HTML界面
import ui

# 增加最大递归深度
sys.setrecursionlimit(2000)

class PDFMergerAPI:
    def __init__(self, mode='local'):
        """
        初始化 API
        mode: 'local' 本地模式 (pywebview) | 'server' 服务器模式 (FastAPI)
        """
        self.mode = mode
        self.source_files = {}
        self.file_mapping = {}  # 服务器模式：虚拟ID -> 真实路径
        self.temp_dir = None
        
        if mode == 'server':
            import tempfile
            from pathlib import Path
            self.temp_dir = Path(tempfile.gettempdir()) / "pdfm_server"
            self.temp_dir.mkdir(exist_ok=True)
            print(f"📁 临时文件目录: {self.temp_dir}")

    def select_pdfs(self):
        """选择PDF文件"""
        if self.mode == 'server':
            # 服务器模式：返回特殊标记，让前端触发文件上传
            return 'BROWSER_UPLOAD_REQUIRED'
        
        # 本地模式：使用系统文件对话框
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
    
    def upload_files(self, files_data):
        """
        服务器模式专用：处理上传的文件
        files_data: [(filename, file_content_bytes), ...]
        """
        if self.mode != 'server':
            return {'success': False, 'error': '仅服务器模式可用'}
        
        result = []
        import fitz
        
        for filename, content in files_data:
            if not filename.lower().endswith('.pdf'):
                continue
            
            # 生成虚拟ID
            file_id = str(uuid.uuid4())
            file_path = self.temp_dir / f"{file_id}.pdf"
            
            # 保存文件
            with open(file_path, 'wb') as f:
                f.write(content)
            
            # 记录映射
            self.file_mapping[file_id] = str(file_path)
            
            # 获取页数
            try:
                doc = fitz.open(str(file_path))
                page_count = doc.page_count
                doc.close()
                
                result.append({
                    'path': file_id,  # 返回虚拟ID
                    'name': filename,
                    'page_count': page_count
                })
            except Exception as e:
                print(f"解析文件失败: {e}")
        
        return result

    def get_file_info(self, file_path):
        """获取单个文件的信息"""
        try:
            # 服务器模式：转换虚拟ID为真实路径
            if self.mode == 'server' and file_path in self.file_mapping:
                file_path = self.file_mapping[file_path]
            
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
            # 服务器模式：转换虚拟ID为真实路径
            if self.mode == 'server' and file_path in self.file_mapping:
                file_path = self.file_mapping[file_path]
            
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

            # 服务器模式：转换虚拟ID为真实路径
            if self.mode == 'server':
                for item in page_list:
                    if item['path'] in self.file_mapping:
                        item['path'] = self.file_mapping[item['path']]
                
                # 生成输出文件路径
                output_id = str(uuid.uuid4())
                output_path = self.temp_dir / f"merged_{output_id}.pdf"
                output_path = str(output_path)
            else:
                # 本地模式：处理路径
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

            # 构建基础返回结果
            result = {
                'success': True, 
                'message': success_msg, 
                'output_path': output_path,
                'thumbnail': first_page_thumb
            }
            
            # 服务器模式：不再返回PDF内容，只返回下载ID
            if self.mode == 'server':
                try:
                    # 记录映射（用于后续下载）
                    output_id = os.path.basename(output_path).replace('merged_', '').replace('.pdf', '')
                    self.file_mapping[output_id] = output_path
                    
                    # 获取文件大小
                    file_size = os.path.getsize(output_path)
                    
                    result['output_path'] = output_id
                    result['download_id'] = output_id  # 用于下载的ID
                    result['file_size'] = file_size
                    
                    print(f"✅ 服务器模式：合并完成，文件大小 {file_size} 字节，下载ID: {output_id}")
                except Exception as e:
                    print(f"❌ 处理文件失败: {e}")
                    return {'success': False, 'error': f'处理文件失败: {str(e)}'}
            
            return result

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
        """保存文件对话框"""
        if self.mode == 'server':
            # 服务器模式：返回特殊标记，让前端触发下载
            return 'BROWSER_DOWNLOAD'
        
        # 本地模式：使用系统文件对话框
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
    
    def download_merged_pdf(self, file_id):
        """
        下载合并后的PDF文件（服务器模式专用）
        返回文件路径供FastAPI直接发送文件
        """
        if self.mode != 'server':
            return {'success': False, 'error': '仅服务器模式可用'}
        
        if file_id not in self.file_mapping:
            return {'success': False, 'error': '文件不存在'}
        
        file_path = self.file_mapping[file_id]
        if not os.path.exists(file_path):
            return {'success': False, 'error': '文件已被删除'}
        
        return {'success': True, 'file_path': file_path}

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
                    path = p['path']
                    # 服务器模式：转换虚拟ID为真实路径
                    if self.mode == 'server' and path in self.file_mapping:
                        path = self.file_mapping[path]
                    if path not in seen:
                        paths.append(path)
                        seen.add(path)
            else:
                # 服务器模式：转换虚拟ID为真实路径
                if self.mode == 'server':
                    paths = [self.file_mapping.get(p, p) for p in file_paths]
                else:
                    paths = file_paths
                paths = list(set(paths)) # 去重
        
        return filltable.logic.generate_data(paths, date_range)

    def generate_train_ticket_form(self, pages):
        """提取所选铁路电子客票，并按出行日期、发车时间排序。"""
        resolved_pages = []
        for page in pages or []:
            item = dict(page)
            path = item.get('path')
            if self.mode == 'server' and path in self.file_mapping:
                item['path'] = self.file_mapping[path]
            resolved_pages.append(item)
        return train_ticket.generate_form(resolved_pages)

    def get_page_size_categories(self, pages):
        """Return physical page-size categories for the selected PDF pages."""
        import fitz

        categories = {}
        for item in pages or []:
            path = item.get('path')
            if self.mode == 'server' and path in self.file_mapping:
                path = self.file_mapping[path]
            page_index = int(item.get('pageIndex', item.get('page_index', 0)))
            try:
                doc = fitz.open(path)
                page = doc[page_index]
                short_side = min(page.rect.width, page.rect.height)
                long_side = max(page.rect.width, page.rect.height)
                doc.close()
            except Exception as e:
                return {'success': False, 'error': f"读取页面尺寸失败：{e}"}

            category_id = f"{short_side:.1f}:{long_side:.1f}"
            if category_id not in categories:
                categories[category_id] = {
                    'id': category_id,
                    'shortSide': round(short_side, 3),
                    'longSide': round(long_side, 3),
                    'label': f"{short_side / 72 * 25.4:.0f} × {long_side / 72 * 25.4:.0f} mm",
                    'count': 0,
                }
            categories[category_id]['count'] += 1

        result = sorted(categories.values(), key=lambda item: (-item['count'], item['shortSide'], item['longSide']))
        if not result:
            return {'success': False, 'error': '请先选择页面'}
        return {'success': True, 'categories': result, 'defaultCategoryId': result[0]['id']}

    def normalize_page_sizes(self, pages, target_short_side):
        """Scale selected pages uniformly so their short edge matches the target."""
        import fitz
        import tempfile
        from pathlib import Path

        try:
            target_short_side = float(target_short_side)
            if target_short_side <= 0:
                return {'success': False, 'error': '目标尺寸无效'}

            output_dir = self.temp_dir or (Path(tempfile.gettempdir()) / 'pdfm_normalized')
            output_dir.mkdir(parents=True, exist_ok=True)
            output_path = output_dir / f"normalized_{uuid.uuid4().hex}.pdf"
            output_doc = fitz.open()
            results = []

            for item in pages or []:
                path = item.get('path')
                if self.mode == 'server' and path in self.file_mapping:
                    path = self.file_mapping[path]
                page_index = int(item.get('pageIndex', item.get('page_index', 0)))
                src_doc = fitz.open(path)
                try:
                    src_page = src_doc[page_index]
                    rect = src_page.rect
                    scale = target_short_side / min(rect.width, rect.height)
                    new_page = output_doc.new_page(width=rect.width * scale, height=rect.height * scale)
                    new_page.show_pdf_page(new_page.rect, src_doc, page_index)
                    results.append({
                        'clientId': item.get('clientId'),
                        'pageIndex': output_doc.page_count - 1,
                        'width': round(new_page.rect.width, 3),
                        'height': round(new_page.rect.height, 3),
                    })
                finally:
                    src_doc.close()

            output_doc.save(str(output_path))
            output_doc.close()
            if self.mode == 'server':
                file_id = uuid.uuid4().hex
                self.file_mapping[file_id] = str(output_path)
                result_path = file_id
            else:
                result_path = str(output_path)
            return {'success': True, 'path': result_path, 'pages': results}
        except Exception as e:
            return {'success': False, 'error': str(e)}

    def save_csv_dialog(self, filename='报销单.csv'):
        """打开保存CSV对话框"""
        if self.mode == 'server':
            # 服务器模式：返回特殊标记，让前端触发下载
            return 'BROWSER_DOWNLOAD'
        
        # 本地模式：使用系统文件对话框
        file_types = ('CSV Files (*.csv)', 'All Files (*.*)')
        result = window.create_file_dialog(webview.FileDialog.SAVE, file_types=file_types, save_filename=filename)
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
            # 服务器模式：生成CSV内容并返回Base64
            if self.mode == 'server':
                from io import StringIO
                output = StringIO()
                writer = csv.writer(output)
                writer.writerow(['序号', '发票号', '来源', '人数', '日期', '起点', '终点', '票额'])
                total = 0.0
                for r in rows:
                    writer.writerow([r['id'], r.get('invoiceNo', '未识别'), r.get('source', ''), 
                                   r['people'], r['date'], r['start'], r['end'], r['amount']])
                    total += float(r['amount'])
                writer.writerow(['', '', '', '', '', '', '合计', f'{total:.2f}'])
                
                csv_content = output.getvalue()
                csv_base64 = base64.b64encode(csv_content.encode('utf-8-sig')).decode()
                return {
                    'success': True,
                    'content': csv_base64,
                    'filename': '报销单.csv'
                }
            
            # 本地模式：直接保存文件
            with open(path, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f)
                writer.writerow(['序号', '发票号', '来源', '人数', '日期', '起点', '终点', '票额'])
                total = 0.0
                for r in rows:
                    writer.writerow([r['id'], r.get('invoiceNo', '未识别'), r.get('source', ''), 
                                   r['people'], r['date'], r['start'], r['end'], r['amount']])
                    total += float(r['amount'])
                writer.writerow(['', '', '', '', '', '', '合计', f'{total:.2f}'])
            return {'success': True}
        except Exception as e:
            return {'success': False, 'error': str(e)}

    def save_train_ticket_csv(self, path, rows):
        """保存按日期分组的高铁票表单。"""
        try:
            from io import StringIO

            output = StringIO() if self.mode == 'server' else None
            stream = output if output is not None else open(path, 'w', newline='', encoding='utf-8-sig')
            try:
                writer = csv.writer(stream)
                writer.writerow(['日期', '发车时间', '行程', '车次', '金额', '退票费', '来源'])
                current_date = None
                group_fare_total = 0.0
                group_refund_total = 0.0
                fare_total = 0.0
                refund_total = 0.0
                for row in rows:
                    if current_date is not None and row['date'] != current_date:
                        writer.writerow([current_date, '', '小计', '', f'{group_fare_total:.2f}', f'{group_refund_total:.2f}', ''])
                        group_fare_total = 0.0
                        group_refund_total = 0.0
                    current_date = row['date']
                    amount = float(row['amount'])
                    refund_fee = float(row.get('refundFee', 0))
                    group_fare_total += amount
                    group_refund_total += refund_fee
                    fare_total += amount
                    refund_total += refund_fee
                    writer.writerow([
                        row['date'], row['time'], row['route'], row.get('trainNo', ''),
                        f'{amount:.2f}', f'{refund_fee:.2f}', row.get('source', '')
                    ])
                if current_date is not None:
                    writer.writerow([current_date, '', '小计', '', f'{group_fare_total:.2f}', f'{group_refund_total:.2f}', ''])
                writer.writerow(['', '', '票价合计（不含退票费）', '', f'{fare_total:.2f}', '', ''])
                writer.writerow(['', '', '总计（含退票费）', '', f'{fare_total + refund_total:.2f}', f'{refund_total:.2f}', ''])
            finally:
                if output is None:
                    stream.close()

            if output is not None:
                content = base64.b64encode(output.getvalue().encode('utf-8-sig')).decode()
                return {'success': True, 'content': content, 'filename': '高铁票表单.csv'}
            return {'success': True}
        except Exception as e:
            return {'success': False, 'error': str(e)}
    
    def save_statistics_csv(self, path, amounts):
        """保存统计数据到CSV文件"""
        try:
            # 服务器模式：生成CSV内容并返回Base64
            if self.mode == 'server':
                from io import StringIO
                output = StringIO()
                writer = csv.writer(output)
                writer.writerow(['序号', '发票号', '来源页面', '金额'])
                total = 0.0
                for idx, item in enumerate(amounts):
                    pages_str = ', '.join(item.get('pages', []))
                    writer.writerow([idx + 1, item.get('invoiceNo', '未识别'), pages_str, item['amount']])
                    total += float(item['amount'])
                writer.writerow(['', '', '合计（已去重）', f'{total:.2f}'])
                
                csv_content = output.getvalue()
                csv_base64 = base64.b64encode(csv_content.encode('utf-8-sig')).decode()
                return {
                    'success': True,
                    'content': csv_base64,
                    'filename': '发票统计.csv'
                }
            
            # 本地模式：直接保存文件
            with open(path, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f)
                writer.writerow(['序号', '发票号', '来源页面', '金额'])
                total = 0.0
                for idx, item in enumerate(amounts):
                    pages_str = ', '.join(item.get('pages', []))
                    writer.writerow([idx + 1, item.get('invoiceNo', '未识别'), pages_str, item['amount']])
                    total += float(item['amount'])
                writer.writerow(['', '', '合计（已去重）', f'{total:.2f}'])
            return {'success': True}
        except Exception as e:
            return {'success': False, 'error': str(e)}
        except Exception as e:
            return {'success': False, 'error': str(e)}
    
    def print_pdf(self, file_path):
        """打印PDF文件 - 读取文件内容并转换为Base64返回"""
        try:
            # 服务器模式：转换虚拟ID为真实路径
            if self.mode == 'server' and file_path in self.file_mapping:
                file_path = self.file_mapping[file_path]
            
            file_path = str(file_path)
            if not os.path.exists(file_path):
                return {'success': False, 'error': '文件不存在'}
            
            # 读取文件并转换为Base64
            with open(file_path, "rb") as pdf_file:
                encoded_string = base64.b64encode(pdf_file.read()).decode('utf-8')
            
            return {
                'success': True, 
                'data': f'data:application/pdf;base64,{encoded_string}',
                'message': '准备打印'
            }
        
        except Exception as e:
            return {'success': False, 'error': str(e)}

    def calculate_invoice_amounts(self, pages_info):
        """从电子发票PDF中提取金额信息（按发票号去重，检测重复）"""
        try:
            # 服务器模式：转换虚拟ID为真实路径
            if self.mode == 'server':
                for page in pages_info:
                    if page['path'] in self.file_mapping:
                        page['path'] = self.file_mapping[page['path']]
            
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

                    # 铁路电子客票复用“高铁票表单”的成熟识别逻辑。
                    # 普通客票统计票价；退票凭证统计实际发生的退票费。
                    if train_ticket.is_train_ticket_text(text):
                        doc.close()
                        ticket = train_ticket.extract_ticket(path, page_index, page_label)
                        invoice_no = ticket.get('invoiceNo')
                        if invoice_no == '未识别':
                            invoice_no = None
                        amount = (
                            ticket.get('refundFee', '0.00')
                            if float(ticket.get('refundFee', 0) or 0) > 0
                            else ticket.get('amount', '0.00')
                        )

                        if invoice_no:
                            if invoice_no in invoice_dict:
                                invoice_dict[invoice_no]['pages'].append(page_label)
                                invoice_dict[invoice_no]['isDuplicate'] = True
                            else:
                                invoice_dict[invoice_no] = {
                                    'amount': amount,
                                    'invoiceNo': invoice_no,
                                    'pages': [page_label],
                                    'isDuplicate': False,
                                    'fileName': file_name,
                                    'pageIndex': page_index
                                }
                        else:
                            unrecognized_invoices.append({
                                'amount': amount,
                                'invoiceNo': '未识别',
                                'pages': [page_label],
                                'isDuplicate': False,
                                'fileName': file_name,
                                'pageIndex': page_index
                            })
                        continue
                    
                    # 提取发票号码（优先使用“发票号码”标签定位）
                    invoice_no = None
                    has_invoice_label = "发票号码" in text
                    
                    try:
                        # 获取页面上所有文本块及其坐标
                        blocks = page.get_text("dict")["blocks"]
                        
                        # 查找"发票号码"关键字的位置
                        keyword_rect = None
                        keyword_text = None
                        for block in blocks:
                            if "lines" in block:
                                for line in block["lines"]:
                                    for span in line["spans"]:
                                        text_content = span["text"]
                                        if "发票号码" in text_content:
                                            keyword_rect = span["bbox"]  # (x0, y0, x1, y1)
                                            keyword_text = text_content
                                            break
                                    if keyword_rect:
                                        break
                            if keyword_rect:
                                break
                        
                        # 如果号码和标签在同一个文本块中，直接提取。
                        if keyword_rect:
                            inline_text = keyword_text.split("发票号码", 1)[1]
                            inline_match = re.search(r'(?<!\d)(\d{18,20})(?!\d)', inline_text)
                            if inline_match:
                                invoice_no = inline_match.group(1)

                        # 否则只查找标签同一基线且位于其右侧的18-20位数字。
                        # 备注区的银行账号即使长度相同，也不会进入候选。
                        if keyword_rect and not invoice_no:
                            kw_y0, kw_y1 = keyword_rect[1], keyword_rect[3]
                            kw_x1 = keyword_rect[2]  # 关键字右边界
                            kw_center_y = (kw_y0 + kw_y1) / 2
                            kw_height = kw_y1 - kw_y0
                            
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
                                            
                                            # 用文本基线中心判断同一行，避免仅有边缘重叠的备注文本。
                                            span_center_y = (span_y0 + span_y1) / 2
                                            span_height = span_y1 - span_y0
                                            same_line = abs(span_center_y - kw_center_y) <= max(
                                                2.0, min(kw_height, span_height) * 0.25
                                            )
                                            # 在关键字右侧
                                            is_right = span_x0 >= kw_x1 - 2
                                            number_match = re.fullmatch(r'\d{18,20}', span_text)
                                            
                                            if same_line and is_right and number_match:
                                                # 记录候选项：(距离, 文本)
                                                distance = span_x0 - kw_x1
                                                candidates.append((distance, span_text))
                            
                            # 选择距离最近的数字
                            if candidates:
                                candidates.sort(key=lambda x: x[0])
                                invoice_no = candidates[0][1]
                    
                    except Exception as e:
                        print(f"坐标定位失败: {e}")
                    
                    # 坐标定位失败时，仍然只匹配“发票号码”标签紧邻的号码。
                    if not invoice_no:
                        labeled_match = re.search(
                            r'发票号码[：:\s]*?(\d{18,20})(?!\d)', text
                        )
                        if labeled_match:
                            invoice_no = labeled_match.group(1)

                    # 兼容没有“发票号码”标签的旧版票据；有标签时绝不从备注区猜号码。
                    if not invoice_no and not has_invoice_label:
                        generic_match = re.search(r'\b(\d{18,20})\b', text)
                        if generic_match:
                            invoice_no = generic_match.group(1)
                    
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


# ==================== 自动路由生成器 ====================
def create_auto_server(api_instance, host='0.0.0.0', port=8000):
    """自动为 PDFMergerAPI 生成 FastAPI 路由"""
    from fastapi import FastAPI, Request, UploadFile, File
    from fastapi.responses import HTMLResponse
    from fastapi.middleware.cors import CORSMiddleware
    from typing import List
    import inspect
    import uvicorn
    
    app = FastAPI(title="PDF Merger API Server")
    
    # 配置 CORS
    app.add_middleware(
        CORSMiddleware,
        allow_origins=["*"],
        allow_credentials=True,
        allow_methods=["*"],
        allow_headers=["*"],
    )
    
    # 返回前端页面
    @app.get("/", response_class=HTMLResponse)
    async def root():
        html = ui.html_content
        # 在 </head> 前插入垫片
        shim_script = '<script src="/shim.js"></script>'
        html = html.replace("</head>", f"{shim_script}</head>")
        return html
    
    # 返回垫片脚本
    @app.get("/shim.js")
    async def get_shim():
        with open("shim.js", "r", encoding="utf-8") as f:
            content = f.read()
        from fastapi.responses import Response
        return Response(content=content, media_type="application/javascript")
    
    # 特殊处理：文件上传
    @app.post("/api/upload_files")
    async def upload_files(files: List[UploadFile] = File(...)):
        files_data = []
        for file in files:
            content = await file.read()
            files_data.append((file.filename, content))
        result = api_instance.upload_files(files_data)
        return result
    
    # 特殊处理：文件下载
    @app.get("/api/download/{file_id}")
    async def download_file(file_id: str):
        from fastapi.responses import FileResponse
        result = api_instance.download_merged_pdf(file_id)
        if result['success']:
            return FileResponse(
                path=result['file_path'],
                filename='merged.pdf',
                media_type='application/pdf'
            )
        else:
            return {'success': False, 'error': result.get('error', '未知错误')}
    
    # 自动生成所有方法的路由
    for name, method in inspect.getmembers(api_instance, predicate=inspect.ismethod):
        if name.startswith('_') or name in ['upload_files']:  # 跳过私有方法和已处理的方法
            continue
        
        # 使用默认参数捕获当前方法（避免闭包问题）
        def create_endpoint(method_func=method):
            async def endpoint(request: Request):
                try:
                    data = await request.json()
                    # 调用实际方法
                    result = method_func(**data)
                    return result
                except Exception as e:
                    return {'success': False, 'error': str(e)}
            return endpoint
        
        # 注册路由
        app.post(f"/api/{name}")(create_endpoint())
        print(f"✅ 注册路由: /api/{name}")
    
    print("=" * 60)
    print(f"🌐 访问地址: http://localhost:{port}")
    print("=" * 60)
    
    uvicorn.run(app, host=host, port=port, log_level="info")


if __name__ == '__main__':
    import argparse
    
    # 解析命令行参数
    parser = argparse.ArgumentParser(description='PDF Merger Tool')
    parser.add_argument('--server', action='store_true', help='启动 FastAPI 服务器模式')
    parser.add_argument('--port', type=int, default=8000, help='服务器端口 (默认: 8000)')
    parser.add_argument('--host', type=str, default='0.0.0.0', help='服务器地址 (默认: 0.0.0.0)')
    args = parser.parse_args()
    
    if args.server:
        # ==================== 服务器模式 ====================
        print("=" * 60)
        print("🚀 启动 FastAPI 服务器模式...")
        print("=" * 60)
        
        try:
            # 创建服务器模式的 API 实例
            api = PDFMergerAPI(mode='server')
            # 使用自动路由生成器
            create_auto_server(api, host=args.host, port=args.port)
        except ImportError as e:
            print("❌ 错误：缺少依赖")
            print(f"详情: {e}")
            print("请安装：pip install fastapi uvicorn python-multipart")
            sys.exit(1)
    else:
        # ==================== 本地模式 (pywebview) ====================
        if webview is None:
            print("❌ 错误：桌面模式需要 pywebview，请先安装项目依赖")
            sys.exit(1)
        api = PDFMergerAPI(mode='local')
        
        # 将 HTML 写入临时文件
        base_dir = os.path.dirname(os.path.abspath(__file__))
        temp_html_path = os.path.join(base_dir, 'app_gui_temp.html')
        
        with open(temp_html_path, 'w', encoding='utf-8') as f:
            f.write(ui.html_content)
        
        window = webview.create_window(
            '发票打印工具 - 专业版 v3.4',
            url=temp_html_path,
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
        
        webview.start(debug=False)
