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


if __name__ == '__main__':
    api = PDFMergerAPI()
    window = webview.create_window(
        'PDF合并工具 - 专业版',
        html=ui.html_content,
        width=1200,
        height=800,
        resizable=True,
        js_api=api
    )
    webview.start()