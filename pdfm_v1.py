import os
import json
import base64
from io import BytesIO
from PIL import Image
import PyPDF2
from PyPDF2 import PdfMerger, PdfWriter, PdfReader
import webview
import traceback
import sys

# 增加最大递归深度
sys.setrecursionlimit(2000)

class PDFMergerAPI:
    def __init__(self):
        self.pdf_files = [] # 存储完整路径
        self.current_preview = None

    def select_pdfs(self):
        """选择PDF文件"""
        file_types = ('PDF Files (*.pdf)', 'All Files (*.*)')
        files = window.create_file_dialog(
            webview.FileDialog.OPEN,
            allow_multiple=True,
            file_types=file_types
        )
        return files if files else []

    def clear_files(self):
        """清空文件列表"""
        self.pdf_files = []
        return True

    def get_file_count(self):
        return len(self.pdf_files)

    def merge_pdfs(self, output_path, mode='normal'):
        """合并PDF文件"""
        try:
            if not self.pdf_files:
                return {'success': False, 'error': '文件列表为空'}

            if not output_path:
                return {'success': False, 'error': '未选择输出文件路径'}

            if isinstance(output_path, list):
                output_path = output_path[0] if output_path else None

            if mode == 'normal':
                return self._merge_normal(output_path)
            elif mode == 'invoice':
                return self._merge_invoice(output_path)
            else:
                return {'success': False, 'error': 'Unknown merge mode'}
        except Exception as e:
            error_msg = f'合并失败：{str(e)}\n{traceback.format_exc()}'
            print(error_msg)
            return {'success': False, 'error': error_msg}

    def _merge_normal(self, output_path):
        """普通合并模式"""
        try:
            merger = PdfMerger()
            for pdf_file in self.pdf_files:
                if not os.path.exists(pdf_file):
                    return {'success': False, 'error': f'文件不存在: {pdf_file}'}
                merger.append(pdf_file)

            output_dir = os.path.dirname(output_path)
            if output_dir and not os.path.exists(output_dir):
                os.makedirs(output_dir)

            with open(output_path, 'wb') as output_file:
                merger.write(output_file)
            merger.close()

            return {'success': True, 'message': f'合并完成！输出文件：{output_path}'}
        except Exception as e:
            return {'success': False, 'error': str(e)}

    def _merge_invoice(self, output_path):
        """发票合并模式"""
        try:
            output_dir = os.path.dirname(output_path)
            if output_dir and not os.path.exists(output_dir):
                os.makedirs(output_dir)

            import fitz  # PyMuPDF
            output_doc = fitz.open()
            a4_width, a4_height = fitz.paper_size("a4")

            for i in range(0, len(self.pdf_files), 2):
                new_page = output_doc.new_page(width=a4_width, height=a4_height)
                if i < len(self.pdf_files):
                    self._place_invoice_on_page(output_doc, new_page, self.pdf_files[i], a4_width, a4_height, True)
                if i + 1 < len(self.pdf_files):
                    self._place_invoice_on_page(output_doc, new_page, self.pdf_files[i+1], a4_width, a4_height, False)

            output_doc.save(output_path)
            output_doc.close()
            return {'success': True, 'message': f'发票合并完成！输出文件：{output_path}'}
        except Exception as e:
            return {'success': False, 'error': f'发票合并失败：{str(e)}'}

    def _place_invoice_on_page(self, output_doc, target_page, source_path, a4_width, a4_height, is_top):
        import fitz
        if not os.path.exists(source_path): return
        try:
            source_doc = fitz.open(source_path)
            source_page = source_doc[0]
            pdf_width = source_page.rect.width
            pdf_height = source_page.rect.height

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
                source_doc, 0
            )
            source_doc.close()
        except Exception as e:
            print(f"处理发票页面出错: {e}")

    def preview_pdf(self, index, page_num=0, zoom_factor=2.0):
        """预览PDF文件的指定页面"""
        if 0 <= index < len(self.pdf_files):
            pdf_path = self.pdf_files[index]
            try:
                if not os.path.exists(pdf_path):
                    return {'success': False, 'error': f'文件不存在: {pdf_path}'}

                import fitz
                doc = fitz.open(pdf_path)
                total_pages = doc.page_count

                if page_num >= total_pages: page_num = 0
                page = doc[page_num]
                page_rect = page.rect

                # 智能缩放限制
                actual_zoom = float(zoom_factor)
                if page_rect.width > 1500 or page_rect.height > 2000:
                    actual_zoom = min(1.5, actual_zoom)

                matrix = fitz.Matrix(actual_zoom, actual_zoom)
                pix = page.get_pixmap(matrix=matrix, alpha=False)
                img_data = pix.tobytes("png")
                
                doc.close()
                img_base64 = base64.b64encode(img_data).decode('utf-8')

                return {
                    'success': True,
                    'image': f'data:image/png;base64,{img_base64}',
                    'total_pages': total_pages,
                    'current_page': page_num
                }
            except Exception as e:
                return {'success': False, 'error': str(e)}
        return {'success': False, 'error': '索引无效'}

    def save_file_dialog(self):
        file_types = ('PDF Files (*.pdf)', 'All Files (*.*)')
        result = window.create_file_dialog(webview.FileDialog.SAVE, file_types=file_types, save_filename='merged.pdf')
        if result and isinstance(result, list): return result[0]
        return result if result else None

    def update_file_order(self, new_order):
        try:
            valid_files = [f for f in new_order if os.path.exists(f)]
            self.pdf_files = valid_files
            return {'success': True, 'count': len(self.pdf_files)}
        except Exception as e:
            return {'success': False, 'error': str(e)}

# HTML界面
html_content = """
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>PDF合并工具</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; background-color: #f5f5f5; height: 100vh; display: flex; flex-direction: column; }
        
        .toolbar { background-color: #2c3e50; color: white; padding: 10px 20px; display: flex; align-items: center; gap: 15px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); flex-shrink: 0; }
        .toolbar button { background-color: #3498db; color: white; border: none; padding: 8px 16px; border-radius: 4px; cursor: pointer; font-size: 14px; transition: background-color 0.3s; }
        .toolbar button:hover { background-color: #2980b9; }
        .toolbar button:disabled { background-color: #7f8c8d; cursor: not-allowed; }
        
        .toolbar .zoom-controls { display: flex; align-items: center; gap: 10px; margin-left: auto; }
        .toolbar .zoom-level { color: white; font-size: 14px; min-width: 50px; text-align: center; }
        
        .toolbar .quality-selector { display: flex; align-items: center; gap: 10px; color: white; font-size: 14px; }
        .toolbar .quality-selector select { padding: 4px 8px; border-radius: 4px; border: 1px solid #34495e; background-color: #2c3e50; color: white; font-size: 14px; cursor: pointer; }
        
        .toolbar .separator { width: 1px; height: 30px; background-color: #34495e; }
        .toolbar .mode-selector { display: flex; gap: 10px; align-items: center; }
        .toolbar .mode-selector label { cursor: pointer; display: flex; align-items: center; gap: 5px; }

        .main-content { flex: 1; display: flex; overflow: hidden; }
        
        .preview-panel { width: 280px; background-color: white; border-right: 1px solid #ddd; display: flex; flex-direction: column; height: 100%; flex-shrink: 0; }
        .preview-panel .file-list-header { padding: 20px 20px 10px; border-bottom: 1px solid #eee; }
        .preview-panel .file-list-container { flex: 1; overflow-y: auto; padding: 10px 20px; }
        .preview-panel .file-list-controls { padding: 10px 20px; border-top: 1px solid #eee; background-color: #f8f9fa; display: flex; flex-direction: column; gap: 5px; }
        .preview-panel .file-list-controls button { background-color: #3498db; color: white; border: none; padding: 6px 12px; border-radius: 4px; cursor: pointer; font-size: 12px; width: 100%; }
        .preview-panel .file-list-controls button:hover { background-color: #2980b9; }
        .preview-panel .file-list-controls button:disabled { background-color: #bdc3c7; }
        .preview-panel .file-list-controls .move-controls { display: flex; gap: 5px; }
        
        .file-list { list-style: none; }
        .file-item { padding: 8px 12px; margin-bottom: 5px; background-color: #ecf0f1; border-radius: 4px; cursor: pointer; font-size: 13px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; border: 1px solid transparent; }
        .file-item:hover { background-color: #bdc3c7; }
        .file-item.active { background-color: #3498db; color: white; border-color: #2980b9; }
        
        /* 核心 CSS 修改开始 */
        .display-area { 
            flex: 1; 
            background-color: #ecf0f1; 
            overflow: auto; /* 保持滚动条开启 */
            position: relative; 
            padding: 20px;
        }

        .pdf-viewer { 
            display: flex; 
            flex-direction: column; 
            /* 关键修改：移除 align-items: center 
               这样容器宽度不够时，元素会靠左排列，右边出现滚动条 */
            align-items: flex-start; 
            min-height: 100%;
        }
        
        .pdf-page { 
            background-color: white; 
            /* 关键修改：margin: 10px auto; 
               当图片小于窗口时，auto 会自动居中
               当图片大于窗口时，auto 失效（变成靠左），从而保证左边不被切掉 */
            margin: 10px auto; 
            box-shadow: 0 4px 6px rgba(0,0,0,0.1); 
            display: flex; 
            align-items: center; 
            justify-content: center; 
            position: relative;
            width: 800px; 
            transition: width 0.2s ease;
            flex-shrink: 0; /* 防止被挤压 */
        }
        /* 核心 CSS 修改结束 */
        
        .pdf-page img { 
            width: 100%; 
            height: auto; 
            display: block; 
        }
        
        .pdf-page .page-number { position: absolute; bottom: 10px; right: 20px; background-color: rgba(0,0,0,0.5); color: white; padding: 5px 10px; border-radius: 4px; font-size: 12px; }
        
        .empty-state { text-align: center; color: #7f8c8d; margin-top: 100px; width: 100%; }
        .status-bar { background-color: #34495e; color: white; padding: 8px 20px; font-size: 12px; display: flex; justify-content: space-between; flex-shrink: 0; }
        
        .progress-bar { position: fixed; top: 50%; left: 50%; transform: translate(-50%, -50%); background-color: white; padding: 30px; border-radius: 8px; box-shadow: 0 4px 20px rgba(0,0,0,0.3); display: none; z-index: 1000; }
        .progress-bar.active { display: block; }
        .progress-bar .progress { width: 300px; height: 8px; background-color: #ecf0f1; border-radius: 4px; overflow: hidden; margin-top: 15px; }
        .progress-bar .progress-fill { height: 100%; background-color: #3498db; width: 0%; transition: width 0.3s; }
    </style>
</head>
<body>
    <div class="toolbar">
        <button onclick="selectFiles()">选择PDF文件</button>
        <button onclick="clearFiles()">清空列表</button>
        <div class="separator"></div>
        <div class="mode-selector">
            <label><input type="radio" name="mergeMode" value="normal" checked> 普通合并</label>
            <label><input type="radio" name="mergeMode" value="invoice"> 发票合并</label>
        </div>
        <div class="separator"></div>
        <button onclick="mergePDFs()">开始合并</button>

        <div class="separator"></div>
        <div class="zoom-controls">
            <button onclick="zoomOut()">-</button>
            <span class="zoom-level" id="zoomLevel">100%</span>
            <button onclick="zoomIn()">+</button>
            <button onclick="resetZoom()">重置</button>
        </div>
        
        <div class="separator"></div>
        <div class="quality-selector">
            <label>预览质量:</label>
            <select id="qualitySelect" onchange="changeQuality()">
                <option value="1.0">标准 (100%)</option>
                <option value="2.0">高清 (200%)</option>
                <option value="3.0" selected>超清 (300%)</option>
            </select>
        </div>
    </div>

    <div class="main-content">
        <div class="preview-panel">
            <div class="file-list-header"><h3>文件列表</h3></div>
            <div class="file-list-container"><ul class="file-list" id="fileList"></ul></div>
            <div class="file-list-controls">
                 <button onclick="addFiles()">增加文件</button>
                 <button onclick="removeSelectedFile()" id="removeBtn" disabled>删除选中</button>
                 <div class="move-controls">
                     <button onclick="moveUp()" id="upBtn" disabled>上移</button>
                     <button onclick="moveDown()" id="downBtn" disabled>下移</button>
                 </div>
            </div>
        </div>

        <div class="display-area">
            <div class="empty-state" id="emptyState">
                <h2>PDF合并工具</h2>
                <p>请选择PDF文件开始操作</p>
                <p style="font-size:12px; margin-top:10px; color:#999">按住 Ctrl + 鼠标滚轮可快速缩放</p>
            </div>
            <div class="pdf-viewer" id="pdfViewer" style="display: none;"></div>
        </div>
    </div>

    <div class="status-bar">
        <span id="statusText">就绪</span>
        <span id="fileCount">文件数：0</span>
    </div>

    <div class="progress-bar" id="progressBar">
        <h3>正在合并...</h3>
        <div class="progress"><div class="progress-fill" id="progressFill"></div></div>
    </div>

    <script>
        let pdfFiles = []; 
        let currentMode = 'normal';
        let currentZoom = 1.0;
        let selectedFileIndex = -1;
        let currentQuality = 2.0;

        function updateStatus(text) {
            document.getElementById('statusText').textContent = text;
        }

        document.querySelectorAll('input[name="mergeMode"]').forEach(radio => {
            radio.addEventListener('change', (e) => {
                currentMode = e.target.value;
                updateStatus(currentMode === 'normal' ? '普通合并模式' : '发票合并模式');
            });
        });

        function changeQuality() {
            const select = document.getElementById('qualitySelect');
            currentQuality = parseFloat(select.value);
            if (selectedFileIndex !== -1) {
                showMultiPagePreview(selectedFileIndex);
            }
        }

        function getBasename(path) { return path.split(/[\\\\/]/).pop(); }

        async function syncToBackend() {
            try {
                const paths = pdfFiles.map(f => f.path);
                await pywebview.api.update_file_order(paths);
            } catch (error) { console.error("同步出错:", error); }
        }

        async function selectFiles() {
            try {
                const paths = await pywebview.api.select_pdfs();
                if (paths && paths.length > 0) {
                    pdfFiles = paths.map(p => ({ path: p, name: getBasename(p) }));
                    await syncToBackend();
                    updateFileList();
                    updateStatus(`已选择 ${paths.length} 个文件`);
                    if (pdfFiles.length > 0) selectFile(0);
                }
            } catch (error) { 
                alert('选择文件失败：' + error); 
                console.error(error);
            }
        }

        async function addFiles() {
            try {
                const paths = await pywebview.api.select_pdfs();
                if (paths && paths.length > 0) {
                    let added = 0;
                    const currentSet = new Set(pdfFiles.map(f => f.path));
                    paths.forEach(p => {
                        if (!currentSet.has(p)) {
                            pdfFiles.push({ path: p, name: getBasename(p) });
                            added++;
                        }
                    });
                    if (added > 0) {
                        await syncToBackend();
                        updateFileList();
                        updateStatus(`已添加 ${added} 个文件`);
                    }
                }
            } catch (error) { 
                alert('添加文件失败：' + error); 
                console.error(error);
            }
        }

        async function clearFiles() {
            pdfFiles = [];
            selectedFileIndex = -1;
            await pywebview.api.clear_files();
            updateFileList();
            hidePreview();
            updateStatus('已清空文件列表');
        }

        function updateFileList() {
            const list = document.getElementById('fileList');
            list.innerHTML = '';
            pdfFiles.forEach((file, index) => {
                const li = document.createElement('li');
                li.className = 'file-item' + (index === selectedFileIndex ? ' active' : '');
                li.textContent = file.name;
                li.title = file.path;
                li.onclick = () => selectFile(index);
                list.appendChild(li);
            });
            document.getElementById('fileCount').textContent = `文件数：${pdfFiles.length}`;
            updateButtonStates();
        }

        async function selectFile(index) {
            if (index < 0 || index >= pdfFiles.length) return;
            selectedFileIndex = index;
            updateFileList();
            await showMultiPagePreview(index);
        }

        async function showMultiPagePreview(index) {
            const viewer = document.getElementById('pdfViewer');
            try {
                const firstPage = await pywebview.api.preview_pdf(index, 0, currentQuality);
                
                if (!firstPage.success) {
                    await syncToBackend();
                    viewer.innerHTML = `<div style="padding:20px; color:red">加载失败，请重试</div>`;
                    return;
                }

                document.getElementById('emptyState').style.display = 'none';
                viewer.style.display = 'flex';
                viewer.innerHTML = '';

                const totalPages = firstPage.total_pages;
                appendPageToViewer(firstPage, 0, totalPages);

                if (totalPages > 1) {
                    for (let i = 1; i < totalPages; i++) {
                        const pageResult = await pywebview.api.preview_pdf(index, i, currentQuality);
                        if (pageResult.success) {
                            appendPageToViewer(pageResult, i, totalPages);
                        }
                    }
                }
                
                updateZoom();

            } catch (error) {
                console.error(error);
                viewer.innerHTML = `<div style="padding:20px">预览出错: ${error}</div>`;
            }
        }

        function appendPageToViewer(pageData, pageIndex, totalPages) {
            const viewer = document.getElementById('pdfViewer');
            const div = document.createElement('div');
            div.className = 'pdf-page';
            const baseWidth = 800;
            div.style.width = (baseWidth * currentZoom) + 'px';
            
            div.innerHTML = `
                <img src="${pageData.image}" alt="Page ${pageIndex + 1}">
                <div class="page-number">${pageIndex + 1} / ${totalPages}</div>
            `;
            viewer.appendChild(div);
        }

        function hidePreview() {
            document.getElementById('emptyState').style.display = 'block';
            document.getElementById('pdfViewer').style.display = 'none';
        }

        async function removeSelectedFile() {
            if (selectedFileIndex >= 0 && selectedFileIndex < pdfFiles.length) {
                pdfFiles.splice(selectedFileIndex, 1);
                if (pdfFiles.length === 0) selectedFileIndex = -1;
                else if (selectedFileIndex >= pdfFiles.length) selectedFileIndex = pdfFiles.length - 1;
                
                await syncToBackend();
                updateFileList();
                if (selectedFileIndex >= 0) selectFile(selectedFileIndex);
                else hidePreview();
            }
        }

        async function moveUp() {
            if (selectedFileIndex > 0) {
                [pdfFiles[selectedFileIndex], pdfFiles[selectedFileIndex - 1]] = 
                [pdfFiles[selectedFileIndex - 1], pdfFiles[selectedFileIndex]];
                selectedFileIndex--;
                await syncToBackend();
                updateFileList();
                selectFile(selectedFileIndex);
            }
        }

        async function moveDown() {
            if (selectedFileIndex >= 0 && selectedFileIndex < pdfFiles.length - 1) {
                [pdfFiles[selectedFileIndex], pdfFiles[selectedFileIndex + 1]] = 
                [pdfFiles[selectedFileIndex + 1], pdfFiles[selectedFileIndex]];
                selectedFileIndex++;
                await syncToBackend();
                updateFileList();
                selectFile(selectedFileIndex);
            }
        }

        async function mergePDFs() {
            if (pdfFiles.length === 0) { alert('请先选择文件'); return; }
            await syncToBackend();
            
            const output = await pywebview.api.save_file_dialog();
            if (!output) return;

            document.getElementById('progressBar').classList.add('active');
            document.getElementById('progressFill').style.width = '30%';

            setTimeout(async () => {
                const res = await pywebview.api.merge_pdfs(output, currentMode);
                document.getElementById('progressFill').style.width = '100%';
                setTimeout(() => {
                    document.getElementById('progressBar').classList.remove('active');
                    document.getElementById('progressFill').style.width = '0%';
                    alert(res.success ? res.message : '失败: ' + res.error);
                }, 300);
            }, 100);
        }

        function updateButtonStates() {
            document.getElementById('removeBtn').disabled = selectedFileIndex === -1;
            document.getElementById('upBtn').disabled = selectedFileIndex <= 0;
            document.getElementById('downBtn').disabled = selectedFileIndex === -1 || selectedFileIndex >= pdfFiles.length - 1;
        }

        function updateZoom() {
            const baseWidth = 800;
            const newWidth = baseWidth * currentZoom;
            const pages = document.querySelectorAll('.pdf-page');
            pages.forEach(page => {
                page.style.width = newWidth + 'px';
            });
            document.getElementById('zoomLevel').textContent = Math.round(currentZoom * 100) + '%';
        }
        
        function zoomIn() { if(currentZoom < 3.0) { currentZoom += 0.1; updateZoom(); } }
        function zoomOut() { if(currentZoom > 0.3) { currentZoom -= 0.1; updateZoom(); } }
        function resetZoom() { currentZoom = 1.0; updateZoom(); }

        document.querySelector('.display-area').addEventListener('wheel', (e) => {
            if (e.ctrlKey) {
                e.preventDefault();
                e.deltaY > 0 ? zoomOut() : zoomIn();
            }
        });
    </script>
</body>
</html>
"""

if __name__ == '__main__':
    api = PDFMergerAPI()
    window = webview.create_window(
        'PDF合并工具',
        html=html_content,
        width=1200,
        height=800,
        resizable=True,
        js_api=api
    )
    webview.start()