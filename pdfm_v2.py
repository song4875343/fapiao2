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
        """
        获取页面图片
        quality: 1.0 (缩略图), 5.0 (超高清)
        """
        try:
            file_path = str(file_path)
            if not os.path.exists(file_path):
                return {'success': False, 'error': '文件不存在'}

            import fitz
            doc = fitz.open(file_path)
            if page_index >= doc.page_count:
                return {'success': False, 'error': '页码越界'}

            page = doc[page_index]
            zoom = float(quality)
            
            # 安全限制：防止内存溢出 (限制最大边长 8000px)
            rect = page.rect
            if (rect.width * zoom > 8000) or (rect.height * zoom > 8000):
                scale_limit = 8000 / max(rect.width, rect.height)
                zoom = min(zoom, scale_limit)
            
            matrix = fitz.Matrix(zoom, zoom)
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

            if mode == 'normal':
                writer = PdfWriter()
                open_files = {} 
                try:
                    for item in page_list:
                        path = str(item['path'])
                        page_idx = int(item['page_index'])
                        
                        if path not in open_files:
                            if not os.path.exists(path): continue
                            open_files[path] = PdfReader(path)
                        
                        reader = open_files[path]
                        if 0 <= page_idx < len(reader.pages):
                            writer.add_page(reader.pages[page_idx])

                    with open(output_path, 'wb') as f:
                        writer.write(f)
                except Exception as e:
                    raise e
                
                return {'success': True, 'message': '合并成功', 'output_path': output_path}
            
            elif mode == 'invoice':
                return self._merge_invoice_by_pages(page_list, output_path)

        except Exception as e:
            return {'success': False, 'error': str(e)}

    def _merge_invoice_by_pages(self, page_list, output_path):
        """发票合并逻辑"""
        try:
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
            return {'success': True, 'message': '发票合并成功', 'output_path': output_path}
        except Exception as e:
            return {'success': False, 'error': str(e)}

    def _place_page_on_canvas(self, output_doc, target_page, page_info, a4_width, a4_height, is_top):
        import fitz
        path = str(page_info['path'])
        page_idx = int(page_info['page_index'])
        if not os.path.exists(path): return

        try:
            src_doc = fitz.open(path)
            src_page = src_doc[page_idx]
            
            pdf_width = src_page.rect.width
            pdf_height = src_page.rect.height

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
                src_doc, page_idx
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

# HTML界面
html_content = """
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>PDF合并工具 - 完美版</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; background-color: #f5f5f5; height: 100vh; display: flex; flex-direction: column; }
        
        /* 工具栏 */
        .toolbar { background-color: #2c3e50; color: white; padding: 10px 20px; display: flex; align-items: center; gap: 15px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); flex-shrink: 0; }
        .toolbar button { background-color: #3498db; color: white; border: none; padding: 8px 16px; border-radius: 4px; cursor: pointer; font-size: 14px; transition: background-color 0.3s; }
        .toolbar button:hover { background-color: #2980b9; }
        .toolbar button.danger { background-color: #e74c3c; }
        .toolbar button.danger:hover { background-color: #c0392b; }
        .toolbar button.success { background-color: #27ae60; }
        .toolbar button.success:hover { background-color: #2ecc71; }
        
        .toolbar .separator { width: 1px; height: 30px; background-color: #34495e; }
        .toolbar .mode-selector { display: flex; gap: 10px; align-items: center; }
        .toolbar .mode-selector label { cursor: pointer; display: flex; align-items: center; gap: 5px; }

        .main-content { flex: 1; display: flex; overflow: hidden; }
        
        /* 左侧面板 */
        .left-sidebar { 
            width: 240px; background-color: white; border-right: 1px solid #ddd; 
            display: flex; flex-direction: column; height: 100%; flex-shrink: 0; 
        }

        .source-panel { flex: 1; display: flex; flex-direction: column; border-bottom: 1px solid #ddd; min-height: 200px; }
        .history-panel { flex: 1; display: flex; flex-direction: column; background-color: #fcfcfc; }

        .panel-header { padding: 10px 15px; border-bottom: 1px solid #eee; background: #f8f9fa; display: flex; justify-content: space-between; align-items: center; }
        .panel-header h3 { font-size: 13px; color: #2c3e50; font-weight: 600; }
        
        .list-container { list-style: none; flex: 1; overflow-y: auto; padding: 5px; }
        
        /* 源文件项 */
        .list-item { 
            padding: 8px 10px; margin-bottom: 2px; background-color: #ecf0f1; border-radius: 4px; 
            font-size: 12px; color: #555; display: flex; justify-content: space-between; align-items: center; 
            cursor: pointer; /* 变手指 */
            transition: background-color 0.2s;
        }
        .list-item:hover { background-color: #bdc3c7; }
        
        /* 历史记录项 */
        .history-item { 
            padding: 8px 10px; margin-bottom: 4px; background-color: #e8f6f3; border: 1px solid #d1f2eb; 
            border-radius: 4px; font-size: 12px; color: #16a085; cursor: pointer; transition: all 0.2s;
        }
        .history-item:hover { background-color: #d1f2eb; transform: translateX(2px); }
        .history-item .time { font-size: 10px; opacity: 0.7; display: block; margin-top: 2px; }

        /* 右侧区域 */
        .right-area { 
            flex: 1; background-color: #e0e5ec; position: relative; display: flex; flex-direction: column;
        }

        /* 模式1：工作台 */
        .workspace-view {
            flex: 1; overflow-y: auto; padding: 20px; display: flex; flex-direction: column;
        }

        /* 模式2：垂直检查 */
        .review-view {
            flex: 1; overflow-y: auto; background-color: #555; padding: 20px;
            display: none; flex-direction: column; align-items: center;
        }

        .page-grid {
            display: flex; flex-wrap: wrap; gap: 15px; align-content: flex-start; min-height: 200px; padding-bottom: 50px;
        }

        .page-card {
            width: 140px; height: 200px; background: white; border-radius: 6px; box-shadow: 0 2px 5px rgba(0,0,0,0.1);
            display: flex; flex-direction: column; position: relative; cursor: grab; border: 2px solid transparent; user-select: none;
            transition: transform 0.1s, box-shadow 0.1s;
        }
        .page-card:active { cursor: grabbing; }
        .page-card:hover { transform: translateY(-2px); box-shadow: 0 5px 15px rgba(0,0,0,0.2); border-color: #3498db; }
        .page-card.dragging { opacity: 0.5; border: 2px dashed #3498db; }

        .card-preview {
            flex: 1; padding: 10px; display: flex; align-items: center; justify-content: center;
            overflow: hidden; background: #fafafa; border-radius: 6px 6px 0 0; pointer-events: none;
        }
        .card-preview img { max-width: 100%; max-height: 100%; box-shadow: 0 1px 3px rgba(0,0,0,0.2); }
        .card-info {
            height: 30px; background: white; border-top: 1px solid #eee; display: flex;
            align-items: center; justify-content: center; font-size: 11px; color: #666; border-radius: 0 0 6px 6px;
        }
        .btn-delete {
            position: absolute; top: -8px; right: -8px; width: 24px; height: 24px;
            background: #e74c3c; color: white; border-radius: 50%; display: flex; align-items: center; justify-content: center;
            cursor: pointer; box-shadow: 0 2px 4px rgba(0,0,0,0.2); opacity: 0; transition: opacity 0.2s; z-index: 10;
        }
        .page-card:hover .btn-delete { opacity: 1; }

        .review-page {
            background-color: white; margin: 10px 0; box-shadow: 0 4px 10px rgba(0,0,0,0.3);
            width: 800px; min-height: 400px; transition: width 0.2s;
        }
        .review-page img { width: 100%; display: block; }
        
        .review-toolbar {
            position: absolute; top: 10px; left: 50%; transform: translateX(-50%);
            background: rgba(0,0,0,0.8); color: white; padding: 10px 20px; border-radius: 30px;
            display: flex; gap: 15px; align-items: center; z-index: 100;
        }
        .review-toolbar button { background: none; border: 1px solid #777; padding: 4px 10px; border-radius: 4px; font-size: 12px; cursor: pointer; color: white;}
        .review-toolbar button:hover { border-color: white; }

        /* 全屏高清预览层 */
        .preview-modal {
            position: fixed; top: 0; left: 0; right: 0; bottom: 0;
            background: rgba(0,0,0,0.95); z-index: 3000; display: none; flex-direction: column;
        }
        .preview-toolbar-top {
            height: 50px; background: #2c3e50; display: flex; align-items: center; justify-content: center;
            gap: 20px; padding: 0 20px; color: white; flex-shrink: 0;
        }
        .preview-content {
            flex: 1; overflow: auto; display: flex; align-items: flex-start; padding: 40px;
            background-color: #333;
        }
        .preview-image {
            display: block; margin: 0 auto; box-shadow: 0 0 20px rgba(0,0,0,0.5); background: white;
            transition: width 0.2s ease;
        }
        .close-btn { position: absolute; right: 20px; top: 10px; background: none; border: none; color: white; font-size: 24px; cursor: pointer; }

        .status-bar { background-color: #34495e; color: white; padding: 8px 20px; font-size: 12px; display: flex; justify-content: space-between; flex-shrink: 0; }
        .progress-overlay {
            position: fixed; top: 0; left: 0; right: 0; bottom: 0; background: rgba(0,0,0,0.5); display: none;
            justify-content: center; align-items: center; z-index: 2000;
        }
        .progress-box { background: white; padding: 20px; border-radius: 8px; width: 300px; text-align: center; }
        .progress-bar { height: 6px; background: #eee; border-radius: 3px; margin-top: 10px; overflow: hidden; }
        .progress-fill { height: 100%; background: #3498db; width: 0%; transition: width 0.3s; }
        .empty-state { display: flex; flex-direction: column; align-items: center; justify-content: center; height: 100%; color: #95a5a6; }
    </style>
</head>
<body>
    <div class="toolbar">
        <button onclick="addFiles()">+ 添加PDF文件</button>
        <div class="separator"></div>
        <div class="mode-selector">
            <label><input type="radio" name="mergeMode" value="normal" checked> 普通合并</label>
            <label><input type="radio" name="mergeMode" value="invoice"> 发票合并</label>
        </div>
        <div class="separator"></div>
        <button onclick="clearAll()" class="danger">清空全部</button>
        <div style="flex:1"></div>
        <button onclick="switchToWorkspace()" id="btnBackEdit" style="display:none; margin-right: 10px;">&lt; 返回编辑</button>
        <button onclick="startMerge()" class="success">合并并保存</button>
    </div>

    <div class="main-content">
        <div class="left-sidebar">
            <div class="source-panel">
                <div class="panel-header"><h3>源文件列表 (点击预览)</h3></div>
                <ul class="list-container" id="sourceList"></ul>
            </div>
            <div class="history-panel">
                <div class="panel-header"><h3>合并生成记录</h3></div>
                <ul class="list-container" id="historyList"></ul>
            </div>
        </div>

        <div class="right-area">
            <!-- 工作台 -->
            <div class="workspace-view" id="workspaceView">
                <div id="emptyState" class="empty-state">
                    <h2>工作台为空</h2>
                    <p>点击左上角“添加PDF文件”开始</p>
                    <p>拖拽排序，<b>双击查看超清大图</b></p>
                </div>
                <div id="pageGrid" class="page-grid"></div>
            </div>

            <!-- 垂直检查 -->
            <div class="review-view" id="reviewView">
                <div class="review-toolbar">
                    <span>检查文件 (Ctrl+滚轮缩放)</span>
                    <button onclick="reviewZoomOut()">-</button>
                    <span id="reviewZoomLevel">100%</span>
                    <button onclick="reviewZoomIn()">+</button>
                    <button onclick="switchToWorkspace()">关闭检查</button>
                </div>
                <div id="reviewContent" style="width: 100%; display: flex; flex-direction: column; align-items: center; padding-top: 50px;"></div>
            </div>
        </div>
    </div>

    <!-- 双击预览 Modal -->
    <div class="preview-modal" id="previewModal">
        <div class="preview-toolbar-top">
            <span>超清预览 (500%)</span>
            <div class="separator" style="background:#555"></div>
            <button onclick="zoomOut()">- 缩小</button>
            <span class="zoom-display" id="zoomLevel">100%</span>
            <button onclick="zoomIn()">+ 放大</button>
            <button class="close-btn" onclick="closePreview()">×</button>
        </div>
        <div class="preview-content" id="previewContainer">
            <img src="" class="preview-image" id="previewImage">
        </div>
    </div>

    <div class="status-bar">
        <span id="statusText">就绪</span>
        <span id="totalStats">总页数：0</span>
    </div>

    <div class="progress-overlay" id="progressOverlay">
        <div class="progress-box">
            <h3 id="progressText">处理中...</h3>
            <div class="progress-bar"><div class="progress-fill" id="progressFill"></div></div>
        </div>
    </div>

    <script>
        let allPages = [];
        let sourceFiles = [];
        let historyFiles = [];
        let isProcessing = false;
        
        let currentPreviewZoom = 1.0;
        let currentReviewZoom = 1.0;
        const BASE_WIDTH = 800; 

        function generateUUID() {
            return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, c => {
                var r = Math.random() * 16 | 0, v = c == 'x' ? r : (r & 0x3 | 0x8);
                return v.toString(16);
            });
        }

        async function addFiles() {
            if (isProcessing) return;
            switchToWorkspace();
            isProcessing = true;
            showProgress('正在分析文件...', 30);
            try {
                const newFiles = await pywebview.api.select_pdfs();
                if (newFiles && newFiles.length > 0) {
                    newFiles.forEach(f => {
                        if (!sourceFiles.find(sf => sf.path === f.path)) sourceFiles.push(f);
                    });
                    renderSourceList();

                    let newPages = [];
                    newFiles.forEach(file => {
                        for (let i = 0; i < file.page_count; i++) {
                            newPages.push({
                                id: generateUUID(),
                                path: file.path,
                                pageIndex: i,
                                fileName: file.name
                            });
                        }
                    });

                    allPages = allPages.concat(newPages);
                    document.getElementById('emptyState').style.display = 'none';
                    renderPageGrid();
                    loadThumbnails(newPages);
                }
            } catch (error) { alert('添加失败: ' + error); } 
            finally {
                isProcessing = false;
                hideProgress();
                updateStats();
            }
        }

        function renderSourceList() {
            const list = document.getElementById('sourceList');
            list.innerHTML = '';
            sourceFiles.forEach(f => {
                const li = document.createElement('li');
                li.className = 'list-item';
                // 支持点击预览源文件
                li.onclick = () => loadReview(f.path);
                li.innerHTML = `
                    <div style="overflow:hidden;text-overflow:ellipsis;white-space:nowrap;width:150px;font-weight:500">${f.name}</div>
                    <span>${f.page_count}p</span>
                `;
                list.appendChild(li);
            });
        }

        async function startMerge() {
            if (allPages.length === 0) { alert('请先添加文件'); return; }
            const output = await pywebview.api.save_file_dialog();
            if (!output) return; 

            showProgress('正在合并...', 50);
            const mode = document.querySelector('input[name="mergeMode"]:checked').value;
            const mergeData = allPages.map(p => ({ path: p.path, page_index: p.pageIndex }));

            setTimeout(async () => {
                const res = await pywebview.api.merge_pages(mergeData, output, mode);
                document.getElementById('progressFill').style.width = '100%';
                
                setTimeout(() => {
                    hideProgress();
                    if (res.success) {
                        addHistory(res.output_path);
                        alert('合并成功！点击左侧历史记录可检查。');
                    } else {
                        alert('合并错误: ' + res.error);
                    }
                }, 500);
            }, 100);
        }

        function addHistory(path) {
            const name = path.replace(/\\\\/g, '/').split('/').pop();
            const now = new Date();
            const timeStr = `${now.getHours().toString().padStart(2,'0')}:${now.getMinutes().toString().padStart(2,'0')}`;
            
            historyFiles.unshift({ path: path, name: name, time: timeStr });
            renderHistoryList();
            loadReview(path);
        }

        function renderHistoryList() {
            const list = document.getElementById('historyList');
            list.innerHTML = '';
            historyFiles.forEach((f) => {
                const li = document.createElement('li');
                li.className = 'history-item';
                li.onclick = () => loadReview(f.path);
                li.innerHTML = `
                    <div style="font-weight:bold; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;">${f.name}</div>
                    <span class="time">${f.time} - 点击预览</span>
                `;
                list.appendChild(li);
            });
        }

        // ================= 检查模式 (500% 高清 + Ctrl滚轮) =================
        async function loadReview(path) {
            document.getElementById('workspaceView').style.display = 'none';
            document.getElementById('reviewView').style.display = 'flex';
            document.getElementById('btnBackEdit').style.display = 'block';
            
            const container = document.getElementById('reviewContent');
            container.innerHTML = '<div style="color:white; margin-top:50px;">正在加载 500% 超清预览...</div>';

            const info = await pywebview.api.get_file_info(path);
            if (!info.success) {
                container.innerHTML = `<div style="color:red; margin-top:50px;">无法读取: ${info.error}</div>`;
                return;
            }

            container.innerHTML = '';
            currentReviewZoom = 1.0;
            updateReviewZoomUI();

            for (let i = 0; i < info.page_count; i++) {
                const pageDiv = document.createElement('div');
                pageDiv.className = 'review-page';
                pageDiv.style.width = BASE_WIDTH + 'px'; 
                
                const img = document.createElement('img');
                img.alt = `Page ${i+1}`;
                
                // 请求 5.0 (500%) 高清图
                pywebview.api.get_page_image(path, i, 5.0).then(res => {
                    if (res.success) img.src = res.image;
                });
                
                pageDiv.appendChild(img);
                
                const label = document.createElement('div');
                label.textContent = `Page ${i+1}`;
                label.style.textAlign = 'center'; label.style.padding = '5px'; label.style.color = '#555';
                pageDiv.appendChild(label);

                container.appendChild(pageDiv);
            }
        }

        function switchToWorkspace() {
            document.getElementById('reviewView').style.display = 'none';
            document.getElementById('workspaceView').style.display = 'flex';
            document.getElementById('btnBackEdit').style.display = 'none';
        }

        function updateReviewZoomUI() {
            document.getElementById('reviewZoomLevel').textContent = Math.round(currentReviewZoom * 100) + '%';
            const width = BASE_WIDTH * currentReviewZoom;
            document.querySelectorAll('.review-page').forEach(el => { el.style.width = width + 'px'; });
        }
        function reviewZoomIn() { if(currentReviewZoom < 3.0) { currentReviewZoom += 0.2; updateReviewZoomUI(); } }
        function reviewZoomOut() { if(currentReviewZoom > 0.4) { currentReviewZoom -= 0.2; updateReviewZoomUI(); } }

        // 支持 Ctrl + 滚轮 在检查模式缩放
        document.getElementById('reviewView').addEventListener('wheel', (e) => {
            if (e.ctrlKey) {
                e.preventDefault();
                e.deltaY > 0 ? reviewZoomOut() : reviewZoomIn();
            }
        });

        // ================= 编辑模式 =================
        function renderPageGrid() {
            const grid = document.getElementById('pageGrid');
            grid.innerHTML = ''; 
            allPages.forEach((page) => {
                const card = document.createElement('div');
                card.className = 'page-card';
                card.draggable = true;
                card.dataset.id = page.id;
                card.innerHTML = `
                    <div class="btn-delete" title="删除" onclick="deletePage('${page.id}')">×</div>
                    <div class="card-preview"><img id="img-${page.id}" src="" style="opacity:0.3"></div>
                    <div class="card-info" title="${page.fileName}">${page.fileName} - P${page.pageIndex + 1}</div>
                `;
                card.ondblclick = () => openPreview(page);
                addDragEvents(card);
                grid.appendChild(card);
            });
        }

        async function loadThumbnails(pagesToLoad) {
            for (const page of pagesToLoad) {
                const imgEl = document.getElementById(`img-${page.id}`);
                if (imgEl && !imgEl.src.startsWith('data')) {
                    pywebview.api.get_page_image(page.path, page.pageIndex, 0.5).then(res => {
                        if (res.success) { imgEl.src = res.image; imgEl.style.opacity = '1'; }
                    });
                }
            }
        }

        let dragSrcEl = null;
        function addDragEvents(item) {
            item.addEventListener('dragstart', function(e) {
                dragSrcEl = this;
                e.dataTransfer.effectAllowed = 'move';
                e.dataTransfer.setData('text/html', this.innerHTML);
                this.classList.add('dragging');
            });
            item.addEventListener('dragover', function(e) {
                e.preventDefault(); e.dataTransfer.dropEffect = 'move';
                const targetCard = e.target.closest('.page-card');
                if (targetCard && targetCard !== dragSrcEl) {
                    const grid = document.getElementById('pageGrid');
                    const children = Array.from(grid.children);
                    if (children.indexOf(dragSrcEl) < children.indexOf(targetCard)) targetCard.after(dragSrcEl);
                    else targetCard.before(dragSrcEl);
                }
                return false;
            });
            item.addEventListener('dragend', function(e) {
                this.classList.remove('dragging');
                const newOrderIds = Array.from(document.querySelectorAll('.page-card')).map(el => el.dataset.id);
                const newAllPages = [];
                newOrderIds.forEach(id => {
                    const p = allPages.find(x => x.id === id);
                    if (p) newAllPages.push(p);
                });
                allPages = newAllPages;
            });
        }

        async function openPreview(page) {
            const modal = document.getElementById('previewModal');
            const img = document.getElementById('previewImage');
            modal.style.display = 'flex';
            img.src = ''; img.alt = '正在渲染超清大图 (500%)...';
            
            currentPreviewZoom = 1.0;
            updatePreviewZoom();

            try {
                const res = await pywebview.api.get_page_image(page.path, page.pageIndex, 5.0);
                if (res.success) img.src = res.image;
                else img.alt = '加载失败: ' + res.error;
            } catch (e) { console.error(e); }
        }

        function closePreview() { document.getElementById('previewModal').style.display = 'none'; }
        function updatePreviewZoom() {
            const img = document.getElementById('previewImage');
            document.getElementById('zoomLevel').textContent = Math.round(currentPreviewZoom * 100) + '%';
            img.style.width = (BASE_WIDTH * currentPreviewZoom) + 'px';
        }
        function zoomIn() { if (currentPreviewZoom < 4.0) { currentPreviewZoom += 0.25; updatePreviewZoom(); } }
        function zoomOut() { if (currentPreviewZoom > 0.25) { currentPreviewZoom -= 0.25; updatePreviewZoom(); } }

        document.getElementById('previewModal').addEventListener('wheel', (e) => {
            if (e.ctrlKey) { e.preventDefault(); e.deltaY > 0 ? zoomOut() : zoomIn(); }
        });
        document.addEventListener('keydown', (e) => { if (e.key === 'Escape') closePreview(); });

        window.deletePage = function(id) {
            event.stopPropagation();
            const index = allPages.findIndex(p => p.id === id);
            if (index !== -1) {
                allPages.splice(index, 1);
                document.querySelector(`.page-card[data-id="${id}"]`).remove();
                updateStats();
                if (allPages.length === 0) document.getElementById('emptyState').style.display = 'flex';
            }
        }
        function clearAll() {
            allPages = []; 
            document.getElementById('pageGrid').innerHTML = '';
            // 同时清空源文件列表
            sourceFiles = [];
            renderSourceList();
            
            document.getElementById('emptyState').style.display = 'flex'; 
            updateStats();
        }
        function updateStats() { document.getElementById('totalStats').textContent = `总页数：${allPages.length}`; }
        function showProgress(text, pct) {
            const el = document.getElementById('progressOverlay');
            document.getElementById('progressText').textContent = text;
            document.getElementById('progressFill').style.width = pct + '%';
            el.style.display = 'flex';
        }
        function hideProgress() { document.getElementById('progressOverlay').style.display = 'none'; document.getElementById('progressFill').style.width = '0%'; }
    </script>
</body>
</html>
"""

if __name__ == '__main__':
    api = PDFMergerAPI()
    window = webview.create_window(
        'PDF合并工具 - 完美版',
        html=html_content,
        width=1200,
        height=800,
        resizable=True,
        js_api=api
    )
    webview.start()