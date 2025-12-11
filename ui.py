html_content = """
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>PDF合并工具 - 专业版</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; background-color: #f5f5f5; height: 100vh; display: flex; flex-direction: column; user-select: none; }
        
        .toolbar { background-color: #2c3e50; color: white; padding: 10px 20px; display: flex; align-items: center; gap: 15px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); flex-shrink: 0; z-index: 100; }
        .toolbar button { background-color: #3498db; color: white; border: none; padding: 8px 16px; border-radius: 4px; cursor: pointer; font-size: 14px; transition: background-color 0.3s; }
        .toolbar button:hover { background-color: #2980b9; }
        .toolbar button.danger { background-color: #e74c3c; }
        .toolbar button.danger:hover { background-color: #c0392b; }
        .toolbar button.success { background-color: #27ae60; }
        .toolbar button.success:hover { background-color: #2ecc71; }
        .toolbar button.secondary { background-color: #7f8c8d; }
        .toolbar button.secondary:hover { background-color: #95a5a6; }
        .toolbar button.warning { background-color: #f39c12; }
        .toolbar button.warning:hover { background-color: #d68910; }
        .toolbar .separator { width: 1px; height: 30px; background-color: #34495e; }
        .toolbar .mode-selector { display: flex; gap: 10px; align-items: center; }
        .toolbar .mode-selector label { cursor: pointer; display: flex; align-items: center; gap: 5px; }

        .config-btn { padding: 4px 8px !important; font-size: 12px !important; margin-left: 5px; display: none; }
        
        .main-content { flex: 1; display: flex; overflow: hidden; }
        .left-sidebar { width: 240px; background-color: white; border-right: 1px solid #ddd; display: flex; flex-direction: column; height: 100%; flex-shrink: 0; }
        .source-panel { flex: 1; display: flex; flex-direction: column; border-bottom: 1px solid #ddd; min-height: 200px; }
        .history-panel { flex: 1; display: flex; flex-direction: column; background-color: #fcfcfc; }
        .panel-header { padding: 10px 15px; border-bottom: 1px solid #eee; background: #f8f9fa; display: flex; justify-content: space-between; align-items: center; }
        .panel-header h3 { font-size: 13px; color: #2c3e50; font-weight: 600; }
        .list-container { list-style: none; flex: 1; overflow-y: auto; padding: 5px; }
        .list-item { padding: 8px 10px; margin-bottom: 2px; background-color: #ecf0f1; border-radius: 4px; font-size: 12px; color: #555; display: flex; justify-content: space-between; align-items: center; cursor: pointer; transition: background-color 0.2s; }
        .list-item:hover { background-color: #bdc3c7; }
        .history-item { padding: 8px 10px; margin-bottom: 4px; background-color: #e8f6f3; border: 1px solid #d1f2eb; border-radius: 4px; font-size: 12px; color: #16a085; cursor: pointer; transition: all 0.2s; }
        .history-item:hover { background-color: #d1f2eb; transform: translateX(2px); }

        .right-area { flex: 1; background-color: #e0e5ec; position: relative; display: flex; flex-direction: column; }
        .workspace-view { flex: 1; overflow-y: auto; padding: 20px; display: flex; flex-direction: column; position: relative; }
        .review-view { flex: 1; overflow-y: auto; background-color: #555; padding: 20px; display: none; flex-direction: column; align-items: center; }
        .page-grid { display: flex; flex-wrap: wrap; gap: 15px; align-content: flex-start; min-height: 200px; padding-bottom: 80px; }
        
        .page-card { width: 140px; height: 190px; background: white; border-radius: 6px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); display: flex; flex-direction: column; position: relative; cursor: grab; border: 2px solid transparent; transition: transform 0.1s, box-shadow 0.1s; }
        .page-card.selected { border-color: #3498db; background-color: #ebf5fb; box-shadow: 0 0 0 2px rgba(52, 152, 219, 0.3); }
        .page-card:active { cursor: grabbing; }
        .page-card:hover { transform: translateY(-2px); box-shadow: 0 5px 15px rgba(0,0,0,0.2); }
        .page-card.dragging { opacity: 0.5; border: 2px dashed #3498db; }
        .card-preview { flex: 1; padding: 10px; display: flex; align-items: center; justify-content: center; overflow: hidden; background: #fafafa; border-radius: 6px 6px 0 0; pointer-events: none; }
        .card-preview img { max-width: 100%; max-height: 100%; box-shadow: 0 1px 3px rgba(0,0,0,0.2); transition: transform 0.3s ease; }
        .card-info { height: 28px; background: white; border-top: 1px solid #eee; display: flex; align-items: center; justify-content: center; font-size: 11px; color: #666; border-radius: 0 0 6px 6px; }
        .check-mark { position: absolute; top: 5px; right: 5px; width: 20px; height: 20px; background: #3498db; color: white; border-radius: 50%; display: none; align-items: center; justify-content: center; font-size: 12px; box-shadow: 0 2px 4px rgba(0,0,0,0.2); z-index: 10; }
        .page-card.selected .check-mark { display: flex; }

        .batch-toolbar { position: fixed; bottom: 40px; left: 50%; transform: translateX(-50%) translateY(100px); background: #2c3e50; color: white; padding: 10px 20px; border-radius: 30px; display: flex; gap: 20px; align-items: center; z-index: 500; box-shadow: 0 5px 20px rgba(0,0,0,0.3); transition: transform 0.3s cubic-bezier(0.175, 0.885, 0.32, 1.275); }
        .batch-toolbar.visible { transform: translateX(-50%) translateY(0); }
        .batch-btn { background: none; border: 1px solid rgba(255,255,255,0.3); color: white; padding: 5px 15px; border-radius: 20px; cursor: pointer; font-size: 13px; }
        .batch-btn:hover { background: rgba(255,255,255,0.1); border-color: white; }
        .batch-btn.delete { color: #ffadad; border-color: rgba(255,173,173,0.3); }
        .batch-btn.delete:hover { background: rgba(255,0,0,0.2); }
        .batch-btn.fill-form { color: #f9e79f; border-color: rgba(249, 231, 159, 0.3); }
        .batch-btn.fill-form:hover { background: rgba(249, 231, 159, 0.1); }

        .selection-box { position: absolute; border: 1px solid #3498db; background-color: rgba(52, 152, 219, 0.2); display: none; z-index: 1000; pointer-events: none; }
        .review-page { background-color: white; margin: 10px 0; box-shadow: 0 4px 10px rgba(0,0,0,0.3); width: 800px; min-height: 400px; transition: width 0.2s; position: relative; display: flex; align-items: center; justify-content: center; }
        .review-page img { width: 100%; display: block; min-height: 200px; }
        .review-loading { position: absolute; color: #999; font-size: 14px; z-index: 0; }
        .review-toolbar { position: absolute; top: 10px; left: 50%; transform: translateX(-50%); background: rgba(0,0,0,0.8); color: white; padding: 10px 20px; border-radius: 30px; display: flex; gap: 15px; align-items: center; z-index: 100; }
        .review-toolbar button { background: none; border: 1px solid #777; padding: 4px 10px; border-radius: 4px; font-size: 12px; cursor: pointer; color: white;}
        .review-toolbar button:hover { border-color: white; }

        .preview-modal, .common-modal { position: fixed; top: 0; left: 0; right: 0; bottom: 0; background: rgba(0,0,0,0.95); z-index: 3000; display: none; flex-direction: column; }
        .common-modal { background: rgba(0,0,0,0.5); justify-content: center; align-items: center; }
        .modal-content { background: white; padding: 25px; border-radius: 8px; width: 600px; max-height: 80vh; display: flex; flex-direction: column; box-shadow: 0 5px 25px rgba(0,0,0,0.2); }
        .modal-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px; border-bottom: 1px solid #eee; padding-bottom: 10px; }
        .modal-body { flex: 1; overflow-y: auto; margin-bottom: 20px; }
        .modal-footer { display: flex; justify-content: flex-end; gap: 10px; }
        
        .preview-toolbar-top { height: 50px; background: #2c3e50; display: flex; align-items: center; padding: 0 20px; color: white; flex-shrink: 0; justify-content: space-between; }
        .preview-toolbar-group { display: flex; align-items: center; gap: 15px; }
        .preview-content { flex: 1; overflow: auto; display: flex; align-items: flex-start; padding: 40px; background-color: #333; }
        .preview-image { display: block; margin: 0 auto; box-shadow: 0 0 20px rgba(0,0,0,0.5); background: white; transition: width 0.2s ease, transform 0.3s ease; }
        
        .status-bar { background-color: #34495e; color: white; padding: 8px 20px; font-size: 12px; display: flex; justify-content: space-between; flex-shrink: 0; z-index: 600; position:relative;}
        .progress-overlay { position: fixed; top: 0; left: 0; right: 0; bottom: 0; background: rgba(0,0,0,0.5); display: none; justify-content: center; align-items: center; z-index: 2000; }
        .progress-box { background: white; padding: 20px; border-radius: 8px; width: 300px; text-align: center; }
        .progress-bar { height: 6px; background: #eee; border-radius: 3px; margin-top: 10px; overflow: hidden; }
        .progress-fill { height: 100%; background: #3498db; width: 0%; transition: width 0.3s; }
        .empty-state { display: flex; flex-direction: column; align-items: center; justify-content: center; height: 100%; color: #95a5a6; }

        .route-table, .result-table { width: 100%; border-collapse: collapse; font-size: 13px; }
        .route-table th, .route-table td, .result-table th, .result-table td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        .route-table th, .result-table th { background-color: #f2f2f2; }
        .route-table input { width: 100%; border: none; background: transparent; }
        .del-row-btn { color: #e74c3c; cursor: pointer; text-align: center; font-weight: bold; }
        
        /* 合计行样式 */
        .result-table tfoot tr { font-weight: bold; background-color: #f8f9fa; }
    </style>
</head>
<body>
    <div class="toolbar">
        <button onclick="addFiles()">+ 添加PDF文件</button>
        <div class="separator"></div>
        <div class="mode-selector">
            <label><input type="radio" name="mergeMode" value="normal" checked onchange="toggleConfigBtn()"> 普通合并</label>
            <label><input type="radio" name="mergeMode" value="invoice" onchange="toggleConfigBtn()"> 发票合并</label>
            <button id="btnConfig" class="config-btn warning" onclick="openRouteConfig()">配置路线</button>
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
            <div class="workspace-view" id="workspaceView">
                <div id="emptyState" class="empty-state">
                    <h2>工作台为空</h2>
                    <p>点击左上角“添加PDF文件”开始</p>
                    <p>拖拽框选，Ctrl+点击多选</p>
                </div>
                <div id="pageGrid" class="page-grid"></div>
                <div id="selectionBox" class="selection-box"></div>
            </div>

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

    <div class="batch-toolbar" id="batchToolbar">
        <span id="batchCount">已选 0 项</span>
        <div style="width:1px; height:15px; background:rgba(255,255,255,0.3)"></div>
        <button class="batch-btn" onclick="batchRotate(-90)">↺ 左旋</button>
        <button class="batch-btn" onclick="batchRotate(90)">↻ 右旋</button>
        <div style="width:1px; height:15px; background:rgba(255,255,255,0.3)"></div>
        <button class="batch-btn fill-form" id="btnFillForm" onclick="openDateModal()" style="display:none">📝 填写报销单</button>
        <button class="batch-btn delete" onclick="batchDelete()">删除所选</button>
    </div>

    <!-- 路线配置 Modal -->
    <div class="common-modal" id="routeConfigModal">
        <div class="modal-content">
            <div class="modal-header"><h3>发票路线配置</h3><button onclick="closeModal('routeConfigModal')">×</button></div>
            <div class="modal-body">
                <table class="route-table" id="routeTable">
                    <thead><tr><th>起点</th><th>终点</th><th width="80">票额(约)</th><th width="50">操作</th></tr></thead>
                    <tbody></tbody>
                </table>
                <button onclick="addRouteRow()" style="margin-top:10px; color:#3498db; background:none; border:none; cursor:pointer;">+ 添加一行</button>
            </div>
            <div class="modal-footer">
                <button onclick="closeModal('routeConfigModal')" class="secondary">取消</button>
                <button onclick="saveRoutes()" class="success">保存配置</button>
            </div>
        </div>
    </div>

    <!-- 日期选择 Modal -->
    <div class="common-modal" id="dateModal">
        <div class="modal-content" style="width: 400px; height: auto;">
            <div class="modal-header"><h3>填写报销单</h3></div>
            <div class="modal-body">
                <p style="margin-bottom:10px; color:#666; font-size:13px;">请输入报销时间范围（例如：2025年7-12月），程序将自动分配工作日。</p>
                <input type="text" id="dateRangeInput" placeholder="2025年7-12月" style="width:100%; padding:8px; border:1px solid #ddd; border-radius:4px;">
            </div>
            <div class="modal-footer">
                <button onclick="closeModal('dateModal')" class="secondary">取消</button>
                <button onclick="submitFillForm()" class="success">生成表格</button>
            </div>
        </div>
    </div>

    <!-- 结果展示 Modal -->
    <div class="common-modal" id="resultModal">
        <div class="modal-content" style="width: 800px;">
            <div class="modal-header"><h3>生成的报销单</h3><button onclick="closeModal('resultModal')">×</button></div>
            <div class="modal-body">
                <table class="result-table" id="resultTable">
                    <thead><tr><th>序号</th><th>人数</th><th>日期</th><th>起点</th><th>终点</th><th>票额</th></tr></thead>
                    <tbody></tbody>
                    <tfoot></tfoot>
                </table>
            </div>
            <div class="modal-footer">
                <button onclick="copyTableToClipboard()" class="warning">复制 (Excel)</button>
                <button onclick="saveTableToCSV()" class="success">保存 CSV</button>
                <button onclick="closeModal('resultModal')" class="secondary">关闭</button>
            </div>
        </div>
    </div>

    <!-- 预览 Modal -->
    <div class="preview-modal" id="previewModal">
        <div class="preview-toolbar-top">
            <div class="preview-toolbar-group">
                <button onclick="closePreview()" class="secondary" style="font-size: 14px; padding: 6px 12px;">&lt; 返回编辑</button>
                <span style="font-size: 14px; opacity: 0.8">超清预览 (5.0x)</span>
            </div>
            <div class="preview-toolbar-group">
                <button onclick="zoomOut()">-</button>
                <span class="zoom-display" id="zoomLevel">100%</span>
                <button onclick="zoomIn()">+</button>
            </div>
            <div class="preview-toolbar-group">
                <button onclick="rotateCurrentPage(-90)" title="左旋">↺</button>
                <button onclick="rotateCurrentPage(90)" title="右旋">↻</button>
            </div>
        </div>
        <div class="preview-content" id="previewContainer"><img src="" class="preview-image" id="previewImage"></div>
    </div>

    <div class="status-bar"><span id="statusText">就绪</span><span id="totalStats">总页数：0</span></div>
    <div class="progress-overlay" id="progressOverlay">
        <div class="progress-box"><h3 id="progressText">处理中...</h3><div class="progress-bar"><div class="progress-fill" id="progressFill"></div></div></div>
    </div>

    <script>
        let allPages = [];
        let sourceFiles = [];
        let historyFiles = [];
        let historyCache = {}; 
        let isProcessing = false;
        
        let selectedPageIds = new Set();
        let isSelecting = false;
        let startX, startY;
        const workspace = document.getElementById('workspaceView');
        const selectionBox = document.getElementById('selectionBox');
        
        let currentPreviewZoom = 1.0;
        let currentReviewZoom = 1.0;
        let currentPreviewPageId = null;
        const BASE_WIDTH = 800; 
        let reviewObserver = null;
        
        // 缓存当前的报销单数据用于保存CSV
        let currentTableData = [];

        function generateUUID() { return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, c => { var r = Math.random() * 16 | 0, v = c == 'x' ? r : (r & 0x3 | 0x8); return v.toString(16); }); }

        function toggleConfigBtn() {
            const mode = document.querySelector('input[name="mergeMode"]:checked').value;
            const btn = document.getElementById('btnConfig');
            const fillBtn = document.getElementById('btnFillForm');
            if (mode === 'invoice') {
                btn.style.display = 'inline-block';
                if (selectedPageIds.size > 0) fillBtn.style.display = 'inline-block';
            } else {
                btn.style.display = 'none';
                fillBtn.style.display = 'none';
            }
        }

        async function openRouteConfig() {
            const routes = await pywebview.api.get_routes();
            const tbody = document.querySelector('#routeTable tbody'); tbody.innerHTML = '';
            routes.forEach(r => addRouteRow(r));
            document.getElementById('routeConfigModal').style.display = 'flex';
        }

        function addRouteRow(data = {start:'', end:'', price:''}) {
            const tbody = document.querySelector('#routeTable tbody');
            const tr = document.createElement('tr');
            tr.innerHTML = `<td><input type="text" value="${data.start||''}" placeholder="起点"></td><td><input type="text" value="${data.end||''}" placeholder="终点"></td><td><input type="number" value="${data.price||''}" placeholder="0.0"></td><td class="del-row-btn" onclick="this.parentElement.remove()">×</td>`;
            tbody.appendChild(tr);
        }

        async function saveRoutes() {
            const rows = Array.from(document.querySelectorAll('#routeTable tbody tr'));
            const data = rows.map(tr => { const i = tr.querySelectorAll('input'); return { start: i[0].value, end: i[1].value, price: parseFloat(i[2].value) || 0 }; }).filter(i => i.start && i.end);
            await pywebview.api.save_routes(data); closeModal('routeConfigModal'); alert('配置已保存');
        }

        function closeModal(id) { document.getElementById(id).style.display = 'none'; }

        function openDateModal() { document.getElementById('dateRangeInput').value = '2025年7-12月'; document.getElementById('dateModal').style.display = 'flex'; }

        async function submitFillForm() {
            const dateStr = document.getElementById('dateRangeInput').value;
            if (!dateStr) return;
            const selectedPaths = []; selectedPageIds.forEach(id => { const p = allPages.find(x => x.id === id); if (p) selectedPaths.push(p.path); });
            if (selectedPaths.length === 0) { alert("请先选择发票"); return; }

            showProgress('正在分析发票...', 50);
            setTimeout(async () => {
                const res = await pywebview.api.generate_reimbursement_form(selectedPaths, dateStr);
                hideProgress(); closeModal('dateModal');
                
                if (res.success) {
                    currentTableData = res.rows;
                    const tbody = document.querySelector('#resultTable tbody');
                    const tfoot = document.querySelector('#resultTable tfoot');
                    tbody.innerHTML = ''; tfoot.innerHTML = '';
                    
                    let totalAmount = 0.0;
                    res.rows.forEach(r => {
                        const tr = document.createElement('tr');
                        tr.innerHTML = `<td>${r.id}</td><td>${r.people}</td><td>${r.date}</td><td>${r.start}</td><td>${r.end}</td><td>${r.amount}</td>`;
                        tbody.appendChild(tr);
                        totalAmount += parseFloat(r.amount);
                    });
                    
                    tfoot.innerHTML = `<tr><td colspan="4"></td><td>合计</td><td>${totalAmount.toFixed(2)}</td></tr>`;
                    document.getElementById('resultModal').style.display = 'flex';
                } else {
                    alert(res.error);
                }
            }, 100);
        }

        // --- 修复版复制功能 (解决undefined报错和python转义问题) ---
        function copyTableToClipboard() {
            if (!currentTableData || currentTableData.length === 0) return;
            
            // 注意：这里使用双反斜杠 \\t 和 \\n，防止Python字符串转义导致JS语法错误
            let text = "序号\\t人数\\t日期\\t起点\\t终点\\t票额\\n";
            let total = 0.0;
            currentTableData.forEach(r => {
                text += `${r.id}\\t${r.people}\\t${r.date}\\t${r.start}\\t${r.end}\\t${r.amount}\\n`;
                total += parseFloat(r.amount);
            });
            text += `\\t\\t\\t\\t合计\\t${total.toFixed(2)}`;
            
            // 安全检查：防止 navigator.clipboard 为 undefined
            if (navigator && navigator.clipboard && navigator.clipboard.writeText) {
                navigator.clipboard.writeText(text)
                    .then(() => alert("表格已复制到剪贴板，可直接粘贴到 Excel"))
                    .catch(() => fallbackCopy(text));
            } else {
                fallbackCopy(text);
            }
        }

        function fallbackCopy(text) {
            var textArea = document.createElement("textarea");
            textArea.value = text;
            textArea.style.position = "fixed";
            textArea.style.left = "-9999px";
            document.body.appendChild(textArea);
            textArea.focus();
            textArea.select();
            
            try {
                var successful = document.execCommand('copy');
                if(successful) alert("表格已复制到剪贴板 (兼容模式)");
                else alert("复制失败，浏览器限制");
            } catch (err) {
                alert("复制失败: " + err);
            }
            document.body.removeChild(textArea);
        }

        async function saveTableToCSV() {
            if (!currentTableData || currentTableData.length === 0) return;
            const path = await pywebview.api.save_csv_dialog();
            if (path) {
                const res = await pywebview.api.save_csv_data(path, currentTableData);
                if (res.success) alert("保存成功！"); else alert("保存失败: " + res.error);
            }
        }

        // ================= 常规逻辑 =================
        function getRelativeCoordinates(e, container) { const rect = container.getBoundingClientRect(); return { x: e.clientX - rect.left + container.scrollLeft, y: e.clientY - rect.top + container.scrollTop }; }
        workspace.addEventListener('mousedown', (e) => { if (e.target.closest('.page-card')) return; if (e.offsetX > workspace.clientWidth || e.offsetY > workspace.clientHeight) return; isSelecting = true; const c = getRelativeCoordinates(e, workspace); startX = c.x; startY = c.y; if (!e.ctrlKey) clearSelection(); selectionBox.style.left = startX + 'px'; selectionBox.style.top = startY + 'px'; selectionBox.style.width = '0px'; selectionBox.style.height = '0px'; selectionBox.style.display = 'block'; });
        window.addEventListener('mousemove', (e) => { if (!isSelecting) return; const c = getRelativeCoordinates(e, workspace); const w = Math.abs(c.x - startX); const h = Math.abs(c.y - startY); const l = Math.min(c.x, startX); const t = Math.min(c.y, startY); selectionBox.style.width = w + 'px'; selectionBox.style.height = h + 'px'; selectionBox.style.left = l + 'px'; selectionBox.style.top = t + 'px'; checkSelection(l, t, w, h, e.ctrlKey); });
        window.addEventListener('mouseup', () => { if (isSelecting) { isSelecting = false; selectionBox.style.display = 'none'; updateBatchToolbar(); } });
        function checkSelection(l, t, w, h, isCtrl) { const r = l + w; const b = t + h; const ws = workspace.getBoundingClientRect(); document.querySelectorAll('.page-card').forEach(c => { const cr = c.getBoundingClientRect(); const cl = cr.left - ws.left + workspace.scrollLeft; const ct = cr.top - ws.top + workspace.scrollTop; if (!(r < cl || l > (cl + cr.width) || b < ct || t > (ct + cr.height))) { selectedPageIds.add(c.dataset.id); c.classList.add('selected'); } }); }
        function onCardClick(e, id) { e.stopPropagation(); if (e.ctrlKey) { if (selectedPageIds.has(id)) { selectedPageIds.delete(id); document.querySelector(`.page-card[data-id="${id}"]`).classList.remove('selected'); } else { selectedPageIds.add(id); document.querySelector(`.page-card[data-id="${id}"]`).classList.add('selected'); } } else { clearSelection(); selectedPageIds.add(id); document.querySelector(`.page-card[data-id="${id}"]`).classList.add('selected'); } updateBatchToolbar(); }
        function clearSelection() { selectedPageIds.clear(); document.querySelectorAll('.page-card.selected').forEach(el => el.classList.remove('selected')); updateBatchToolbar(); }
        function updateBatchToolbar() { const b = document.getElementById('batchToolbar'); if (selectedPageIds.size > 0) { b.classList.add('visible'); document.getElementById('batchCount').textContent = `已选 ${selectedPageIds.size} 项`; const mode = document.querySelector('input[name="mergeMode"]:checked').value; document.getElementById('btnFillForm').style.display = (mode === 'invoice') ? 'inline-block' : 'none'; } else { b.classList.remove('visible'); } }
        function batchRotate(angle) { selectedPageIds.forEach(id => { rotatePageById(id, angle); }); }
        function batchDelete() { if (!confirm(`确定要删除选中的 ${selectedPageIds.size} 个页面吗？`)) return; const d = Array.from(selectedPageIds); allPages = allPages.filter(p => !selectedPageIds.has(p.id)); d.forEach(id => { const el = document.querySelector(`.page-card[data-id="${id}"]`); if(el) el.remove(); }); clearSelection(); updateStats(); if (allPages.length === 0) document.getElementById('emptyState').style.display = 'flex'; }
        function rotatePageById(id, angle) { const p = allPages.find(x => x.id === id); if (p) { let r = (p.rotation || 0) + angle; r = (r % 360 + 360) % 360; p.rotation = r; const t = document.getElementById(`img-${p.id}`); if (t) t.style.transform = `rotate(${r}deg)`; } }

        async function addFiles() { if (isProcessing) return; switchToWorkspace(); isProcessing = true; showProgress('正在分析文件...', 30); try { const f = await pywebview.api.select_pdfs(); if (f && f.length > 0) { f.forEach(i => { if (!sourceFiles.find(s => s.path === i.path)) sourceFiles.push(i); }); renderSourceList(); let n = []; f.forEach(file => { for (let i = 0; i < file.page_count; i++) n.push({ id: generateUUID(), path: file.path, pageIndex: i, fileName: file.name, rotation: 0 }); }); allPages = allPages.concat(n); document.getElementById('emptyState').style.display = 'none'; renderPageGrid(); loadThumbnails(n); } } catch (e) { alert('添加失败: ' + e); } finally { isProcessing = false; hideProgress(); updateStats(); } }
        function renderSourceList() { const l = document.getElementById('sourceList'); l.innerHTML = ''; sourceFiles.forEach(f => { const li = document.createElement('li'); li.className = 'list-item'; li.onclick = () => loadReview(f.path); li.innerHTML = `<div style="overflow:hidden;text-overflow:ellipsis;white-space:nowrap;width:150px;font-weight:500">${f.name}</div><span>${f.page_count}p</span>`; l.appendChild(li); }); }
        function renderPageGrid() { const g = document.getElementById('pageGrid'); const c = {}; g.querySelectorAll('.page-card img').forEach(i => { if (i.src && i.src.startsWith('data:')) c[i.id.replace('img-', '')] = i.src; }); g.innerHTML = ''; allPages.forEach(p => { const d = document.createElement('div'); d.className = 'page-card'; d.draggable = true; d.dataset.id = p.id; const src = c[p.id] || ''; const op = src ? 'opacity:1' : 'opacity:0.3'; d.innerHTML = `<div class="check-mark">✓</div><div class="card-preview"><img id="img-${p.id}" src="${src}" style="${op}; transform: rotate(${p.rotation||0}deg)"></div><div class="card-info" title="${p.fileName}">${p.fileName} - P${p.pageIndex + 1}</div>`; d.onclick = (e) => onCardClick(e, p.id); d.ondblclick = (e) => { e.stopPropagation(); openPreview(p); }; addDragEvents(d); g.appendChild(d); }); }
        
        async function startMerge() { if (allPages.length === 0) { alert('请先添加文件'); return; } const o = await pywebview.api.save_file_dialog(); if (!o) return; showProgress('正在合并...', 50); const m = document.querySelector('input[name="mergeMode"]:checked').value; const d = allPages.map(p => ({ path: p.path, page_index: p.pageIndex, rotation: p.rotation })); setTimeout(async () => { const r = await pywebview.api.merge_pages(d, o, m); document.getElementById('progressFill').style.width = '100%'; setTimeout(() => { hideProgress(); if (r.success) { if (r.thumbnail) historyCache[r.output_path] = r.thumbnail; addHistory(r.output_path); alert(r.message); } else alert('合并错误: ' + r.error); }, 500); }, 100); }
        function addHistory(p) { const n = p.replace(/\\\\/g, '/').split('/').pop(); const d = new Date(); historyFiles.unshift({ path: p, name: n, time: `${d.getHours().toString().padStart(2,'0')}:${d.getMinutes().toString().padStart(2,'0')}` }); renderHistoryList(); loadReview(p); }
        function renderHistoryList() { const l = document.getElementById('historyList'); l.innerHTML = ''; historyFiles.forEach(f => { const li = document.createElement('li'); li.className = 'history-item'; li.onclick = () => loadReview(f.path); li.innerHTML = `<div style="font-weight:bold; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;">${f.name}</div><span class="time">${f.time} - 点击预览</span>`; l.appendChild(li); }); }

        let dragSrcEl = null;
        function addDragEvents(item) { item.addEventListener('dragstart', function(e) { clearSelection(); dragSrcEl = this; e.dataTransfer.effectAllowed = 'move'; e.dataTransfer.setData('text/html', this.innerHTML); this.classList.add('dragging'); }); item.addEventListener('dragover', function(e) { e.preventDefault(); e.dataTransfer.dropEffect = 'move'; const t = e.target.closest('.page-card'); if (t && t !== dragSrcEl) { const g = document.getElementById('pageGrid'); const c = Array.from(g.children); if (c.indexOf(dragSrcEl) < c.indexOf(t)) t.after(dragSrcEl); else t.before(dragSrcEl); } return false; }); item.addEventListener('dragend', function(e) { this.classList.remove('dragging'); const ids = Array.from(document.querySelectorAll('.page-card')).map(el => el.dataset.id); const n = []; ids.forEach(id => { const p = allPages.find(x => x.id === id); if (p) n.push(p); }); allPages = n; }); }
        
        function findCachedThumb(path, index) { const p = allPages.find(x => x.path === path && x.pageIndex === index); if (p) { const el = document.getElementById(`img-${p.id}`); if (el && el.src.startsWith('data:')) return el.src; } return null; }
        
        async function loadReview(path) {
            document.getElementById('workspaceView').style.display = 'none'; document.getElementById('reviewView').style.display = 'flex'; document.getElementById('btnBackEdit').style.display = 'block';
            const c = document.getElementById('reviewContent'); c.innerHTML = '<div style="color:white; margin-top:50px;">正在获取文件信息...</div>';
            if (reviewObserver) reviewObserver.disconnect();
            const info = await pywebview.api.get_file_info(path);
            if (!info.success) { c.innerHTML = `<div style="color:red; margin-top:50px;">错误：${info.error}</div>`; return; }
            c.innerHTML = ''; currentReviewZoom = 1.0; updateReviewZoomUI();
            
            for (let i = 0; i < info.page_count; i++) {
                const div = document.createElement('div'); div.className = 'review-page'; div.style.width = BASE_WIDTH + 'px'; div.dataset.path = path; div.dataset.index = i;
                let src = findCachedThumb(path, i); if (!src && i === 0 && historyCache[path]) src = historyCache[path];
                const img = document.createElement('img'); img.alt = `Page ${i+1}`;
                if (src) { img.src = src; img.dataset.status = 'thumb'; } else { img.src = ''; div.innerHTML += '<span class="review-loading">等待加载...</span>'; }
                div.appendChild(img); c.appendChild(div);
            }
            const prioritizeVisibles = async () => { const n = Array.from(c.querySelectorAll('.review-page')).slice(0, 2); for (const div of n) { const img = div.querySelector('img'); if (!img.src) { const r = await pywebview.api.get_page_image(div.dataset.path, div.dataset.index, 0.5); if (r.success) { img.src = r.image; img.dataset.status = 'thumb'; } } loadHighResImage(div, img); } };
            prioritizeVisibles();
            reviewObserver = new IntersectionObserver((entries, observer) => { entries.forEach(entry => { if (entry.isIntersecting) { const div = entry.target; const img = div.querySelector('img'); if (!img.src || img.dataset.status === 'thumb') loadHighResImage(div, img); observer.unobserve(div); } }); }, { root: document.getElementById('reviewView'), rootMargin: '200px', threshold: 0.01 });
            document.querySelectorAll('.review-page').forEach(div => reviewObserver.observe(div));
        }

        async function loadHighResImage(div, img) { if (div.dataset.loading === 'true') return; div.dataset.loading = 'true'; const res = await pywebview.api.get_page_image(div.dataset.path, parseInt(div.dataset.index), 5.0); if (res.success) { img.src = res.image; img.dataset.status = 'hd'; const l = div.querySelector('.review-loading'); if (l) l.remove(); } delete div.dataset.loading; }
        
        function switchToWorkspace() { if (reviewObserver) reviewObserver.disconnect(); document.getElementById('reviewView').style.display = 'none'; document.getElementById('workspaceView').style.display = 'flex'; document.getElementById('btnBackEdit').style.display = 'none'; }
        async function openPreview(page) { currentPreviewPageId = page.id; document.getElementById('previewModal').style.display = 'flex'; const img = document.getElementById('previewImage'); img.style.transform = `rotate(${page.rotation||0}deg)`; currentPreviewZoom = 1.0; updatePreviewZoom(); const t = document.getElementById(`img-${page.id}`); img.src = (t && t.src.startsWith('data:')) ? t.src : ''; const res = await pywebview.api.get_page_image(page.path, page.pageIndex, 5.0); if (res.success) img.src = res.image; }
        function rotateCurrentPage(angle) { if (currentPreviewPageId) rotatePageById(currentPreviewPageId, angle); const p = allPages.find(x => x.id === currentPreviewPageId); if (p) document.getElementById('previewImage').style.transform = `rotate(${p.rotation}deg)`; }
        function closePreview() { document.getElementById('previewModal').style.display = 'none'; currentPreviewPageId = null; }
        function updateReviewZoomUI() { document.getElementById('reviewZoomLevel').textContent = Math.round(currentReviewZoom * 100) + '%'; const w = BASE_WIDTH * currentReviewZoom; document.querySelectorAll('.review-page').forEach(el => { el.style.width = w + 'px'; }); }
        function reviewZoomOut() { if(currentReviewZoom > 0.4) { currentReviewZoom -= 0.2; updateReviewZoomUI(); } }
        function reviewZoomIn() { if(currentReviewZoom < 3.0) { currentReviewZoom += 0.2; updateReviewZoomUI(); } }
        function updatePreviewZoom() { document.getElementById('zoomLevel').textContent = Math.round(currentPreviewZoom * 100) + '%'; document.getElementById('previewImage').style.width = (BASE_WIDTH * currentPreviewZoom) + 'px'; }
        function zoomIn() { if (currentPreviewZoom < 4.0) { currentPreviewZoom += 0.25; updatePreviewZoom(); } }
        function zoomOut() { if (currentPreviewZoom > 0.25) { currentPreviewZoom -= 0.25; updatePreviewZoom(); } }
        document.getElementById('previewModal').addEventListener('wheel', (e) => { if (e.ctrlKey) { e.preventDefault(); e.deltaY > 0 ? zoomOut() : zoomIn(); } });
        document.getElementById('reviewView').addEventListener('wheel', (e) => { if (e.ctrlKey) { e.preventDefault(); e.deltaY > 0 ? reviewZoomOut() : reviewZoomIn(); } });
        document.addEventListener('keydown', (e) => { if (e.key === 'Escape') closePreview(); });
        async function loadThumbnails(pages) { for (const p of pages) { const el = document.getElementById(`img-${p.id}`); if (el && !el.src.startsWith('data')) { pywebview.api.get_page_image(p.path, p.pageIndex, 0.5).then(res => { if (res.success) { el.src = res.image; el.style.opacity = '1'; } }); } } }
        function clearAll() { allPages = []; document.getElementById('pageGrid').innerHTML = ''; sourceFiles = []; renderSourceList(); clearSelection(); document.getElementById('emptyState').style.display = 'flex'; updateStats(); }
        function updateStats() { document.getElementById('totalStats').textContent = `总页数：${allPages.length}`; }
        function showProgress(t, p) { document.getElementById('progressOverlay').style.display = 'flex'; document.getElementById('progressText').textContent = t; document.getElementById('progressFill').style.width = p + '%'; }
        function hideProgress() { document.getElementById('progressOverlay').style.display = 'none'; document.getElementById('progressFill').style.width = '0%'; }
    </script>
</body>
</html>
"""