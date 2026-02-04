html_content = """
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>发票打印工具 - 专业版</title>
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
        .list-container { list-style: none; flex: 1; overflow-y: auto; padding: 5px; overflow-x: hidden; position: relative; }
        .list-item-wrapper { position: relative; margin-bottom: 2px; }
        .list-item { padding: 8px 10px; background-color: #ecf0f1; border-radius: 4px; font-size: 12px; color: #555; display: flex; justify-content: space-between; align-items: center; cursor: pointer; transition: all 0.3s; position: relative; }
        .list-item:hover { background-color: #bdc3c7; }
        .list-item-wrapper.slide-left .list-item { transform: translateX(-60px); }
        .list-item-delete-btn { position: absolute; right: -60px; top: 0; bottom: 0; width: 60px; background-color: #27ae60; color: white; display: flex; align-items: center; justify-content: center; cursor: pointer; font-weight: normal; font-size: 12px; border-radius: 0 4px 4px 0; transition: all 0.3s; z-index: 1; }
        .list-item-wrapper.slide-left .list-item-delete-btn { right: 0; background-color: #e74c3c; }
        .list-item-delete-btn:hover { background-color: #c0392b; }
        .history-item { padding: 8px 10px; background-color: #e8f6f3; border: 1px solid #d1f2eb; border-radius: 4px; font-size: 12px; color: #16a085; cursor: pointer; transition: all 0.3s; display: flex; flex-direction: column; gap: 2px; }
        .history-item:hover { background-color: #d1f2eb; }
        .list-item-wrapper.slide-left .history-item { transform: translateX(-60px); }

        .right-area { flex: 1; background-color: #e0e5ec; position: relative; display: flex; flex-direction: column; }
        .workspace-view { flex: 1; overflow-y: auto; padding: 20px; display: flex; flex-direction: column; position: relative; }
        .review-view { flex: 1; overflow-y: auto; background-color: #555; padding: 20px; display: none; flex-direction: column; align-items: center; }
        .help-view { flex: 1; overflow-y: auto; background: #f5f5f5; padding: 0; display: none; }
        .help-view::-webkit-scrollbar { width: 0; height: 0; }
        .help-view { scrollbar-width: none; -ms-overflow-style: none; }
        .help-container { width: 100%; min-height: 100%; background: white; overflow: visible; display: flex; flex-direction: column; }
        .help-header { background: #2c3e50; color: white; padding: 30px 40px; flex-shrink: 0; }
        .help-header h1 { font-size: 28px; margin-bottom: 8px; }
        .help-header p { font-size: 15px; opacity: 0.9; }
        .help-content { padding: 30px 40px; max-width: 1200px; margin: 0 auto; width: 100%; flex: 1; }
        .help-section { margin-bottom: 40px; }
        .help-section h2 { color: #2c3e50; font-size: 24px; margin-bottom: 20px; padding-bottom: 10px; border-bottom: 3px solid #3498db; }
        .help-section h3 { color: #34495e; font-size: 18px; margin: 20px 0 10px 0; }
        .help-section p, .help-section li { color: #555; font-size: 15px; margin-bottom: 10px; }
        .help-section ul, .help-section ol { margin-left: 20px; margin-bottom: 15px; }
        .feature-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(280px, 1fr)); gap: 20px; margin-top: 20px; }
        .feature-card { background: #f8f9fa; padding: 20px; border-radius: 8px; border-left: 4px solid #3498db; }
        .feature-card h4 { color: #2c3e50; font-size: 16px; margin-bottom: 10px; }
        .feature-card p { color: #666; font-size: 14px; margin: 0; }
        .tip-box { background: #e8f6f3; border-left: 4px solid #27ae60; padding: 15px 20px; margin: 20px 0; border-radius: 4px; }
        .tip-box strong { color: #27ae60; display: block; margin-bottom: 5px; }
        .warning-box { background: #fff3cd; border-left: 4px solid #f39c12; padding: 15px 20px; margin: 20px 0; border-radius: 4px; }
        .warning-box strong { color: #f39c12; display: block; margin-bottom: 5px; }
        .shortcut-table { width: 100%; border-collapse: collapse; margin-top: 15px; }
        .shortcut-table th, .shortcut-table td { padding: 12px; text-align: left; border-bottom: 1px solid #ddd; }
        .shortcut-table th { background: #f2f2f2; color: #2c3e50; font-weight: 600; }
        .shortcut-table tr:hover { background: #f8f9fa; }
        .kbd { display: inline-block; padding: 3px 8px; background: #f4f4f4; border: 1px solid #ccc; border-radius: 3px; font-family: monospace; font-size: 13px; box-shadow: 0 1px 2px rgba(0,0,0,0.1); }
        .help-footer { background: #f8f9fa; padding: 20px; text-align: center; color: #666; font-size: 13px; border-top: 1px solid #e0e0e0; flex-shrink: 0; }
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

        .batch-toolbar { position: fixed; bottom: 40px; left: 50%; transform: translateX(-50%) translateY(100px); background: #2c3e50; color: white; padding: 8px 20px; border-radius: 30px; display: flex; gap: 12px; align-items: center; z-index: 500; box-shadow: 0 5px 20px rgba(0,0,0,0.3); transition: transform 0.3s cubic-bezier(0.175, 0.885, 0.32, 1.275); white-space: nowrap; }
        .batch-toolbar.visible { transform: translateX(-50%) translateY(0); }
        .batch-btn { background: none; border: 1px solid rgba(255,255,255,0.3); color: white; padding: 4px 12px; border-radius: 20px; cursor: pointer; font-size: 12px; white-space: nowrap; }
        .batch-btn:hover { background: rgba(255,255,255,0.1); border-color: white; }
        .batch-btn.delete { color: #ffadad; border-color: rgba(255,173,173,0.3); }
        .batch-btn.delete:hover { background: rgba(255,0,0,0.2); }
        .batch-btn.fill-form { color: #f9e79f; border-color: rgba(249, 231, 159, 0.3); }
        .batch-btn.fill-form:hover { background: rgba(249, 231, 159, 0.1); }
        .batch-btn.info { color: #85c1e9; border-color: rgba(133, 193, 233, 0.3); }
        .batch-btn.info:hover { background: rgba(133, 193, 233, 0.1); }

        .selection-box { position: absolute; border: 1px solid #3498db; background-color: rgba(52, 152, 219, 0.2); display: none; z-index: 1000; pointer-events: none; }
        .review-page { background-color: white; margin: 10px 0; box-shadow: 0 4px 10px rgba(0,0,0,0.3); width: 800px; min-height: 400px; transition: width 0.2s; position: relative; display: flex; align-items: center; justify-content: center; }
        .review-page img { width: 100%; display: block; min-height: 200px; }
        .review-loading { position: absolute; color: #999; font-size: 14px; z-index: 0; }
        .review-toolbar { position: absolute; top: 10px; left: 50%; transform: translateX(-50%); background: rgba(0,0,0,0.8); color: white; padding: 10px 20px; border-radius: 30px; display: flex; gap: 15px; align-items: center; z-index: 100; white-space: nowrap; }
        .review-toolbar button { background: none; border: 1px solid #777; padding: 6px 12px; border-radius: 4px; font-size: 13px; cursor: pointer; color: white; white-space: nowrap; }
        .review-toolbar button:hover { border-color: white; }
        .review-toolbar span { white-space: nowrap; }

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
            <label><input type="radio" name="mergeMode" value="normal" checked onchange="toggleConfigBtn()"> 普通</label>
            <label><input type="radio" name="mergeMode" value="invoice" onchange="toggleConfigBtn()"> 发票</label>
            <button id="btnConfig" class="config-btn warning" onclick="openRouteConfig()">配置路线</button>
        </div>
        <div class="separator"></div>
        <button onclick="clearAll()" class="danger">清空全部</button>
        <div style="flex:1"></div>
        <button onclick="showHelp()" style="background-color: #95a5a6; color: white; border: none; padding: 8px 16px; border-radius: 4px; cursor: pointer; font-size: 14px; margin-right: 10px;">帮助</button>
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
                    <div style="margin-top: 25px; color: #95a5a6; font-size: 13px; line-height: 1.8; max-width: 500px;">
                        <p style="margin: 5px 0;">• 将单文件发票每两张合并到一页A4纸，节省50%打印成本</p>
                        <p style="margin: 5px 0;">• 对大量发票进行金额统计，自动识别发票号并去重</p>
                        <p style="margin: 5px 0;">• 自动编写出租车发票报销单，智能分配日期和路线</p>
                    </div>
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
                    <button onclick="printCurrentFile()" style="margin-left: 15px;">🖨️ 打印</button>
                    <button onclick="switchToWorkspace()">关闭检查</button>
                </div>
                <div id="reviewContent" style="width: 100%; display: flex; flex-direction: column; align-items: center; padding-top: 50px;"></div>
            </div>

            <div class="help-view" id="helpView">
                <div class="help-container">
                    <div class="help-header">
                        <h1>📄 发票打印工具 - 专业版</h1>
                        <p>功能强大的PDF发票处理与合并工具</p>
                    </div>
                    <div class="help-content">
                        <div class="help-section">
                            <h2>🚀 快速开始</h2>
                            <ol>
                                <li>点击工具栏的 <strong>"+ 添加PDF文件"</strong> 按钮选择发票文件</li>
                                <li>在工作台中查看和管理所有页面</li>
                                <li>选择合并模式（普通/发票）</li>
                                <li>点击 <strong>"合并并保存"</strong> 生成最终文件</li>
                            </ol>
                            <div class="tip-box">
                                <strong>💡 提示</strong>
                                如果没有选中任何页面，将合并所有页面；如果选中了部分页面，则只合并选中的页面。
                            </div>
                        </div>

                        <div class="help-section">
                            <h2>✨ 核心功能</h2>
                            <div class="feature-grid">
                                <div class="feature-card">
                                    <h4>📁 文件管理</h4>
                                    <p>支持批量添加PDF文件，左侧列表显示所有源文件，点击可预览，右键删除不需要的文件。</p>
                                </div>
                                <div class="feature-card">
                                    <h4>🔄 页面操作</h4>
                                    <p>支持页面旋转、拖拽排序、批量删除等操作，灵活调整页面顺序和方向。</p>
                                </div>
                                <div class="feature-card">
                                    <h4>🖼️ 高清预览</h4>
                                    <p>双击页面卡片进入超清预览模式（5倍清晰度），支持缩放和旋转操作。</p>
                                </div>
                                <div class="feature-card">
                                    <h4>📋 多选功能</h4>
                                    <p>支持框选和Ctrl+点击多选，批量处理多个页面，提高工作效率。</p>
                                </div>
                                <div class="feature-card">
                                    <h4>🧾 发票模式</h4>
                                    <p>专为发票设计的合并模式，自动将两张发票排版到一页A4纸上，节省打印成本。</p>
                                </div>
                                <div class="feature-card">
                                    <h4>📊 报销单生成</h4>
                                    <p>自动识别发票金额和号码，智能分配日期和路线，一键生成报销单表格。</p>
                                </div>
                                <div class="feature-card">
                                    <h4>💰 金额统计</h4>
                                    <p>自动提取发票金额，按发票号去重统计，检测重复发票，支持导出CSV。</p>
                                </div>
                                <div class="feature-card">
                                    <h4>📜 合并记录</h4>
                                    <p>自动保存合并历史，点击记录可快速预览已生成的文件。</p>
                                </div>
                            </div>
                        </div>

                        <div class="help-section">
                            <h2>📖 详细操作指南</h2>
                            <h3>1. 普通合并模式</h3>
                            <p>适用于一般PDF文件的合并，保持原始页面尺寸和布局。</p>
                            <ul>
                                <li>添加PDF文件后，所有页面会显示在工作台</li>
                                <li>可以拖拽页面卡片调整顺序</li>
                                <li>选中页面后可以旋转或删除</li>
                                <li>点击"合并并保存"选择保存位置</li>
                            </ul>

                            <h3>2. 发票合并模式</h3>
                            <p>专为发票打印优化，自动将两张发票排版到一页A4纸上。</p>
                            <ul>
                                <li>切换到"发票"模式</li>
                                <li>添加发票PDF文件</li>
                                <li>系统会自动将每两页发票合并到一页A4纸</li>
                                <li>上下各一张，居中对齐，自动缩放</li>
                            </ul>
                            <div class="tip-box">
                                <strong>💡 发票模式优势</strong>
                                使用发票模式可以节省50%的打印纸张，同时保持发票清晰可读。
                            </div>

                            <h3>3. 路线配置（发票模式）</h3>
                            <p>为报销单功能配置常用路线和票额。</p>
                            <ul>
                                <li>切换到"发票"模式后，点击"配置路线"按钮</li>
                                <li>添加常用的起点、终点和大致票额</li>
                                <li>系统会根据发票金额自动匹配最接近的路线</li>
                                <li>配置会自动保存，下次使用时无需重新设置</li>
                            </ul>

                            <h3>4. 生成报销单</h3>
                            <p>自动识别发票信息，生成规范的报销单表格。</p>
                            <ul>
                                <li>选中需要报销的发票页面</li>
                                <li>点击底部工具栏的"📝 报销单"按钮</li>
                                <li>输入报销时间范围（如：2025年7-12月）</li>
                                <li>系统会自动识别金额、匹配路线、分配日期</li>
                                <li>生成的表格可以复制到Excel或保存为CSV文件</li>
                            </ul>
                            <div class="warning-box">
                                <strong>⚠️ 注意</strong>
                                报销单功能会自动检测重复的发票号，重复的发票会被标记并去重，确保不会重复报销。
                            </div>

                            <h3>5. 金额统计</h3>
                            <p>快速统计选中发票的总金额，自动去重。</p>
                            <ul>
                                <li>选中需要统计的发票页面</li>
                                <li>点击底部工具栏的"💰 统计"按钮</li>
                                <li>系统会显示每张发票的号码、金额和来源</li>
                                <li>自动检测重复发票并显示警告</li>
                                <li>结果可以复制到Excel或保存为CSV</li>
                            </ul>
                        </div>

                        <div class="help-section">
                            <h2>⌨️ 快捷键</h2>
                            <table class="shortcut-table">
                                <thead>
                                    <tr>
                                        <th>快捷键</th>
                                        <th>功能</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    <tr>
                                        <td><span class="kbd">Ctrl</span> + <span class="kbd">点击</span></td>
                                        <td>多选页面（可选中多个不连续的页面）</td>
                                    </tr>
                                    <tr>
                                        <td><span class="kbd">拖拽框选</span></td>
                                        <td>框选多个页面（按住鼠标左键拖动）</td>
                                    </tr>
                                    <tr>
                                        <td><span class="kbd">Ctrl</span> + <span class="kbd">滚轮</span></td>
                                        <td>在预览模式下缩放页面</td>
                                    </tr>
                                    <tr>
                                        <td><span class="kbd">双击</span></td>
                                        <td>打开单页超清预览</td>
                                    </tr>
                                    <tr>
                                        <td><span class="kbd">右键</span></td>
                                        <td>在源文件或历史记录上右键显示删除按钮</td>
                                    </tr>
                                    <tr>
                                        <td><span class="kbd">Esc</span></td>
                                        <td>关闭预览窗口</td>
                                    </tr>
                                </tbody>
                            </table>
                        </div>

                        <div class="help-section">
                            <h2>❓ 常见问题</h2>
                            <h3>Q: 为什么有些发票识别不出金额？</h3>
                            <p>A: 可能是发票格式不标准或扫描质量较差。建议使用清晰的电子发票PDF文件，避免使用扫描件。</p>
                            
                            <h3>Q: 报销单中的日期是如何分配的？</h3>
                            <p>A: 系统会在您指定的时间范围内，自动筛选出所有工作日（周一至周五），然后随机分配给每张发票。</p>
                            
                            <h3>Q: 如何处理重复的发票？</h3>
                            <p>A: 系统会自动识别发票号，重复的发票会被标记并在统计时去重。在报销单和金额统计功能中都会显示重复警告。</p>
                            
                            <h3>Q: 发票模式和普通模式有什么区别？</h3>
                            <p>A: 普通模式保持原始页面尺寸，适合一般PDF合并。发票模式会将两张发票自动排版到一页A4纸上，专为打印优化。</p>
                        </div>

                        <div class="help-section">
                            <h2>💡 使用建议</h2>
                            <div class="tip-box">
                                <strong>最佳实践</strong>
                                <ul style="margin: 10px 0 0 20px;">
                                    <li>使用电子发票PDF文件，避免扫描件，识别准确率更高</li>
                                    <li>在配置路线时，设置常用的几条路线即可，系统会自动匹配</li>
                                    <li>生成报销单前先使用"统计"功能检查是否有重复发票</li>
                                    <li>定期清理合并记录，避免列表过长</li>
                                    <li>大批量处理时，可以分批添加文件，提高响应速度</li>
                                </ul>
                            </div>
                        </div>
                    </div>
                    <div class="help-footer">
                        <p>发票打印工具 - 专业版 v3.2</p>
                        <p>让发票处理更简单、更高效</p>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <div class="batch-toolbar" id="batchToolbar">
        <span id="batchCount">已选 0 项</span>
        <div style="width:1px; height:15px; background:rgba(255,255,255,0.3)"></div>
        <button class="batch-btn" onclick="batchRotate(-90)">↺ 左旋</button>
        <button class="batch-btn" onclick="batchRotate(90)">↻ 右旋</button>
        <div style="width:1px; height:15px; background:rgba(255,255,255,0.3)"></div>
        <button class="batch-btn fill-form" id="btnFillForm" onclick="openDateModal()" style="display:none">📝 报销单</button>
        <button class="batch-btn info" id="btnStatistics" onclick="statisticsAmount()" style="display:none">💰 统计</button>
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

    <!-- 结果展示 Modal（报销单） -->
    <div class="common-modal" id="resultModal">
        <div class="modal-content" style="width: 900px;">
            <div class="modal-header"><h3>生成的报销单</h3><button onclick="closeModal('resultModal')">×</button></div>
            <div class="modal-body">
                <table class="result-table" id="resultTable">
                    <thead><tr><th>序号</th><th>发票号</th><th>人数</th><th>日期</th><th>起点</th><th>终点</th><th>票额</th></tr></thead>
                    <tbody></tbody>
                    <tfoot></tfoot>
                </table>
            </div>
            <div class="modal-footer">
                <button onclick="copyReimbursementTable()" class="warning">复制 (Excel)</button>
                <button onclick="saveReimbursementCSV()" class="success">保存 CSV</button>
                <button onclick="closeModal('resultModal')" class="secondary">关闭</button>
            </div>
        </div>
    </div>

    <!-- 金额统计 Modal -->
    <div class="common-modal" id="statisticsModal">
        <div class="modal-content" style="width: 900px;">
            <div class="modal-header"><h3>发票金额统计</h3><button onclick="closeModal('statisticsModal')">×</button></div>
            <div class="modal-body">
                <!-- 内容由JS动态生成 -->
            </div>
            <div class="modal-footer">
                <button onclick="copyStatisticsTable()" class="warning">复制 (Excel)</button>
                <button onclick="saveStatisticsCSV()" class="success">保存 CSV</button>
                <button onclick="closeModal('statisticsModal')" class="secondary">关闭</button>
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
        
        // 记录当前右键激活的列表项
        let activeContextItem = null;
        
        // 记录当前预览的文件路径
        let currentReviewFilePath = null;

        function generateUUID() { return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, c => { var r = Math.random() * 16 | 0, v = c == 'x' ? r : (r & 0x3 | 0x8); return v.toString(16); }); }

        function toggleConfigBtn() {
            const mode = document.querySelector('input[name="mergeMode"]:checked').value;
            const btn = document.getElementById('btnConfig');
            const fillBtn = document.getElementById('btnFillForm');
            const statsBtn = document.getElementById('btnStatistics');
            if (mode === 'invoice') {
                btn.style.display = 'inline-block';
                if (selectedPageIds.size > 0) {
                    fillBtn.style.display = 'inline-block';
                    statsBtn.style.display = 'inline-block';
                }
            } else {
                btn.style.display = 'none';
                fillBtn.style.display = 'none';
                statsBtn.style.display = 'none';
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
                        tr.innerHTML = `<td>${r.id}</td><td>${r.invoiceNo}</td><td>${r.people}</td><td>${r.date}</td><td>${r.start}</td><td>${r.end}</td><td>${r.amount}</td>`;
                        tbody.appendChild(tr);
                        totalAmount += parseFloat(r.amount);
                    });
                    
                    tfoot.innerHTML = `<tr><td colspan="5"></td><td>合计</td><td>${totalAmount.toFixed(2)}</td></tr>`;
                    
                    // 显示重复提示
                    if (res.duplicateDetails && res.duplicateDetails.length > 0) {
                        const modalBody = document.querySelector('#resultModal .modal-body');
                        const warningDiv = document.createElement('div');
                        warningDiv.style.cssText = 'margin-top: 15px; padding: 12px; background: #fff3cd; border: 1px solid #ffc107; border-radius: 4px; color: #856404; font-size: 13px;';
                        warningDiv.innerHTML = '<strong>⚠️ 发现重复发票（已自动去重）：</strong><br><br>' + 
                            res.duplicateDetails.map(d => 
                                `<div style="margin-bottom: 8px;">
                                    <strong>发票号：${d.invoiceNo}</strong><br>
                                    出现次数：${d.count} 次<br>
                                    文件：${d.files.join('、')}
                                </div>`
                            ).join('');
                        modalBody.appendChild(warningDiv);
                    }
                    
                    document.getElementById('resultModal').style.display = 'flex';
                } else {
                    alert(res.error);
                }
            }, 100);
        }

        // 报销单复制功能
        function copyReimbursementTable() {
            if (!currentTableData || currentTableData.length === 0) return;
            
            let text = "序号\\t发票号\\t人数\\t日期\\t起点\\t终点\\t票额\\n";
            let total = 0.0;
            currentTableData.forEach(r => {
                text += `${r.id}\\t${r.invoiceNo}\\t${r.people}\\t${r.date}\\t${r.start}\\t${r.end}\\t${r.amount}\\n`;
                total += parseFloat(r.amount);
            });
            text += `\\t\\t\\t\\t\\t合计\\t${total.toFixed(2)}`;
            
            if (navigator && navigator.clipboard && navigator.clipboard.writeText) {
                navigator.clipboard.writeText(text)
                    .then(() => alert("表格已复制到剪贴板，可直接粘贴到 Excel"))
                    .catch(() => fallbackCopy(text));
            } else {
                fallbackCopy(text);
            }
        }

        // 报销单保存CSV
        async function saveReimbursementCSV() {
            if (!currentTableData || currentTableData.length === 0) return;
            const path = await pywebview.api.save_csv_dialog();
            if (path) {
                const res = await pywebview.api.save_reimbursement_csv(path, currentTableData);
                if (res.success) alert("保存成功！"); else alert("保存失败: " + res.error);
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
        
        // 统计数据缓存
        let currentStatisticsData = [];
        
        // 统计发票金额
        async function statisticsAmount() {
            const selectedPages = [];
            selectedPageIds.forEach(id => {
                const p = allPages.find(x => x.id === id);
                if (p) {
                    selectedPages.push({
                        path: p.path,
                        pageIndex: p.pageIndex,
                        fileName: p.fileName,
                        id: p.id
                    });
                }
            });
            
            if (selectedPages.length === 0) {
                alert("请先选择发票");
                return;
            }
            
            showProgress('正在识别发票金额...', 50);
            setTimeout(async () => {
                const res = await pywebview.api.calculate_invoice_amounts(selectedPages);
                hideProgress();
                
                if (res.success) {
                    currentStatisticsData = res.amounts;  // 缓存数据
                    
                    const modalBody = document.querySelector('#statisticsModal .modal-body');
                    modalBody.innerHTML = '';
                    
                    const tableContainer = document.createElement('div');
                    tableContainer.innerHTML = `
                        <table class="result-table" id="statisticsTable">
                            <thead><tr><th>序号</th><th>发票号</th><th>来源页面</th><th>金额</th></tr></thead>
                            <tbody></tbody>
                            <tfoot></tfoot>
                        </table>
                    `;
                    modalBody.appendChild(tableContainer);
                    
                    const tbody = document.querySelector('#statisticsTable tbody');
                    const tfoot = document.querySelector('#statisticsTable tfoot');
                    
                    res.amounts.forEach((item, index) => {
                        const tr = document.createElement('tr');
                        let rowStyle = '';
                        let invoiceDisplay = item.invoiceNo;
                        
                        if (item.isDuplicate) {
                            rowStyle = 'background-color: #fff3cd; color: #856404;';
                            invoiceDisplay += ' ⚠️';
                        }
                        
                        tr.style.cssText = rowStyle;
                        tr.innerHTML = `<td>${index + 1}</td><td>${invoiceDisplay}</td><td>${item.pages.join(', ')}</td><td>${item.amount}</td>`;
                        
                        if (item.isDuplicate) {
                            tr.title = `重复发票 - 已去重统计`;
                        }
                        
                        tbody.appendChild(tr);
                    });
                    
                    tfoot.innerHTML = `<tr><td colspan="3">合计（已去重）</td><td>${res.totalAmount}</td></tr>`;
                    
                    if (res.duplicateDetails && res.duplicateDetails.length > 0) {
                        const warningDiv = document.createElement('div');
                        warningDiv.style.cssText = 'margin-top: 15px; padding: 12px; background: #fff3cd; border: 1px solid #ffc107; border-radius: 4px; color: #856404; font-size: 13px;';
                        warningDiv.innerHTML = '<strong>⚠️ 发现重复发票（已自动去重）：</strong><br><br>' + 
                            res.duplicateDetails.map(d => 
                                `<div style="margin-bottom: 8px;">
                                    <strong>发票号：${d.invoiceNo}</strong><br>
                                    金额：${d.amount} 元<br>
                                    出现次数：${d.count} 次<br>
                                    位置：${d.pages.join('、')}
                                </div>`
                            ).join('');
                        modalBody.appendChild(warningDiv);
                    }
                    
                    document.getElementById('statisticsModal').style.display = 'flex';
                } else {
                    alert(res.error || '识别失败');
                }
            }, 100);
        }

        // 统计表格复制功能
        function copyStatisticsTable() {
            if (!currentStatisticsData || currentStatisticsData.length === 0) return;
            
            let text = "序号\\t发票号\\t来源页面\\t金额\\n";
            let total = 0.0;
            currentStatisticsData.forEach((item, index) => {
                text += `${index + 1}\\t${item.invoiceNo}\\t${item.pages.join(', ')}\\t${item.amount}\\n`;
                total += parseFloat(item.amount);
            });
            text += `\\t\\t合计（已去重）\\t${total.toFixed(2)}`;
            
            if (navigator && navigator.clipboard && navigator.clipboard.writeText) {
                navigator.clipboard.writeText(text)
                    .then(() => alert("表格已复制到剪贴板，可直接粘贴到 Excel"))
                    .catch(() => fallbackCopy(text));
            } else {
                fallbackCopy(text);
            }
        }

        // 统计表格保存CSV
        async function saveStatisticsCSV() {
            if (!currentStatisticsData || currentStatisticsData.length === 0) return;
            const path = await pywebview.api.save_csv_dialog();
            if (path) {
                const res = await pywebview.api.save_statistics_csv(path, currentStatisticsData);
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
        function updateBatchToolbar() { 
            const b = document.getElementById('batchToolbar'); 
            if (selectedPageIds.size > 0) { 
                b.classList.add('visible'); 
                document.getElementById('batchCount').textContent = `已选 ${selectedPageIds.size} 项`; 
                const mode = document.querySelector('input[name="mergeMode"]:checked').value; 
                document.getElementById('btnFillForm').style.display = (mode === 'invoice') ? 'inline-block' : 'none'; 
                document.getElementById('btnStatistics').style.display = (mode === 'invoice') ? 'inline-block' : 'none'; 
            } else { 
                b.classList.remove('visible'); 
            } 
        }
        function batchRotate(angle) { selectedPageIds.forEach(id => { rotatePageById(id, angle); }); }
        function batchDelete() { if (!confirm(`确定要删除选中的 ${selectedPageIds.size} 个页面吗？`)) return; const d = Array.from(selectedPageIds); allPages = allPages.filter(p => !selectedPageIds.has(p.id)); d.forEach(id => { const el = document.querySelector(`.page-card[data-id="${id}"]`); if(el) el.remove(); }); clearSelection(); updateStats(); if (allPages.length === 0) document.getElementById('emptyState').style.display = 'flex'; }
        function rotatePageById(id, angle) { const p = allPages.find(x => x.id === id); if (p) { let r = (p.rotation || 0) + angle; r = (r % 360 + 360) % 360; p.rotation = r; const t = document.getElementById(`img-${p.id}`); if (t) t.style.transform = `rotate(${r}deg)`; } }

        async function addFiles() { if (isProcessing) return; switchToWorkspace(); isProcessing = true; showProgress('正在分析文件...', 30); try { const f = await pywebview.api.select_pdfs(); if (f && f.length > 0) { f.forEach(i => { if (!sourceFiles.find(s => s.path === i.path)) sourceFiles.push(i); }); renderSourceList(); let n = []; f.forEach(file => { for (let i = 0; i < file.page_count; i++) n.push({ id: generateUUID(), path: file.path, pageIndex: i, fileName: file.name, rotation: 0 }); }); allPages = allPages.concat(n); document.getElementById('emptyState').style.display = 'none'; renderPageGrid(); loadThumbnails(n); } } catch (e) { alert('添加失败: ' + e); } finally { isProcessing = false; hideProgress(); updateStats(); } }
        function renderSourceList() { 
            const l = document.getElementById('sourceList'); 
            l.innerHTML = ''; 
            sourceFiles.forEach(f => { 
                // 创建包装容器
                const wrapper = document.createElement('div');
                wrapper.className = 'list-item-wrapper';
                wrapper.dataset.path = f.path;
                
                // 创建列表项
                const li = document.createElement('li'); 
                li.className = 'list-item'; 
                li.onclick = () => loadReview(f.path); 
                li.oncontextmenu = (e) => { e.preventDefault(); showDeleteButton(wrapper); };
                
                const content = document.createElement('div');
                content.style.cssText = 'overflow:hidden;text-overflow:ellipsis;white-space:nowrap;width:150px;font-weight:500';
                content.textContent = f.name;
                
                const pageCount = document.createElement('span');
                pageCount.textContent = f.page_count + 'p';
                
                li.appendChild(content);
                li.appendChild(pageCount);
                
                // 创建删除按钮
                const deleteBtn = document.createElement('div');
                deleteBtn.className = 'list-item-delete-btn';
                deleteBtn.textContent = '删除';
                deleteBtn.onclick = (e) => deleteSourceFile(e, wrapper);
                
                wrapper.appendChild(li);
                wrapper.appendChild(deleteBtn);
                l.appendChild(wrapper); 
            }); 
        }
        
        function showDeleteButton(wrapper) {
            // 恢复之前的项
            if (activeContextItem && activeContextItem !== wrapper) {
                activeContextItem.classList.remove('slide-left');
            }
            // 激活当前项
            wrapper.classList.add('slide-left');
            activeContextItem = wrapper;
        }
        
        function deleteSourceFile(event, wrapper) {
            event.stopPropagation();
            const path = wrapper.dataset.path;
            // 从sourceFiles中移除
            sourceFiles = sourceFiles.filter(f => f.path !== path);
            // 从allPages中移除该文件的所有页面
            allPages = allPages.filter(p => p.path !== path);
            // 刷新界面
            renderSourceList();
            renderPageGrid();
            updateStats();
            if (allPages.length === 0) {
                document.getElementById('emptyState').style.display = 'flex';
            }
        }
        
        // 点击其他地方恢复列表项
        document.addEventListener('click', (e) => {
            if (activeContextItem && !e.target.closest('.list-item-wrapper')) {
                activeContextItem.classList.remove('slide-left');
                activeContextItem = null;
            }
        });
        function renderPageGrid() { const g = document.getElementById('pageGrid'); const c = {}; g.querySelectorAll('.page-card img').forEach(i => { if (i.src && i.src.startsWith('data:')) c[i.id.replace('img-', '')] = i.src; }); g.innerHTML = ''; allPages.forEach(p => { const d = document.createElement('div'); d.className = 'page-card'; d.draggable = true; d.dataset.id = p.id; const src = c[p.id] || ''; const op = src ? 'opacity:1' : 'opacity:0.3'; d.innerHTML = `<div class="check-mark">✓</div><div class="card-preview"><img id="img-${p.id}" src="${src}" style="${op}; transform: rotate(${p.rotation||0}deg)"></div><div class="card-info" title="${p.fileName}">${p.fileName} - P${p.pageIndex + 1}</div>`; d.onclick = (e) => onCardClick(e, p.id); d.ondblclick = (e) => { e.stopPropagation(); loadSinglePageReview(p); }; addDragEvents(d); g.appendChild(d); }); }
        
        async function loadSinglePageReview(page) {
            document.getElementById('workspaceView').style.display = 'none'; 
            document.getElementById('reviewView').style.display = 'flex'; 
            document.getElementById('btnBackEdit').style.display = 'block';
            const c = document.getElementById('reviewContent'); 
            c.innerHTML = '<div style="color:white; margin-top:50px;">正在加载页面...</div>';
            if (reviewObserver) reviewObserver.disconnect();
            
            c.innerHTML = ''; 
            currentReviewZoom = 1.0; 
            updateReviewZoomUI();
            
            const div = document.createElement('div'); 
            div.className = 'review-page'; 
            div.style.width = BASE_WIDTH + 'px'; 
            div.dataset.path = page.path; 
            div.dataset.index = page.pageIndex;
            
            const cachedSrc = document.getElementById(`img-${page.id}`)?.src;
            const img = document.createElement('img'); 
            img.alt = `Page ${page.pageIndex + 1}`;
            
            if (cachedSrc && cachedSrc.startsWith('data:')) { 
                img.src = cachedSrc; 
                img.dataset.status = 'thumb'; 
            } else { 
                img.src = ''; 
                div.innerHTML += '<span class="review-loading">等待加载...</span>'; 
            }
            
            div.appendChild(img); 
            c.appendChild(div);
            
            // 加载高清图
            loadHighResImage(div, img);
        }
        
        async function startMerge() { 
            // 确定要合并的页面
            let pagesToMerge = [];
            if (selectedPageIds.size > 0) {
                // 如果有选中页面，只合并选中的
                selectedPageIds.forEach(id => {
                    const page = allPages.find(p => p.id === id);
                    if (page) pagesToMerge.push(page);
                });
            } else {
                // 如果没有选中，合并所有页面
                pagesToMerge = allPages;
            }
            
            if (pagesToMerge.length === 0) { 
                alert('请先添加文件或选择要合并的页面'); 
                return; 
            } 
            
            const o = await pywebview.api.save_file_dialog(); 
            if (!o) return; 
            
            showProgress('正在合并...', 50); 
            const m = document.querySelector('input[name="mergeMode"]:checked').value; 
            const d = pagesToMerge.map(p => ({ path: p.path, page_index: p.pageIndex, rotation: p.rotation })); 
            
            setTimeout(async () => { 
                const r = await pywebview.api.merge_pages(d, o, m); 
                document.getElementById('progressFill').style.width = '100%'; 
                setTimeout(() => { 
                    hideProgress(); 
                    if (r.success) { 
                        if (r.thumbnail) historyCache[r.output_path] = r.thumbnail; 
                        addHistory(r.output_path); 
                        alert(r.message); 
                    } else alert('合并错误: ' + r.error); 
                }, 500); 
            }, 100); 
        }
        function addHistory(p) { const n = p.replace(/\\\\/g, '/').split('/').pop(); const d = new Date(); historyFiles.unshift({ path: p, name: n, time: `${d.getHours().toString().padStart(2,'0')}:${d.getMinutes().toString().padStart(2,'0')}` }); renderHistoryList(); loadReview(p); }
        
        function renderHistoryList() { 
            const l = document.getElementById('historyList'); 
            l.innerHTML = ''; 
            historyFiles.forEach(f => { 
                // 创建包装容器
                const wrapper = document.createElement('div');
                wrapper.className = 'list-item-wrapper';
                wrapper.dataset.path = f.path;
                
                // 创建列表项
                const li = document.createElement('li'); 
                li.className = 'history-item'; 
                li.onclick = () => loadReview(f.path); 
                li.oncontextmenu = (e) => { e.preventDefault(); showDeleteButton(wrapper); };
                li.innerHTML = `<div style="font-weight:bold; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;">${f.name}</div><span class="time">${f.time} - 点击预览</span>`;
                
                // 创建删除按钮
                const deleteBtn = document.createElement('div');
                deleteBtn.className = 'list-item-delete-btn';
                deleteBtn.textContent = '删除';
                deleteBtn.onclick = (e) => deleteHistoryFile(e, wrapper);
                
                wrapper.appendChild(li);
                wrapper.appendChild(deleteBtn);
                l.appendChild(wrapper); 
            }); 
        }
        
        function deleteHistoryFile(event, wrapper) {
            event.stopPropagation();
            const path = wrapper.dataset.path;
            // 从historyFiles中移除
            historyFiles = historyFiles.filter(f => f.path !== path);
            // 从缓存中移除
            if (historyCache[path]) {
                delete historyCache[path];
            }
            // 刷新界面
            renderHistoryList();
        }

        let dragSrcEl = null;
        function addDragEvents(item) { item.addEventListener('dragstart', function(e) { clearSelection(); dragSrcEl = this; e.dataTransfer.effectAllowed = 'move'; e.dataTransfer.setData('text/html', this.innerHTML); this.classList.add('dragging'); }); item.addEventListener('dragover', function(e) { e.preventDefault(); e.dataTransfer.dropEffect = 'move'; const t = e.target.closest('.page-card'); if (t && t !== dragSrcEl) { const g = document.getElementById('pageGrid'); const c = Array.from(g.children); if (c.indexOf(dragSrcEl) < c.indexOf(t)) t.after(dragSrcEl); else t.before(dragSrcEl); } return false; }); item.addEventListener('dragend', function(e) { this.classList.remove('dragging'); const ids = Array.from(document.querySelectorAll('.page-card')).map(el => el.dataset.id); const n = []; ids.forEach(id => { const p = allPages.find(x => x.id === id); if (p) n.push(p); }); allPages = n; }); }
        
        function findCachedThumb(path, index) { const p = allPages.find(x => x.path === path && x.pageIndex === index); if (p) { const el = document.getElementById(`img-${p.id}`); if (el && el.src.startsWith('data:')) return el.src; } return null; }
        
        async function loadReview(path) {
            currentReviewFilePath = path;  // 保存当前预览的文件路径
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
        
        function switchToWorkspace() { if (reviewObserver) reviewObserver.disconnect(); document.getElementById('reviewView').style.display = 'none'; document.getElementById('helpView').style.display = 'none'; document.getElementById('workspaceView').style.display = 'flex'; document.getElementById('btnBackEdit').style.display = 'none'; currentReviewFilePath = null; }
        
        async function printCurrentFile() {
            if (!currentReviewFilePath) {
                alert('没有可打印的文件');
                return;
            }
            
            const result = await pywebview.api.print_pdf(currentReviewFilePath);
            if (result.success) {
                alert(result.message || '已发送到打印机');
            } else {
                alert('打印失败: ' + result.error);
            }
        }
        
        function showHelp() {
            document.getElementById('workspaceView').style.display = 'none';
            document.getElementById('reviewView').style.display = 'none';
            document.getElementById('helpView').style.display = 'block';
            document.getElementById('btnBackEdit').style.display = 'block';
        }
        
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