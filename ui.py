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
        .batch-btn.train-ticket { color: #a9dfbf; border-color: rgba(169, 223, 191, 0.4); }
        .batch-btn.train-ticket:hover { background: rgba(169, 223, 191, 0.12); }

        .ticket-date-group { margin-bottom: 22px; }
        .ticket-date-heading { color: #2c3e50; font-size: 15px; margin-bottom: 8px; }
        .ticket-table { width: 100%; border-collapse: collapse; }
        .ticket-table th, .ticket-table td { border: 1px solid #ddd; padding: 9px 10px; text-align: left; font-size: 13px; }
        .ticket-table th { background: #f4f6f7; color: #34495e; }
        .ticket-table td.amount, .ticket-table th.amount { text-align: right; }
        .ticket-table tfoot td { background: #f8f9fa; font-weight: 600; }
        .ticket-error-list { background: #fff3cd; border-left: 4px solid #f39c12; color: #795d12; padding: 10px 14px; margin-bottom: 16px; font-size: 13px; }
        .ticket-grand-total { text-align: right; color: #2c3e50; font-size: 16px; font-weight: 600; }
        .size-option-list { border: 1px solid #ddd; border-radius: 4px; overflow: hidden; }
        .size-option { display: flex; align-items: center; justify-content: space-between; gap: 12px; padding: 11px 12px; border-bottom: 1px solid #eee; cursor: pointer; }
        .size-option:last-child { border-bottom: none; }
        .size-option:hover { background: #f8f9fa; }
        .size-option-main { display: flex; align-items: center; gap: 9px; color: #2c3e50; }
        .size-option-count { color: #7f8c8d; font-size: 12px; }

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
        
        /* 可编辑单元格样式 */
        .editable-cell { cursor: text; position: relative; }
        .editable-cell:hover { background-color: #f0f8ff; }
        .editable-cell input { width: 100%; border: none; padding: 4px; background: transparent; font-size: inherit; }
        .edited-row { background-color: #fff3cd !important; }
        
        /* 双击预览提示 */
        .result-table tbody tr { cursor: pointer; }
        .result-table tbody tr:hover { background-color: #f5f5f5; }
        
        /* 校验信息框样式 */
        .validation-info { 
            margin-bottom: 15px; 
            padding: 12px 15px; 
            background: #e3f2fd; 
            border: 1px solid #2196f3; 
            border-radius: 4px; 
            color: #1565c0; 
            font-size: 13px;
            line-height: 1.6;
        }
        .validation-info.warning { 
            background: #fff3e0; 
            border-color: #ff9800; 
            color: #e65100; 
        }
        .validation-info.error { 
            background: #ffebee; 
            border-color: #f44336; 
            color: #c62828; 
        }
        .validation-info strong { 
            display: block; 
            margin-bottom: 5px; 
            font-size: 14px;
        }
        
        /* 修改提示样式 */
        .edit-notice {
            margin-top: 15px;
            padding: 12px 15px;
            background: #fff9e6;
            border: 1px solid #ffd54f;
            border-radius: 4px;
            color: #f57c00;
            font-size: 13px;
        }
        .edit-notice strong {
            display: block;
            margin-bottom: 8px;
        }
        .edit-notice ul {
            margin: 5px 0 0 20px;
            padding: 0;
        }
        .edit-notice li {
            margin: 3px 0;
        }

        .invoice2-view { flex: 1; min-width: 0; overflow: hidden; background: #f4f6f7; display: none; }
        .invoice2-shell { width: 100%; height: 100%; display: grid; grid-template-columns: 252px minmax(0, 1fr); }
        .invoice2-sidebar { min-width: 0; overflow: hidden; background: white; border-right: 1px solid #d7dde1; padding: 15px 14px; }
        .invoice2-main { min-width: 0; min-height: 0; overflow: hidden; display: flex; flex-direction: column; padding: 14px 18px 10px; }
        .invoice2-header { display: flex; flex-direction: column; gap: 10px; margin-bottom: 13px; }
        .invoice2-header h1 { color: #243342; font-size: 19px; font-weight: 650; }
        .invoice2-header p { color: #6b7780; margin-top: 3px; font-size: 12px; }
        .invoice2-mode { width: 100%; display: inline-flex; border: 1px solid #cfd6dc; border-radius: 6px; overflow: hidden; background: white; flex-shrink: 0; }
        .invoice2-mode button { flex: 1; border: 0; border-right: 1px solid #cfd6dc; background: white; color: #4d5a63; padding: 7px 9px; cursor: pointer; }
        .invoice2-mode button:last-child { border-right: 0; }
        .invoice2-mode button.active { background: #2c3e50; color: white; }
        .invoice2-form { background: white; }
        .invoice2-grid { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 9px 10px; }
        .invoice2-field { grid-column: span 1; min-width: 0; }
        .invoice2-field.wide { grid-column: 1 / -1; }
        .invoice2-field label { display: block; color: #53606a; font-size: 11px; font-weight: 600; margin-bottom: 4px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
        .invoice2-input-row { display: flex; min-width: 0; }
        .invoice2-field input { width: 100%; min-width: 0; height: 32px; border: 1px solid #cfd6dc; padding: 0 8px; font-size: 12px; background: white; }
        .invoice2-input-row input { border-radius: 4px 0 0 4px; }
        .invoice2-input-row button { width: 52px; border: 1px solid #cfd6dc; border-left: 0; border-radius: 0 4px 4px 0; background: #f5f7f8; color: #34495e; cursor: pointer; }
        .invoice2-field > input { border-radius: 4px; }
        .invoice2-actions { border-top: 1px solid #edf0f2; margin-top: 11px; padding-top: 8px; }
        .invoice2-actions span { display: block; min-height: 17px; line-height: 17px; }
        .invoice2-actions button, .invoice2-result-actions button { border: 0; border-radius: 4px; padding: 9px 16px; cursor: pointer; font-size: 13px; }
        .invoice2-primary { background: #21865b; color: white; }
        .invoice2-secondary { background: #e8edf0; color: #34495e; }
        .invoice2-results { display: none; height: 100%; min-height: 0; flex-direction: column; }
        .invoice2-result-empty { flex: 1; min-height: 0; display: flex; flex-direction: column; align-items: center; justify-content: center; color: #879199; text-align: center; }
        .invoice2-result-empty strong { color: #53616b; font-size: 16px; font-weight: 600; margin-bottom: 8px; }
        .invoice2-result-empty span { font-size: 12px; }
        .invoice2-kpis { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); border: 1px solid #dce1e5; border-radius: 6px; overflow: hidden; background: white; }
        .invoice2-kpi { padding: 10px 14px; border-right: 1px solid #e5e9ec; min-width: 0; }
        .invoice2-kpi:nth-child(3n) { border-right: 0; }
        .invoice2-kpi:nth-child(n+4) { border-top: 1px solid #e5e9ec; }
        .invoice2-kpi span { display: block; color: #74808a; font-size: 10px; margin-bottom: 4px; }
        .invoice2-kpi strong { display: block; color: #263746; font-size: 16px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
        .invoice2-kpi.total strong { color: #21865b; }
        .invoice2-result-head { display: flex; align-items: center; justify-content: space-between; gap: 16px; margin: 12px 0 7px; flex-shrink: 0; }
        .invoice2-result-head h2 { color: #2c3e50; font-size: 15px; }
        .invoice2-result-actions { display: flex; gap: 8px; }
        .invoice2-tabs { display: flex; gap: 2px; border-bottom: 1px solid #cfd6dc; flex-shrink: 0; }
        .invoice2-tabs button { border: 0; background: transparent; color: #66737c; padding: 7px 13px; cursor: pointer; border-bottom: 2px solid transparent; }
        .invoice2-tabs button.active { color: #1d6f4e; border-bottom-color: #21865b; font-weight: 600; }
        .invoice2-table-wrap { flex: 1; min-height: 0; background: white; border: 1px solid #dce1e5; border-top: 0; overflow: auto; }
        .invoice2-table { width: 100%; border-collapse: collapse; font-size: 12px; }
        .invoice2-table th { position: sticky; top: 0; z-index: 1; background: #eef2f4; color: #44515a; text-align: left; padding: 9px 10px; white-space: nowrap; }
        .invoice2-table td { border-top: 1px solid #edf0f2; padding: 9px 10px; color: #4b5962; vertical-align: top; }
        .invoice2-table td.amount { text-align: right; font-variant-numeric: tabular-nums; white-space: nowrap; }
        .invoice2-table td.filename { max-width: 260px; word-break: break-all; }
        .invoice2-empty { padding: 34px; text-align: center; color: #879199; }
        .invoice2-files { color: #66737c; font-size: 11px; margin-top: 6px; min-height: 16px; flex-shrink: 0; word-break: break-all; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
        @media (max-width: 820px) {
            .invoice2-view { overflow-y: auto; }
            .invoice2-shell { height: auto; display: block; }
            .invoice2-sidebar { border-right: 0; border-bottom: 1px solid #d7dde1; }
            .invoice2-main { min-height: 460px; overflow-y: auto; padding: 18px 16px 60px; }
            .invoice2-grid { grid-template-columns: minmax(0, 1fr); }
            .invoice2-field, .invoice2-field.wide { grid-column: 1; }
            .invoice2-kpis { grid-template-columns: repeat(2, minmax(0, 1fr)); }
            .invoice2-kpi, .invoice2-kpi:nth-child(3n) { border-right: 1px solid #e5e9ec; }
            .invoice2-kpi:nth-child(2n) { border-right: 0; }
            .invoice2-kpi:nth-child(n+3) { border-top: 1px solid #e5e9ec; }
            .invoice2-table-wrap { max-height: 430px; min-height: 260px; }
        }
    </style>
</head>
<body>
    <div class="toolbar">
        <button id="btnAddFiles" onclick="addFiles()">+ 添加PDF文件</button>
        <div class="separator"></div>
        <div class="mode-selector">
            <label><input type="radio" name="mergeMode" value="normal" checked onchange="switchAppMode()"> 普通</label>
            <label><input type="radio" name="mergeMode" value="invoice" onchange="switchAppMode()"> 发票</label>
            <label><input type="radio" name="mergeMode" value="invoice2" onchange="switchAppMode()"> 发票2</label>
            <button id="btnConfig" class="config-btn warning" onclick="openRouteConfig()">配置路线</button>
        </div>
        <div class="separator"></div>
        <button id="btnClearAll" onclick="clearAll()" class="danger">清空全部</button>
        <div style="flex:1"></div>
        <button id="btnHelp" onclick="showHelp()" style="background-color: #95a5a6; color: white; border: none; padding: 8px 16px; border-radius: 4px; cursor: pointer; font-size: 14px; margin-right: 10px;">帮助</button>
        <button id="btnDirectPrint" onclick="directPrintInvoices()" class="secondary" style="display:none;">打印</button>
        <button id="btnStartMerge" onclick="startMerge()" class="success">合并并保存</button>
        <button onclick="switchToWorkspace()" id="btnBackEdit" style="display:none; margin-left: 10px;">&lt; 返回编辑</button>
    </div>

    <div class="main-content">
        <div class="invoice2-view" id="invoice2View">
            <div class="invoice2-shell">
                <aside class="invoice2-sidebar">
                    <div class="invoice2-header">
                        <div><h1>发票全流程</h1><p>下载、清洗、分类与报销参数</p></div>
                        <div class="invoice2-mode">
                            <button id="invoice2LocalMode" class="active" onclick="setInvoice2Mode('local')">本地文件夹</button>
                            <button id="invoice2EmailMode" onclick="setInvoice2Mode('email')">邮箱下载</button>
                        </div>
                    </div>
                    <div class="invoice2-form">
                        <div class="invoice2-grid">
                            <div class="invoice2-field wide" id="invoice2SourceField">
                                <label>发票源目录</label><div class="invoice2-input-row"><input id="invoice2SourceDir" placeholder="选择包含 PDF 的目录"><button onclick="browseInvoice2Dir('invoice2SourceDir')">浏览</button></div>
                            </div>
                            <div class="invoice2-field wide" id="invoice2RawField" style="display:none">
                                <label>附件下载目录</label><div class="invoice2-input-row"><input id="invoice2RawDir" value="./tem"><button onclick="browseInvoice2Dir('invoice2RawDir')">浏览</button></div>
                            </div>
                            <div class="invoice2-field" id="invoice2StartField" style="display:none"><label>开始日期</label><input type="date" id="invoice2StartDate"></div>
                            <div class="invoice2-field" id="invoice2EndField" style="display:none"><label>结束日期</label><input type="date" id="invoice2EndDate"></div>
                            <div class="invoice2-field wide">
                                <label>分类输出目录</label><div class="invoice2-input-row"><input id="invoice2OutputDir" placeholder="选择结果保存目录"><button onclick="browseInvoice2Dir('invoice2OutputDir')">浏览</button></div>
                            </div>
                            <div class="invoice2-field"><label>单位所在城市</label><input id="invoice2UnitCity" value="郑州"></div>
                            <div class="invoice2-field"><label>家到高铁站最低费用</label><input type="number" id="invoice2HomeMin" value="20" min="0"></div>
                            <div class="invoice2-field"><label>家到高铁站最高费用</label><input type="number" id="invoice2HomeMax" value="30" min="0"></div>
                            <div class="invoice2-field"><label>高铁站到项目地最低费用</label><input type="number" id="invoice2ProjectMin" value="30" min="0"></div>
                        </div>
                        <div class="invoice2-actions"><span id="invoice2FormStatus" style="color:#71808a;font-size:12px;"></span></div>
                    </div>
                </aside>
                <section class="invoice2-main">
                    <div class="invoice2-result-empty" id="invoice2ResultEmpty">
                        <strong>等待处理结果</strong><span>完成左侧参数后，统计和发票明细将在这里显示</span>
                    </div>
                    <div class="invoice2-results" id="invoice2Results">
                        <div class="invoice2-kpis" id="invoice2Kpis"></div>
                        <div class="invoice2-result-head">
                            <h2>发票统计与明细</h2>
                            <div class="invoice2-result-actions"><button class="invoice2-secondary" onclick="sendInvoice2ToPrint()">送入两张一页打印</button></div>
                        </div>
                        <div class="invoice2-tabs" id="invoice2Tabs"></div>
                        <div class="invoice2-table-wrap"><table class="invoice2-table"><thead><tr><th>类别</th><th>日期</th><th>内容</th><th>金额</th><th>发票号</th><th>文件名</th></tr></thead><tbody id="invoice2TableBody"></tbody></table></div>
                        <div class="invoice2-files" id="invoice2Files"></div>
                    </div>
                </section>
            </div>
        </div>
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
                        <p>发票打印工具 - 专业版 v3.3</p>
                        <p>让发票处理更简单、更高效</p>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <div class="batch-toolbar" id="batchToolbar">
        <span id="batchCount">已选 0 项</span>
        <div style="width:1px; height:15px; background:rgba(255,255,255,0.3)"></div>
        <button class="batch-btn" onclick="selectAllPages()">全选</button>
        <button class="batch-btn" onclick="invertSelection()">反选</button>
        <div style="width:1px; height:15px; background:rgba(255,255,255,0.3)"></div>
        <button class="batch-btn" onclick="batchRotate(-90)">↺ 左旋</button>
        <button class="batch-btn" onclick="batchRotate(90)">↻ 右旋</button>
        <div style="width:1px; height:15px; background:rgba(255,255,255,0.3)"></div>
        <button class="batch-btn" id="btnUniformSize" onclick="openUniformSizeModal()">统一大小</button>
        <button class="batch-btn fill-form" id="btnFillForm" onclick="openDateModal()" style="display:none">📝 报销单</button>
        <button class="batch-btn train-ticket" id="btnTrainTicket" onclick="generateTrainTicketForm()" style="display:none">🚄 高铁票表单</button>
        <button class="batch-btn info" id="btnStatistics" onclick="statisticsAmount()" style="display:none">💰 统计</button>
        <button class="batch-btn delete" onclick="batchDelete()">删除所选</button>
    </div>

    <!-- 统一页面大小 Modal -->
    <div class="common-modal" id="uniformSizeModal">
        <div class="modal-content" style="width: 460px; height: auto;">
            <div class="modal-header"><h3>统一页面大小</h3><button onclick="closeModal('uniformSizeModal')">×</button></div>
            <div class="modal-body">
                <div id="sizeCategoryList" class="size-option-list"></div>
            </div>
            <div class="modal-footer">
                <button onclick="closeModal('uniformSizeModal')" class="secondary">取消</button>
                <button onclick="applyUniformSize()" class="success">确定</button>
            </div>
        </div>
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
                    <thead><tr><th>序号</th><th>发票号</th><th>来源</th><th>人数</th><th>日期</th><th>起点</th><th>终点</th><th>票额</th></tr></thead>
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

    <!-- 高铁票表单 Modal -->
    <div class="common-modal" id="trainTicketModal">
        <div class="modal-content" style="width: 860px; max-width: calc(100vw - 40px);">
            <div class="modal-header"><h3>高铁票表单</h3><button onclick="closeModal('trainTicketModal')">×</button></div>
            <div class="modal-body"></div>
            <div class="modal-footer">
                <button onclick="copyTrainTicketTable()" class="warning">复制 (Excel)</button>
                <button onclick="saveTrainTicketCSV()" class="success">保存 CSV</button>
                <button onclick="closeModal('trainTicketModal')" class="secondary">关闭</button>
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
        // ========== 辅助函数：Base64 转 Blob ==========
        // 用于解决浏览器跨域安全限制
        function base64ToBlob(base64Data) {
            // 1. 去掉 data:application/pdf;base64, 前缀（如果有）
            const arr = base64Data.split(',');
            const data = arr.length > 1 ? arr[1] : arr[0];
            
            // 2. 解码 Base64
            const byteString = atob(data);
            const ab = new ArrayBuffer(byteString.length);
            const ia = new Uint8Array(ab);
            
            // 3. 转换为字节数组
            for (let i = 0; i < byteString.length; i++) {
                ia[i] = byteString.charCodeAt(i);
            }
            
            // 4. 生成 Blob 对象
            return new Blob([ab], { type: 'application/pdf' });
        }
        
        // ========== 全局变量 ==========
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
        let currentTrainTicketData = [];
        
        // 记录被编辑的行（用于显示修改提示）
        let editedRows = new Map(); // 改为Map，存储 {rowIndex: {field: {old: xxx, new: xxx}}}
        
        // 记录当前右键激活的列表项
        let activeContextItem = null;
        
        // ========== 请求队列控制器 ==========
        const ImageQueue = {
            queue: [],
            processing: 0,
            maxConcurrent: 2, // ⚡️ 关键：只允许同时加载 2 张图片，防止堵死
            
            // 添加任务
            add: function(taskFn) {
                return new Promise((resolve, reject) => {
                    this.queue.push({
                        fn: taskFn,
                        resolve: resolve,
                        reject: reject
                    });
                    this.run();
                });
            },
            
            // 执行队列
            run: function() {
                if (this.processing >= this.maxConcurrent || this.queue.length === 0) {
                    return;
                }
                
                this.processing++;
                // 取出队列中的第一个任务（先进先出）
                const task = this.queue.shift();
                
                // 执行任务
                task.fn()
                    .then(res => task.resolve(res))
                    .catch(err => task.reject(err))
                    .finally(() => {
                        this.processing--;
                        this.run(); // 递归触发下一个
                    });
            },
            
            // 清空队列（切换文件时调用）
            clear: function() {
                this.queue = [];
                this.processing = 0;
            }
        };
        
        // 记录从哪个模态框跳转到预览的
        let previewSourceModal = null;
        
        // 记录当前预览的文件路径
        let currentReviewFilePath = null;

        let invoice2Mode = 'local';
        let invoice2Result = null;
        let invoice2Category = '全部';

        function generateUUID() { return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, c => { var r = Math.random() * 16 | 0, v = c == 'x' ? r : (r & 0x3 | 0x8); return v.toString(16); }); }

        function switchAppMode() {
            const mode = document.querySelector('input[name="mergeMode"]:checked').value;
            const invoice2 = mode === 'invoice2';
            document.getElementById('invoice2View').style.display = invoice2 ? 'block' : 'none';
            document.querySelector('.left-sidebar').style.display = invoice2 ? 'none' : 'flex';
            document.querySelector('.right-area').style.display = invoice2 ? 'none' : 'flex';
            document.getElementById('btnAddFiles').style.display = invoice2 ? 'none' : 'inline-block';
            document.getElementById('btnClearAll').style.display = invoice2 ? 'none' : 'inline-block';
            const primaryButton = document.getElementById('btnStartMerge');
            primaryButton.style.display = 'inline-block';
            primaryButton.textContent = invoice2 ? '开始处理' : '合并并保存';
            document.getElementById('btnHelp').style.display = invoice2 ? 'none' : 'inline-block';
            document.getElementById('btnDirectPrint').style.display = mode === 'invoice' ? 'inline-block' : 'none';
            document.getElementById('btnBackEdit').style.display = 'none';
            if (invoice2) {
                document.getElementById('batchToolbar').classList.remove('visible');
                document.getElementById('statusText').textContent = invoice2Result ? '发票2处理完成' : '发票2就绪';
            } else {
                switchToWorkspace();
                document.getElementById('statusText').textContent = '就绪';
            }
            toggleConfigBtn();
            updateBatchToolbar();
        }

        function setInvoice2Mode(mode) {
            invoice2Mode = mode;
            document.getElementById('invoice2LocalMode').classList.toggle('active', mode === 'local');
            document.getElementById('invoice2EmailMode').classList.toggle('active', mode === 'email');
            document.getElementById('invoice2SourceField').style.display = mode === 'local' ? 'block' : 'none';
            document.getElementById('invoice2RawField').style.display = mode === 'email' ? 'block' : 'none';
            document.getElementById('invoice2StartField').style.display = mode === 'email' ? 'block' : 'none';
            document.getElementById('invoice2EndField').style.display = mode === 'email' ? 'block' : 'none';
        }

        async function browseInvoice2Dir(inputId) {
            const result = await pywebview.api.select_directory();
            if (result && result.success && result.path) document.getElementById(inputId).value = result.path;
            else if (result && !result.success) alert(result.error);
        }

        function invoice2Options() {
            return {
                mode: invoice2Mode,
                sourceDir: document.getElementById('invoice2SourceDir').value.trim(),
                rawDir: document.getElementById('invoice2RawDir').value.trim(),
                outputDir: document.getElementById('invoice2OutputDir').value.trim(),
                startDate: document.getElementById('invoice2StartDate').value,
                endDate: document.getElementById('invoice2EndDate').value,
                unitCity: document.getElementById('invoice2UnitCity').value.trim() || '郑州',
                homeMin: Number(document.getElementById('invoice2HomeMin').value || 0),
                homeMax: Number(document.getElementById('invoice2HomeMax').value || 0),
                projectMin: Number(document.getElementById('invoice2ProjectMin').value || 0)
            };
        }

        async function runInvoice2() {
            const options = invoice2Options();
            if (!options.outputDir || (invoice2Mode === 'local' && !options.sourceDir)) {
                alert('请选择发票源目录和分类输出目录'); return;
            }
            if (options.homeMin > options.homeMax) { alert('最低费用不能高于最高费用'); return; }
            document.getElementById('invoice2FormStatus').textContent = '正在识别和整理发票';
            showProgress(invoice2Mode === 'email' ? '正在下载并整理发票...' : '正在清洗并分类发票...', 35);
            try {
                const result = await pywebview.api.process_reimbursement(options);
                if (!result.success) { alert('处理失败: ' + (result.error || '未知错误')); return; }
                invoice2Result = result;
                invoice2Category = '全部';
                renderInvoice2Result();
                document.getElementById('invoice2FormStatus').textContent = `完成，共 ${result.totalCount} 张有效发票`;
                document.getElementById('statusText').textContent = '发票2处理完成';
            } catch (error) {
                alert('处理失败: ' + error);
            } finally {
                hideProgress();
            }
        }

        function escapeInvoice2(value) {
            const element = document.createElement('span');
            element.textContent = value == null ? '' : String(value);
            return element.innerHTML;
        }

        function renderInvoice2Result() {
            if (!invoice2Result) return;
            const r = invoice2Result;
            const categories = ['高铁票', '出租票', '住宿票', '其他发票'];
            const kpis = [
                ['有效发票', `${r.totalCount} 张`, 'total'],
                ['总金额', `￥${Number(r.totalAmount).toFixed(2)}`, 'total'],
                ...categories.map(category => [category, `${r.counts[category] || 0} 张 / ￥${Number(r.amounts[category] || 0).toFixed(2)}`, ''])
            ];
            document.getElementById('invoice2Kpis').innerHTML = kpis.map(item =>
                `<div class="invoice2-kpi ${item[2]}"><span>${item[0]}</span><strong title="${item[1]}">${item[1]}</strong></div>`
            ).join('');
            const tabs = ['全部', ...categories];
            document.getElementById('invoice2Tabs').innerHTML = tabs.map(category =>
                `<button class="${invoice2Category === category ? 'active' : ''}" onclick="setInvoice2Category('${category}')">${category}</button>`
            ).join('');
            renderInvoice2Table();
            const names = (r.outputFiles || []).map(file => file.name).join('、');
            document.getElementById('invoice2Files').textContent = `输出：${r.outputDir}${names ? '  |  ' + names : ''}`;
            document.getElementById('invoice2ResultEmpty').style.display = 'none';
            document.getElementById('invoice2Results').style.display = 'flex';
        }

        function setInvoice2Category(category) {
            invoice2Category = category;
            renderInvoice2Result();
        }

        function renderInvoice2Table() {
            const rows = getCurrentInvoice2Rows();
            const body = document.getElementById('invoice2TableBody');
            if (!rows.length) {
                body.innerHTML = '<tr><td colspan="6" class="invoice2-empty">当前分类没有发票</td></tr>';
                return;
            }
            body.innerHTML = rows.map(row => `<tr>
                <td>${escapeInvoice2(row.category)}</td><td>${escapeInvoice2(row.date)}</td>
                <td>${escapeInvoice2(row.content)}</td><td class="amount">￥${Number(row.amount).toFixed(2)}</td>
                <td>${escapeInvoice2(row.invoiceNo)}</td><td class="filename">${escapeInvoice2(row.filename)}</td>
            </tr>`).join('');
        }

        function getCurrentInvoice2Rows() {
            if (!invoice2Result) return [];
            return (invoice2Result.rows || []).filter(
                row => invoice2Category === '全部' || row.category === invoice2Category
            );
        }

        async function sendInvoice2ToPrint() {
            const filePaths = [];
            const seenPaths = new Set();
            getCurrentInvoice2Rows().forEach(row => {
                if (row.path && !seenPaths.has(row.path)) {
                    seenPaths.add(row.path);
                    filePaths.push(row.path);
                }
            });
            if (!filePaths.length) {
                alert('没有可打印的分类发票'); return;
            }
            showProgress('正在载入打印工作台...', 35);
            try {
                const result = await pywebview.api.get_reimbursement_print_files(filePaths);
                if (!result.success || !result.files.length) { alert(result.error || '没有可载入的发票'); return; }
                clearAll();
                sourceFiles = result.files;
                const pages = [];
                result.files.forEach(file => {
                    for (let index = 0; index < file.page_count; index++) pages.push({
                        id: generateUUID(), path: file.path, pageIndex: index, fileName: file.name, rotation: 0
                    });
                });
                allPages = pages;
                document.querySelector('input[name="mergeMode"][value="invoice"]').checked = true;
                switchAppMode();
                renderSourceList();
                document.getElementById('emptyState').style.display = 'none';
                renderPageGrid();
                loadThumbnails(pages);
                updateStats();
            } finally {
                hideProgress();
            }
        }

        function toggleConfigBtn() {
            const mode = document.querySelector('input[name="mergeMode"]:checked').value;
            const btn = document.getElementById('btnConfig');
            const fillBtn = document.getElementById('btnFillForm');
            const trainTicketBtn = document.getElementById('btnTrainTicket');
            const statsBtn = document.getElementById('btnStatistics');
            const uniformSizeBtn = document.getElementById('btnUniformSize');
            if (mode === 'invoice') {
                btn.style.display = 'inline-block';
                uniformSizeBtn.style.display = 'none';
                if (selectedPageIds.size > 0) {
                    fillBtn.style.display = 'inline-block';
                    trainTicketBtn.style.display = 'inline-block';
                    statsBtn.style.display = 'inline-block';
                }
            } else {
                btn.style.display = 'none';
                fillBtn.style.display = 'none';
                trainTicketBtn.style.display = 'none';
                statsBtn.style.display = 'none';
                uniformSizeBtn.style.display = selectedPageIds.size > 0 ? 'inline-block' : 'none';
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
            const selectedPaths = []; 
            const selectedPages = [];
            selectedPageIds.forEach(id => { 
                const p = allPages.find(x => x.id === id); 
                if (p) {
                    selectedPaths.push(p.path);
                    selectedPages.push({
                        path: p.path,
                        pageIndex: p.pageIndex,
                        fileName: p.fileName
                    });
                }
            });
            if (selectedPaths.length === 0) { alert("请先选择发票"); return; }

            showProgress('正在分析发票...', 50);
            setTimeout(async () => {
                const res = await pywebview.api.generate_reimbursement_form(selectedPaths, dateStr);
                hideProgress(); closeModal('dateModal');
                
                if (res.success) {
                    currentTableData = res.rows;
                    editedRows.clear(); // 清空编辑记录
                    
                    const modalBody = document.querySelector('#resultModal .modal-body');
                    modalBody.innerHTML = '';
                    
                    // 显示校验信息
                    const selectedCount = selectedPages.length;
                    const recognizedCount = res.rows.length;
                    const validationDiv = document.createElement('div');
                    
                    let hasIssue = false;
                    let passDetails = [];
                    let issueDetails = [];
                    
                    // 检查数量是否一致
                    if (selectedCount !== recognizedCount) {
                        hasIssue = true;
                        issueDetails.push(`选中 ${selectedCount} 个页面，但只识别出 ${recognizedCount} 个发票`);
                        
                        // 找出未识别的页面
                        const recognizedSources = new Set(res.rows.map(r => r.source));
                        const unrecognized = [];
                        selectedPages.forEach(p => {
                            const pageLabel = `${p.fileName}-P${p.pageIndex + 1}`;
                            if (!recognizedSources.has(pageLabel)) {
                                unrecognized.push(pageLabel);
                            }
                        });
                        
                        if (unrecognized.length > 0) {
                            issueDetails.push(`未识别的页面：${unrecognized.join('、')}`);
                        }
                    } else {
                        passDetails.push(`页面数量一致（${selectedCount}个）`);
                    }
                    
                    // 检查发票号是否有未识别
                    const unrecognizedInvoices = res.rows.filter(r => !r.invoiceNo || r.invoiceNo === '未识别');
                    if (unrecognizedInvoices.length > 0) {
                        hasIssue = true;
                        issueDetails.push(`${unrecognizedInvoices.length} 个发票号未识别`);
                    } else {
                        passDetails.push(`所有发票号已识别`);
                    }
                    
                    // 检查金额是否有0
                    const zeroAmounts = res.rows.filter(r => parseFloat(r.amount) === 0);
                    if (zeroAmounts.length > 0) {
                        hasIssue = true;
                        issueDetails.push(`${zeroAmounts.length} 个发票金额为0`);
                    } else {
                        passDetails.push(`所有发票金额正常`);
                    }
                    
                    if (hasIssue) {
                        validationDiv.className = 'validation-info error';
                        validationDiv.innerHTML = `<strong>⚠️ 校验警告</strong>`;
                        if (passDetails.length > 0) {
                            validationDiv.innerHTML += `<br><span style="color: #2e7d32;">✓ 通过项：${passDetails.join('，')}</span>`;
                        }
                        validationDiv.innerHTML += issueDetails.map(d => `<br>• ${d}`).join('');
                    } else {
                        validationDiv.className = 'validation-info';
                        validationDiv.innerHTML = `<strong>✓ 校验通过</strong><br>` + passDetails.map(d => `• ${d}`).join('<br>');
                    }
                    
                    modalBody.appendChild(validationDiv);
                    
                    // 创建表格容器
                    const tableContainer = document.createElement('div');
                    tableContainer.innerHTML = `
                        <table class="result-table" id="resultTable">
                            <thead><tr><th>序号</th><th>发票号</th><th>来源</th><th>人数</th><th>日期</th><th>起点</th><th>终点</th><th>票额</th></tr></thead>
                            <tbody></tbody>
                            <tfoot></tfoot>
                        </table>
                    `;
                    modalBody.appendChild(tableContainer);
                    
                    const tbody = document.querySelector('#resultTable tbody');
                    const tfoot = document.querySelector('#resultTable tfoot');
                    
                    let totalAmount = 0.0;
                    res.rows.forEach((r, idx) => {
                        const tr = document.createElement('tr');
                        tr.dataset.rowIndex = idx;
                        tr.dataset.edited = 'false';
                        
                        // 双击预览发票 - 在主框架显示
                        tr.ondblclick = () => previewInvoiceInMainFrame(r.source, 'resultModal');
                        
                        // 创建可编辑单元格
                        tr.innerHTML = `
                            <td>${r.id}</td>
                            <td class="editable-cell" data-field="invoiceNo">${r.invoiceNo}</td>
                            <td>${r.source || ''}</td>
                            <td class="editable-cell" data-field="people">${r.people}</td>
                            <td class="editable-cell" data-field="date">${r.date}</td>
                            <td class="editable-cell" data-field="start">${r.start}</td>
                            <td class="editable-cell" data-field="end">${r.end}</td>
                            <td class="editable-cell" data-field="amount">${r.amount}</td>
                        `;
                        
                        // 为可编辑单元格添加点击事件
                        tr.querySelectorAll('.editable-cell').forEach(cell => {
                            cell.onclick = (e) => {
                                e.stopPropagation();
                                makeEditable(e.target, tr, 'reimbursement');
                            };
                        });
                        
                        tbody.appendChild(tr);
                        totalAmount += parseFloat(r.amount);
                    });
                    
                    tfoot.innerHTML = `<tr><td colspan="7">合计</td><td>${totalAmount.toFixed(2)}</td></tr>`;
                    
                    // 显示重复提示
                    if (res.duplicateDetails && res.duplicateDetails.length > 0) {
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
                    
                    // 添加修改提示占位符
                    const editNoticeDiv = document.createElement('div');
                    editNoticeDiv.id = 'editNotice';
                    editNoticeDiv.className = 'edit-notice';
                    editNoticeDiv.style.display = 'none';
                    modalBody.appendChild(editNoticeDiv);
                    
                    document.getElementById('resultModal').style.display = 'flex';
                } else {
                    alert(res.error);
                }
            }, 100);
        }
        
        // 使单元格可编辑
        function makeEditable(cell, row, tableType) {
            if (cell.querySelector('input')) return; // 已经是编辑状态
            
            const originalValue = cell.textContent;
            const field = cell.dataset.field;
            
            const input = document.createElement('input');
            input.type = 'text';
            input.value = originalValue;
            input.style.width = '100%';
            
            cell.textContent = '';
            cell.appendChild(input);
            input.focus();
            input.select();
            
            const finishEdit = () => {
                const newValue = input.value;
                cell.textContent = newValue;
                
                // 更新数据
                const rowIndex = parseInt(row.dataset.rowIndex);
                const dataSource = tableType === 'statistics' ? currentStatisticsData : currentTableData;
                
                if (dataSource[rowIndex]) {
                    const oldValue = dataSource[rowIndex][field];
                    
                    // 只有值真正改变时才记录
                    if (oldValue !== newValue) {
                        dataSource[rowIndex][field] = newValue;
                        
                        // 标记为已编辑
                        row.dataset.edited = 'true';
                        row.classList.add('edited-row');
                        
                        // 记录修改详情
                        if (!editedRows.has(rowIndex)) {
                            editedRows.set(rowIndex, {});
                        }
                        editedRows.get(rowIndex)[field] = {old: oldValue, new: newValue};
                        
                        // 更新修改提示
                        updateEditNotice(tableType);
                    }
                    
                    // 如果是金额字段，重新计算合计
                    if (field === 'amount') {
                        updateTotal(tableType);
                    }
                }
            };
            
            input.onblur = finishEdit;
            input.onkeydown = (e) => {
                if (e.key === 'Enter') {
                    finishEdit();
                } else if (e.key === 'Escape') {
                    cell.textContent = originalValue;
                }
            };
        }
        
        // 更新修改提示
        function updateEditNotice(tableType) {
            const noticeId = tableType === 'statistics' ? 'editNoticeStats' : 'editNotice';
            const noticeDiv = document.getElementById(noticeId);
            if (!noticeDiv) return;
            
            if (editedRows.size === 0) {
                noticeDiv.style.display = 'none';
                return;
            }
            
            const dataSource = tableType === 'statistics' ? currentStatisticsData : currentTableData;
            const editedList = [];
            
            // 字段名称映射
            const fieldNames = {
                'invoiceNo': '发票号',
                'people': '人数',
                'date': '日期',
                'start': '起点',
                'end': '终点',
                'amount': '金额'
            };
            
            editedRows.forEach((changes, idx) => {
                const changeDetails = [];
                for (const [field, values] of Object.entries(changes)) {
                    const fieldName = fieldNames[field] || field;
                    changeDetails.push(`${fieldName}: ${values.old} → ${values.new}`);
                }
                editedList.push(`序号 ${idx + 1}：${changeDetails.join('，')}`);
            });
            
            noticeDiv.innerHTML = `
                <strong>📝 以下数据已被手动修改（${editedRows.size}条）：</strong>
                <ul>
                    ${editedList.map(item => `<li>${item}</li>`).join('')}
                </ul>
            `;
            noticeDiv.style.display = 'block';
        }
        
        // 更新合计
        function updateTotal(tableType) {
            let total = 0.0;
            const dataSource = tableType === 'statistics' ? currentStatisticsData : currentTableData;
            
            dataSource.forEach(r => {
                total += parseFloat(r.amount) || 0;
            });
            
            const tableId = tableType === 'statistics' ? 'statisticsTable' : 'resultTable';
            const tfoot = document.querySelector(`#${tableId} tfoot`);
            
            if (tableType === 'statistics') {
                tfoot.innerHTML = `<tr><td colspan="3">合计（已去重）</td><td>${total.toFixed(2)}</td></tr>`;
            } else {
                // 报销单：序号、发票号、来源、人数、日期、起点、终点、票额 = 8列
                tfoot.innerHTML = `<tr><td colspan="7">合计</td><td>${total.toFixed(2)}</td></tr>`;
            }
        }
        
        // 预览发票（从报销单行）- 旧版本，使用模态框
        async function previewInvoiceFromRow(rowData) {
            if (!rowData.source) {
                alert('无法定位发票来源');
                return;
            }
            
            // 解析来源信息 "文件名-P1"
            const match = rowData.source.match(/^(.+)-P(\\d+)$/);
            if (!match) {
                alert('来源格式错误');
                return;
            }
            
            const fileName = match[1];
            const pageNum = parseInt(match[2]) - 1; // 转换为0-based索引
            
            // 查找对应的文件路径
            const sourceFile = sourceFiles.find(f => f.name === fileName);
            if (!sourceFile) {
                alert('找不到源文件：' + fileName);
                return;
            }
            
            // 打开预览模态框
            document.getElementById('previewModal').style.display = 'flex';
            const img = document.getElementById('previewImage');
            img.style.transform = 'rotate(0deg)';
            currentPreviewZoom = 1.0;
            updatePreviewZoom();
            img.src = '';
            
            // 加载图片
            const res = await pywebview.api.get_page_image(sourceFile.path, pageNum, 5.0);
            if (res.success) {
                img.src = res.image;
            } else {
                alert('加载发票失败');
                closePreview();
            }
        }
        
        // 在主框架中预览发票
        async function previewInvoiceInMainFrame(sourceLabel, fromModal) {
            if (!sourceLabel) {
                alert('无法定位发票来源');
                return;
            }
            
            // 解析来源信息 "文件名-P1"
            const match = sourceLabel.match(/^(.+)-P(\\d+)$/);
            if (!match) {
                alert('来源格式错误');
                return;
            }
            
            const fileName = match[1];
            const pageNum = parseInt(match[2]) - 1;
            
            // 查找对应的文件路径
            const sourceFile = sourceFiles.find(f => f.name === fileName);
            if (!sourceFile) {
                alert('找不到源文件：' + fileName);
                return;
            }
            
            // 记录来源模态框
            previewSourceModal = fromModal;
            
            // 关闭模态框
            closeModal('resultModal');
            closeModal('statisticsModal');
            
            // 切换到review视图
            document.getElementById('workspaceView').style.display = 'none';
            document.getElementById('reviewView').style.display = 'flex';
            document.getElementById('btnBackEdit').style.display = 'block';
            
            const c = document.getElementById('reviewContent');
            c.innerHTML = '<div style="color:white; margin-top:50px;">正在加载发票...</div>';
            
            if (reviewObserver) reviewObserver.disconnect();
            
            c.innerHTML = '';
            currentReviewZoom = 1.0;
            updateReviewZoomUI();
            
            const div = document.createElement('div');
            div.className = 'review-page';
            div.style.width = BASE_WIDTH + 'px';
            div.dataset.path = sourceFile.path;
            div.dataset.index = pageNum;
            
            const img = document.createElement('img');
            img.alt = `${fileName} - P${pageNum + 1}`;
            img.src = '';
            div.innerHTML += '<span class="review-loading">正在加载...</span>';
            
            div.appendChild(img);
            c.appendChild(div);
            
            // 加载高清图（单页预览使用 4.0 清晰度）
            const res = await pywebview.api.get_page_image(sourceFile.path, pageNum, 4.0);
            if (res.success) {
                img.src = res.image;
                img.dataset.status = 'hd';
                const loading = div.querySelector('.review-loading');
                if (loading) loading.remove();
            } else {
                alert('加载发票失败');
                switchToWorkspace();
            }
            
            // 保存当前预览的文件路径
            currentReviewFilePath = sourceFile.path;
        }

        // 报销单复制功能
        function copyReimbursementTable() {
            if (!currentTableData || currentTableData.length === 0) return;
            
            let text = "序号\\t发票号\\t来源\\t人数\\t日期\\t起点\\t终点\\t票额\\n";
            let total = 0.0;
            currentTableData.forEach(r => {
                text += `${r.id}\\t${r.invoiceNo}\\t${r.source || ''}\\t${r.people}\\t${r.date}\\t${r.start}\\t${r.end}\\t${r.amount}\\n`;
                total += parseFloat(r.amount);
            });
            text += `\\t\\t\\t\\t\\t\\t合计\\t${total.toFixed(2)}`;
            
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

        function appendTicketCell(row, value, className = '') {
            const cell = document.createElement('td');
            cell.textContent = value;
            if (className) cell.className = className;
            row.appendChild(cell);
        }

        async function generateTrainTicketForm() {
            const selectedPages = [];
            selectedPageIds.forEach(id => {
                const page = allPages.find(item => item.id === id);
                if (page) selectedPages.push({
                    path: page.path,
                    pageIndex: page.pageIndex,
                    fileName: page.fileName
                });
            });
            if (selectedPages.length === 0) {
                alert('请先选择高铁票');
                return;
            }

            showProgress('正在识别高铁票...', 50);
            try {
                const result = await pywebview.api.generate_train_ticket_form(selectedPages);
                if (!result.success) {
                    alert(result.error || '未识别到高铁票');
                    return;
                }

                currentTrainTicketData = result.rows;
                const modalBody = document.querySelector('#trainTicketModal .modal-body');
                modalBody.innerHTML = '';

                if (result.errors && result.errors.length > 0) {
                    const warning = document.createElement('div');
                    warning.className = 'ticket-error-list';
                    warning.textContent = `有 ${result.errors.length} 个页面未识别：` +
                        result.errors.map(item => `${item.source}（${item.error}）`).join('；');
                    modalBody.appendChild(warning);
                }

                const groups = new Map();
                currentTrainTicketData.forEach(row => {
                    if (!groups.has(row.date)) groups.set(row.date, []);
                    groups.get(row.date).push(row);
                });

                let fareTotal = 0;
                let refundFeeTotal = 0;
                groups.forEach((rows, date) => {
                    const group = document.createElement('section');
                    group.className = 'ticket-date-group';
                    const heading = document.createElement('h4');
                    heading.className = 'ticket-date-heading';
                    heading.textContent = date;
                    group.appendChild(heading);

                    const table = document.createElement('table');
                    table.className = 'ticket-table';
                    table.innerHTML = '<thead><tr><th>日期</th><th>发车时间</th><th>行程</th><th>车次</th><th class="amount">金额（元）</th><th class="amount">退票费（元）</th></tr></thead>';
                    const tbody = document.createElement('tbody');
                    let groupFareTotal = 0;
                    let groupRefundFeeTotal = 0;
                    rows.forEach(item => {
                        const tr = document.createElement('tr');
                        appendTicketCell(tr, item.date);
                        appendTicketCell(tr, item.time);
                        appendTicketCell(tr, item.route);
                        appendTicketCell(tr, item.trainNo || '');
                        appendTicketCell(tr, Number(item.amount).toFixed(2), 'amount');
                        appendTicketCell(tr, Number(item.refundFee || 0).toFixed(2), 'amount');
                        tbody.appendChild(tr);
                        groupFareTotal += Number(item.amount);
                        groupRefundFeeTotal += Number(item.refundFee || 0);
                    });
                    fareTotal += groupFareTotal;
                    refundFeeTotal += groupRefundFeeTotal;
                    table.appendChild(tbody);
                    const footer = document.createElement('tfoot');
                    const footerRow = document.createElement('tr');
                    const label = document.createElement('td');
                    label.colSpan = 4;
                    label.textContent = '小计';
                    footerRow.appendChild(label);
                    appendTicketCell(footerRow, groupFareTotal.toFixed(2), 'amount');
                    appendTicketCell(footerRow, groupRefundFeeTotal.toFixed(2), 'amount');
                    footer.appendChild(footerRow);
                    table.appendChild(footer);
                    group.appendChild(table);
                    modalBody.appendChild(group);
                });

                const total = document.createElement('div');
                total.className = 'ticket-grand-total';
                total.innerHTML = `票价合计（不含退票费）：${fareTotal.toFixed(2)} 元<br>` +
                    `退票费合计：${refundFeeTotal.toFixed(2)} 元<br>` +
                    `总计（含退票费）：${(fareTotal + refundFeeTotal).toFixed(2)} 元`;
                modalBody.appendChild(total);
                document.getElementById('trainTicketModal').style.display = 'flex';
            } finally {
                hideProgress();
            }
        }

        function copyTrainTicketTable() {
            if (!currentTrainTicketData.length) return;
            let text = '日期\\t发车时间\\t行程\\t车次\\t金额\\t退票费\\n';
            let currentDate = '';
            let groupFareTotal = 0;
            let groupRefundFeeTotal = 0;
            let fareTotal = 0;
            let refundFeeTotal = 0;
            currentTrainTicketData.forEach(row => {
                if (currentDate && row.date !== currentDate) {
                    text += `${currentDate}\\t\\t小计\\t\\t${groupFareTotal.toFixed(2)}\\t${groupRefundFeeTotal.toFixed(2)}\\n`;
                    groupFareTotal = 0;
                    groupRefundFeeTotal = 0;
                }
                currentDate = row.date;
                groupFareTotal += Number(row.amount);
                groupRefundFeeTotal += Number(row.refundFee || 0);
                fareTotal += Number(row.amount);
                refundFeeTotal += Number(row.refundFee || 0);
                text += `${row.date}\\t${row.time}\\t${row.route}\\t${row.trainNo || ''}\\t${Number(row.amount).toFixed(2)}\\t${Number(row.refundFee || 0).toFixed(2)}\\n`;
            });
            text += `${currentDate}\\t\\t小计\\t\\t${groupFareTotal.toFixed(2)}\\t${groupRefundFeeTotal.toFixed(2)}\\n`;
            text += `\\t\\t票价合计（不含退票费）\\t\\t${fareTotal.toFixed(2)}\\t\\n`;
            text += `\\t\\t退票费合计\\t\\t\\t${refundFeeTotal.toFixed(2)}\\n`;
            text += `\\t\\t总计（含退票费）\\t\\t${(fareTotal + refundFeeTotal).toFixed(2)}\\t${refundFeeTotal.toFixed(2)}`;
            if (navigator.clipboard && navigator.clipboard.writeText) {
                navigator.clipboard.writeText(text)
                    .then(() => alert('表格已复制到剪贴板，可直接粘贴到 Excel'))
                    .catch(() => fallbackCopy(text));
            } else {
                fallbackCopy(text);
            }
        }

        async function saveTrainTicketCSV() {
            if (!currentTrainTicketData.length) return;
            const path = await pywebview.api.save_csv_dialog('高铁票表单.csv');
            if (!path) return;
            const result = await pywebview.api.save_train_ticket_csv(path, currentTrainTicketData);
            if (result.success) alert('保存成功！'); else alert('保存失败: ' + result.error);
        }

        function getSelectedPageDescriptors() {
            const pages = [];
            selectedPageIds.forEach(id => {
                const page = allPages.find(item => item.id === id);
                if (page) pages.push({
                    clientId: page.id,
                    path: page.path,
                    pageIndex: page.pageIndex,
                    rotation: page.rotation || 0,
                    fileName: page.fileName
                });
            });
            return pages;
        }

        async function openUniformSizeModal() {
            const pages = getSelectedPageDescriptors();
            if (!pages.length) {
                alert('请先选择页面');
                return;
            }
            showProgress('正在识别页面大小...', 45);
            try {
                const result = await pywebview.api.get_page_size_categories(pages);
                if (!result.success) {
                    alert(result.error || '无法识别页面大小');
                    return;
                }
                const list = document.getElementById('sizeCategoryList');
                list.innerHTML = '';
                result.categories.forEach(category => {
                    const label = document.createElement('label');
                    label.className = 'size-option';
                    const main = document.createElement('span');
                    main.className = 'size-option-main';
                    const radio = document.createElement('input');
                    radio.type = 'radio';
                    radio.name = 'uniformPageSize';
                    radio.value = category.id;
                    radio.dataset.shortSide = category.shortSide;
                    radio.checked = category.id === result.defaultCategoryId;
                    const text = document.createElement('span');
                    text.textContent = category.label;
                    main.appendChild(radio);
                    main.appendChild(text);
                    const count = document.createElement('span');
                    count.className = 'size-option-count';
                    count.textContent = `${category.count} 页`;
                    label.appendChild(main);
                    label.appendChild(count);
                    list.appendChild(label);
                });
                document.getElementById('uniformSizeModal').style.display = 'flex';
            } finally {
                hideProgress();
            }
        }

        async function applyUniformSize() {
            const selectedSize = document.querySelector('input[name="uniformPageSize"]:checked');
            if (!selectedSize) return;
            const pages = getSelectedPageDescriptors();
            showProgress('正在统一页面大小...', 55);
            try {
                const result = await pywebview.api.normalize_page_sizes(pages, Number(selectedSize.dataset.shortSide));
                if (!result.success) {
                    alert(result.error || '统一页面大小失败');
                    return;
                }
                const changedPages = [];
                result.pages.forEach(item => {
                    const page = allPages.find(existing => existing.id === item.clientId);
                    if (!page) return;
                    page.path = result.path;
                    page.pageIndex = item.pageIndex;
                    page.normalizedSize = { width: item.width, height: item.height };
                    const image = document.getElementById(`img-${page.id}`);
                    if (image) {
                        image.removeAttribute('src');
                        image.style.opacity = '0.3';
                    }
                    changedPages.push(page);
                });
                closeModal('uniformSizeModal');
                loadThumbnails(changedPages);
                alert(`已统一 ${changedPages.length} 个页面的短边尺寸`);
            } finally {
                hideProgress();
            }
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
                    editedRows.clear(); // 清空编辑记录
                    
                    const modalBody = document.querySelector('#statisticsModal .modal-body');
                    modalBody.innerHTML = '';
                    
                    // 显示校验信息
                    const selectedCount = selectedPages.length;
                    const recognizedCount = res.amounts.length;
                    const validationDiv = document.createElement('div');
                    
                    let hasIssue = false;
                    let passDetails = [];
                    let issueDetails = [];
                    
                    // 检查数量是否一致
                    if (selectedCount !== recognizedCount) {
                        hasIssue = true;
                        issueDetails.push(`选中 ${selectedCount} 个页面，但只识别出 ${recognizedCount} 个发票`);
                        
                        // 找出未识别的页面
                        const recognizedPages = new Set();
                        res.amounts.forEach(item => {
                            item.pages.forEach(page => recognizedPages.add(page));
                        });
                        
                        const unrecognized = [];
                        selectedPages.forEach(p => {
                            const pageLabel = `${p.fileName}-P${p.pageIndex + 1}`;
                            if (!recognizedPages.has(pageLabel)) {
                                unrecognized.push(pageLabel);
                            }
                        });
                        
                        if (unrecognized.length > 0) {
                            issueDetails.push(`未识别的页面：${unrecognized.join('、')}`);
                        }
                    } else {
                        passDetails.push(`页面数量一致（${selectedCount}个）`);
                    }
                    
                    // 检查发票号是否有未识别
                    const unrecognizedInvoices = res.amounts.filter(a => !a.invoiceNo || a.invoiceNo === '未识别');
                    if (unrecognizedInvoices.length > 0) {
                        hasIssue = true;
                        issueDetails.push(`${unrecognizedInvoices.length} 个发票号未识别`);
                    } else {
                        passDetails.push(`所有发票号已识别`);
                    }
                    
                    // 检查金额是否有0
                    const zeroAmounts = res.amounts.filter(a => parseFloat(a.amount) === 0);
                    if (zeroAmounts.length > 0) {
                        hasIssue = true;
                        issueDetails.push(`${zeroAmounts.length} 个发票金额为0`);
                    } else {
                        passDetails.push(`所有发票金额正常`);
                    }
                    
                    if (hasIssue) {
                        validationDiv.className = 'validation-info error';
                        validationDiv.innerHTML = `<strong>⚠️ 校验警告</strong>`;
                        if (passDetails.length > 0) {
                            validationDiv.innerHTML += `<br><span style="color: #2e7d32;">✓ 通过项：${passDetails.join('，')}</span>`;
                        }
                        validationDiv.innerHTML += issueDetails.map(d => `<br>• ${d}`).join('');
                    } else {
                        validationDiv.className = 'validation-info';
                        validationDiv.innerHTML = `<strong>✓ 校验通过</strong><br>` + passDetails.map(d => `• ${d}`).join('<br>');
                    }
                    
                    modalBody.appendChild(validationDiv);
                    
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
                        tr.dataset.rowIndex = index;
                        tr.dataset.edited = 'false';
                        
                        let rowStyle = '';
                        let invoiceDisplay = item.invoiceNo;
                        
                        if (item.isDuplicate) {
                            rowStyle = 'background-color: #fff3cd; color: #856404;';
                            invoiceDisplay += ' ⚠️';
                        }
                        
                        tr.style.cssText = rowStyle;
                        
                        // 创建可编辑单元格
                        tr.innerHTML = `
                            <td>${index + 1}</td>
                            <td class="editable-cell" data-field="invoiceNo">${invoiceDisplay}</td>
                            <td>${item.pages.join(', ')}</td>
                            <td class="editable-cell" data-field="amount">${item.amount}</td>
                        `;
                        
                        if (item.isDuplicate) {
                            tr.title = `重复发票 - 已去重统计`;
                        }
                        
                        // 为可编辑单元格添加点击事件
                        tr.querySelectorAll('.editable-cell').forEach(cell => {
                            cell.onclick = (e) => {
                                e.stopPropagation();
                                makeEditable(e.target, tr, 'statistics');
                            };
                        });
                        
                        // 双击预览发票（非编辑单元格区域）- 在主框架显示
                        tr.style.cursor = 'pointer';
                        tr.ondblclick = (e) => {
                            if (!e.target.classList.contains('editable-cell')) {
                                if (item.pages && item.pages.length > 0) {
                                    previewInvoiceInMainFrame(item.pages[0], 'statisticsModal');
                                }
                            }
                        };
                        
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
                    
                    // 添加修改提示占位符
                    const editNoticeDiv = document.createElement('div');
                    editNoticeDiv.id = 'editNoticeStats';
                    editNoticeDiv.className = 'edit-notice';
                    editNoticeDiv.style.display = 'none';
                    modalBody.appendChild(editNoticeDiv);
                    
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
        function selectAllPages() {
            allPages.forEach(page => selectedPageIds.add(page.id));
            document.querySelectorAll('.page-card').forEach(card => card.classList.add('selected'));
            updateBatchToolbar();
        }
        function invertSelection() {
            const inverted = new Set();
            allPages.forEach(page => { if (!selectedPageIds.has(page.id)) inverted.add(page.id); });
            selectedPageIds = inverted;
            document.querySelectorAll('.page-card').forEach(card => card.classList.toggle('selected', selectedPageIds.has(card.dataset.id)));
            updateBatchToolbar();
        }
        function updateBatchToolbar() { 
            const b = document.getElementById('batchToolbar'); 
            if (selectedPageIds.size > 0) { 
                b.classList.add('visible'); 
                document.getElementById('batchCount').textContent = `已选 ${selectedPageIds.size} 项`; 
                const mode = document.querySelector('input[name="mergeMode"]:checked').value; 
                document.getElementById('btnUniformSize').style.display = (mode === 'normal') ? 'inline-block' : 'none';
                document.getElementById('btnFillForm').style.display = (mode === 'invoice') ? 'inline-block' : 'none'; 
                document.getElementById('btnTrainTicket').style.display = (mode === 'invoice') ? 'inline-block' : 'none';
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
            
            // 加载高清图（单页预览使用 4.0 清晰度）
            loadHighResImage(div, img, 4.0);
        }
        
        async function startMerge() { 
            if (document.querySelector('input[name="mergeMode"]:checked').value === 'invoice2') {
                return runInvoice2();
            }
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

        function printPdfData(pdfData, button, originalText) {
            const blobUrl = URL.createObjectURL(base64ToBlob(pdfData));
            const iframe = document.createElement('iframe');
            iframe.style.position = 'fixed';
            iframe.style.right = '0';
            iframe.style.bottom = '0';
            iframe.style.width = '1px';
            iframe.style.height = '1px';
            iframe.style.border = 'none';
            iframe.style.opacity = '0.01';
            iframe.style.pointerEvents = 'none';

            const cleanup = () => {
                if (document.body.contains(iframe)) document.body.removeChild(iframe);
                URL.revokeObjectURL(blobUrl);
            };

            iframe.onload = () => {
                setTimeout(() => {
                    try {
                        iframe.contentWindow.focus();
                        iframe.contentWindow.print();
                    } catch (error) {
                        console.error('打印调用失败:', error);
                        if (confirm('自动打印被拦截。是否在新窗口打开PDF进行打印？')) {
                            window.open(blobUrl, '_blank');
                        }
                    } finally {
                        if (button) {
                            button.disabled = false;
                            button.innerText = originalText;
                        }
                        // 给系统打印预览保留读取 Blob 的时间。
                        setTimeout(cleanup, 60000);
                    }
                }, 500);
            };
            iframe.onerror = () => {
                cleanup();
                if (button) {
                    button.disabled = false;
                    button.innerText = originalText;
                }
                alert('PDF文件加载失败，无法打印');
            };
            iframe.src = blobUrl;
            document.body.appendChild(iframe);
        }

        async function directPrintInvoices() {
            const button = document.getElementById('btnDirectPrint');
            const originalText = button.innerText;
            const pagesToPrint = selectedPageIds.size > 0
                ? allPages.filter(page => selectedPageIds.has(page.id))
                : allPages.slice();

            if (pagesToPrint.length === 0) {
                alert('请先添加文件或选择要打印的页面');
                return;
            }

            button.disabled = true;
            button.innerText = '处理中...';
            showProgress('正在合并并准备打印...', 50);
            try {
                const pages = pagesToPrint.map(page => ({
                    path: page.path,
                    page_index: page.pageIndex,
                    rotation: page.rotation
                }));
                const result = await pywebview.api.merge_invoice_for_print(pages);
                if (!result.success) throw new Error(result.error || '生成打印文件失败');

                document.getElementById('progressFill').style.width = '100%';
                hideProgress();
                printPdfData(result.data, button, originalText);
            } catch (error) {
                hideProgress();
                button.disabled = false;
                button.innerText = originalText;
                alert('打印失败: ' + error.message);
            }
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
            // ⚡️ 切换文件时，清空之前的等待队列
            ImageQueue.clear();
            
            currentReviewFilePath = path;  // 保存当前预览的文件路径
            
            // 切换到预览视图
            document.getElementById('workspaceView').style.display = 'none';
            document.getElementById('reviewView').style.display = 'flex';
            document.getElementById('btnBackEdit').style.display = 'block';
            
            const c = document.getElementById('reviewContent');
            c.innerHTML = '<div style="color:white; margin-top:50px;">正在获取文件信息...</div>';
            if (reviewObserver) reviewObserver.disconnect();
            
            // 获取文件信息
            const info = await pywebview.api.get_file_info(path);
            if (!info.success) {
                c.innerHTML = `<div style="color:red; margin-top:50px;">错误：${info.error}</div>`;
                return;
            }
            
            c.innerHTML = '';
            currentReviewZoom = 1.0;
            updateReviewZoomUI();
            
            // ========== 统一使用图片方式显示（无论是源文件还是合并文件） ==========
            for (let i = 0; i < info.page_count; i++) {
                const div = document.createElement('div');
                div.className = 'review-page';
                div.style.width = BASE_WIDTH + 'px';
                div.dataset.path = path;
                div.dataset.index = i;
                
                // 尝试从缓存获取缩略图
                let src = findCachedThumb(path, i);
                if (!src && i === 0 && historyCache[path]) {
                    src = historyCache[path];
                }
                
                const img = document.createElement('img');
                img.alt = `Page ${i+1}`;
                
                if (src) {
                    img.src = src;
                    img.dataset.status = 'thumb';
                } else {
                    img.src = '';
                    div.innerHTML += '<span class="review-loading">等待加载...</span>';
                }
                
                div.appendChild(img);
                c.appendChild(div);
            }
            
            // 优先加载前两页
            const prioritizeVisibles = async () => {
                const n = Array.from(c.querySelectorAll('.review-page')).slice(0, 2);
                for (const div of n) {
                    const img = div.querySelector('img');
                    if (!img.src) {
                        const r = await pywebview.api.get_page_image(div.dataset.path, div.dataset.index, 0.5);
                        if (r.success) {
                            img.src = r.image;
                            img.dataset.status = 'thumb';
                        }
                    }
                    loadHighResImage(div, img, 2.0);
                }
            };
            
            prioritizeVisibles();
            
            // 懒加载其他页面
            reviewObserver = new IntersectionObserver((entries, observer) => {
                entries.forEach(entry => {
                    if (entry.isIntersecting) {
                        const div = entry.target;
                        const img = div.querySelector('img');
                        // 只有非高清图才请求
                        if (!img.src || img.dataset.status !== 'hd') {
                            // 调用队列版函数
                            loadHighResImage(div, img, 2.0).then(() => {
                                // 加载成功后取消观察
                                if (img.dataset.status === 'hd') {
                                    observer.unobserve(div);
                                }
                            }).catch(() => {
                                console.log('加载失败，保持观察状态以便重试');
                            });
                        } else {
                            // 已经是高清图，直接取消观察
                            observer.unobserve(div);
                        }
                    }
                });
            }, {
                root: document.getElementById('reviewView'),
                rootMargin: '500px',
                threshold: 0.01
            });
            
            document.querySelectorAll('.review-page').forEach(div => reviewObserver.observe(div));
        }

        async function loadHighResImage(div, img, quality = 2.0) { 
            // 防止重复加载
            if (div.dataset.loading === 'true' || img.dataset.status === 'hd') return Promise.resolve(); 
            div.dataset.loading = 'true'; 
            
            // ⚡️ 将实际的请求逻辑包装成一个函数
            const requestTask = async () => {
                try {
                    const res = await pywebview.api.get_page_image(div.dataset.path, parseInt(div.dataset.index), quality);
                    return res;
                } catch (error) {
                    return { success: false, error: error.message };
                }
            };
            
            // ⚡️ 加入队列，等待调度
            return ImageQueue.add(requestTask).then(res => {
                if (res.success) { 
                    img.src = res.image; 
                    img.dataset.status = 'hd'; 
                    const l = div.querySelector('.review-loading'); 
                    if (l) l.remove(); 
                    delete div.dataset.loading;
                } else {
                    // 加载失败，允许重试（移除 loading 标记）
                    console.error('加载失败:', res.error);
                    delete div.dataset.loading; 
                }
            }).catch(err => {
                console.error('队列执行错误:', err);
                delete div.dataset.loading;
            });
        }
        
        function switchToWorkspace() { 
            if (reviewObserver) reviewObserver.disconnect(); 
            document.getElementById('reviewView').style.display = 'none'; 
            document.getElementById('helpView').style.display = 'none'; 
            document.getElementById('workspaceView').style.display = 'flex'; 
            document.getElementById('btnBackEdit').style.display = 'none'; 
            currentReviewFilePath = null; 
            
            // 如果是从模态框跳转过来的，重新打开模态框
            if (previewSourceModal) {
                document.getElementById(previewSourceModal).style.display = 'flex';
                previewSourceModal = null;
            }
        }
        
        async function printCurrentFile() {
            if (!currentReviewFilePath) {
                alert('没有可打印的文件');
                return;
            }
            
            console.log('🖨️ 准备打印文件:', currentReviewFilePath);
            
            const btn = document.querySelector('.review-toolbar button[onclick="printCurrentFile()"]');
            const originalText = btn ? btn.innerText : '🖨️ 打印';
            if (btn) btn.innerText = '⌛ 处理中...';
            
            try {
                let pdfData = null;
                let blobUrl = null;
                
                // ========== 优先从缓存获取 ==========
                if (window.mergedFilesCache && window.mergedFilesCache[currentReviewFilePath]) {
                    console.log('✅ 从前端缓存获取PDF数据用于打印');
                    const cached = window.mergedFilesCache[currentReviewFilePath];
                    
                    // 使用已有的 Blob URL 或创建新的
                    if (cached.blobUrl) {
                        blobUrl = cached.blobUrl;
                    } else {
                        blobUrl = URL.createObjectURL(cached.blob);
                        cached.blobUrl = blobUrl;
                    }
                } else {
                    // ========== 从服务器获取 ==========
                    console.log('📡 从服务器获取PDF数据用于打印');
                    const result = await pywebview.api.print_pdf(currentReviewFilePath);
                    if (!result.success) {
                        alert('打印失败: ' + result.error);
                        if (btn) btn.innerText = originalText;
                        return;
                    }
                    pdfData = result.data;
                    
                    // 转换为 Blob URL
                    const blob = base64ToBlob(pdfData);
                    blobUrl = URL.createObjectURL(blob);
                }
                
                if (!blobUrl) {
                    alert('没有可以打印的文件');
                    if (btn) btn.innerText = originalText;
                    return;
                }
                
                console.log('✅ PDF 数据已准备，创建打印 iframe');
                
                // 创建 iframe
                const iframe = document.createElement('iframe');
                iframe.style.position = 'fixed';
                iframe.style.right = '0';
                iframe.style.bottom = '0';
                iframe.style.width = '1px';
                iframe.style.height = '1px';
                iframe.style.border = 'none';
                iframe.style.opacity = '0.01';
                iframe.style.pointerEvents = 'none';
                
                // 加载并打印
                iframe.src = blobUrl;
                document.body.appendChild(iframe);
                
                // 定义打印逻辑
                const doPrint = () => {
                    try {
                        iframe.contentWindow.focus();
                        iframe.contentWindow.print();
                        console.log('✅ 打印对话框已调用');
                    } catch (e) {
                        console.error("打印调用被拦截:", e);
                        // 降级方案：如果iframe打印失败，尝试在新窗口打开让用户手动打印
                        if (confirm("自动打印被拦截。是否在新窗口打开PDF进行打印？")) {
                            window.open(blobUrl, '_blank');
                        }
                    }
                };
                
                // 等待加载
                iframe.onload = function() {
                    console.log('✅ iframe 已加载');
                    setTimeout(() => {
                        doPrint();
                        
                        // 恢复按钮
                        if (btn) btn.innerText = originalText;
                        
                        // 延迟清理
                        setTimeout(() => {
                            if (document.body.contains(iframe)) {
                                document.body.removeChild(iframe);
                            }
                            // 只清理非缓存的 URL
                            if (!window.mergedFilesCache || !window.mergedFilesCache[currentReviewFilePath]) {
                                URL.revokeObjectURL(blobUrl);
                            }
                        }, 60000); // 给用户1分钟时间在打印预览框操作
                    }, 500);
                };
                
                // 错误处理
                iframe.onerror = function() {
                    console.error('PDF加载失败');
                    alert('PDF文件加载失败，无法打印');
                    if (document.body.contains(iframe)) {
                        document.body.removeChild(iframe);
                    }
                    if (!window.mergedFilesCache || !window.mergedFilesCache[currentReviewFilePath]) {
                        URL.revokeObjectURL(blobUrl);
                    }
                    if (btn) btn.innerText = originalText;
                };
                
            } catch (e) {
                console.error('打印出错:', e);
                alert('程序错误: ' + e.message);
                if (btn) btn.innerText = originalText;
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
        document.addEventListener('keydown', (e) => {
            if (e.key === 'Escape') closePreview();
            const target = e.target;
            const isEditing = target && (target.matches('input, textarea, select') || target.isContentEditable);
            const modalOpen = Array.from(document.querySelectorAll('.common-modal')).some(modal => modal.style.display === 'flex');
            const workspaceVisible = getComputedStyle(workspace).display !== 'none';
            if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'a' && !isEditing && !modalOpen && workspaceVisible) {
                e.preventDefault();
                selectAllPages();
            }
        });
        async function loadThumbnails(pages) { for (const p of pages) { const el = document.getElementById(`img-${p.id}`); if (el && !el.src.startsWith('data')) { pywebview.api.get_page_image(p.path, p.pageIndex, 0.5).then(res => { if (res.success) { el.src = res.image; el.style.opacity = '1'; } }); } } }
        function clearAll() { allPages = []; document.getElementById('pageGrid').innerHTML = ''; sourceFiles = []; renderSourceList(); clearSelection(); document.getElementById('emptyState').style.display = 'flex'; updateStats(); }
        function updateStats() { document.getElementById('totalStats').textContent = `总页数：${allPages.length}`; }
        function showProgress(t, p) { document.getElementById('progressOverlay').style.display = 'flex'; document.getElementById('progressText').textContent = t; document.getElementById('progressFill').style.width = p + '%'; }
        function hideProgress() { document.getElementById('progressOverlay').style.display = 'none'; document.getElementById('progressFill').style.width = '0%'; }
        const invoice2Today = new Date();
        const invoice2MonthStart = new Date(invoice2Today.getFullYear(), invoice2Today.getMonth(), 1);
        function invoice2DateValue(date) {
            return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
        }
        document.getElementById('invoice2StartDate').value = invoice2DateValue(invoice2MonthStart);
        document.getElementById('invoice2EndDate').value = invoice2DateValue(invoice2Today);
    </script>
</body>
</html>
"""
