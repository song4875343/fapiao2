/**
 * 垫片层 (Shim Layer) - 简化版
 * 自动适配 pywebview 本地模式和浏览器模式
 * 真正的最小侵入式设计
 */

(function() {
    'use strict';
    
    console.log('🔧 Shim Layer 初始化...');
    
    // 检测运行环境
    const isLocalMode = typeof window.pywebview !== 'undefined';
    
    if (isLocalMode) {
        console.log('✅ 本地模式 (pywebview)');
        return; // 本地模式不需要任何处理
    }
    
    console.log('🌐 浏览器模式 (HTTP API)');
    
    // ==================== 浏览器模式垫片 ====================
    
    // 创建虚拟的 pywebview 对象
    window.pywebview = { api: {} };
    
    // 前端缓存：合并后的文件
    window.mergedFilesCache = window.mergedFilesCache || {};
    
    // ==================== 辅助函数 ====================
    
    function base64ToBlob(base64Data, contentType = 'application/pdf') {
        const arr = base64Data.split(',');
        const data = arr.length > 1 ? arr[1] : arr[0];
        const byteString = atob(data);
        const ab = new ArrayBuffer(byteString.length);
        const ia = new Uint8Array(ab);
        for (let i = 0; i < byteString.length; i++) {
            ia[i] = byteString.charCodeAt(i);
        }
        return new Blob([ab], { type: contentType });
    }
    
    function downloadBlob(blob, filename) {
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        setTimeout(() => URL.revokeObjectURL(url), 1000);
    }
    
    function downloadBase64(base64Content, filename) {
        const byteCharacters = atob(base64Content);
        const byteNumbers = new Array(byteCharacters.length);
        for (let i = 0; i < byteCharacters.length; i++) {
            byteNumbers[i] = byteCharacters.charCodeAt(i);
        }
        const byteArray = new Uint8Array(byteNumbers);
        const blob = new Blob([byteArray], { type: 'text/csv;charset=utf-8;' });
        downloadBlob(blob, filename);
    }
    
    // ==================== 自动生成 API 方法 ====================
    
    // 需要自动转换为 HTTP 调用的方法列表
    const autoMethods = [
        'get_file_info', 'get_page_image', 'get_routes', 'save_routes',
        'generate_reimbursement_form', 'calculate_invoice_amounts',
        'print_pdf', 'clear_files'
    ];
    
    // 自动生成所有方法的 HTTP 调用
    autoMethods.forEach(name => {
        window.pywebview.api[name] = async function(...args) {
            // 将参数转换为对象（假设参数名与位置对应）
            const params = {};
            if (args.length > 0) {
                // 简单处理：第一个参数作为主参数
                const argNames = {
                    'get_file_info': ['file_path'],
                    'get_page_image': ['file_path', 'page_index', 'quality'],
                    'generate_reimbursement_form': ['file_paths', 'date_range'],
                    'calculate_invoice_amounts': ['pages_info'],
                    'print_pdf': ['file_path']
                };
                
                const names = argNames[name] || [];
                args.forEach((arg, i) => {
                    if (names[i]) params[names[i]] = arg;
                });
            }
            
            const response = await fetch(`/api/${name}`, {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify(params)
            });
            return await response.json();
        };
    });
    
    // ==================== 特殊处理的方法 ====================
    
    /**
     * 文件选择：触发浏览器上传
     */
    window.pywebview.api.select_pdfs = async function() {
        const input = document.createElement('input');
        input.type = 'file';
        input.accept = '.pdf';
        input.multiple = true;
        
        return new Promise((resolve) => {
            input.onchange = async (e) => {
                const files = Array.from(e.target.files);
                if (files.length === 0) {
                    resolve([]);
                    return;
                }
                
                try {
                    const formData = new FormData();
                    for (let file of files) {
                        formData.append('files', file);
                    }
                    
                    const response = await fetch('/api/upload_files', {
                        method: 'POST',
                        body: formData
                    });
                    const result = await response.json();
                    resolve(result);
                } catch (error) {
                    console.error('上传失败:', error);
                    alert('文件上传失败: ' + error.message);
                    resolve([]);
                }
            };
            
            input.click();
        });
    };
    
    /**
     * 合并页面：处理文件保存和缓存
     */
    window.pywebview.api.merge_pages = async function(page_list, output_path, mode = 'normal') {
        const response = await fetch('/api/merge_pages', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ page_list, output_path, mode })
        });
        
        const result = await response.json();
        
        if (result.success && result.pdf_content) {
            const fileId = result.output_path;
            const pdfBase64 = `data:application/pdf;base64,${result.pdf_content}`;
            const pdfBlob = base64ToBlob(pdfBase64, 'application/pdf');
            
            // 缓存到前端
            window.mergedFilesCache[fileId] = {
                blob: pdfBlob,
                base64: pdfBase64,
                name: 'merged.pdf',
                timestamp: Date.now()
            };
            
            console.log(`✅ 文件已缓存: ${fileId}`);
            
            // 处理文件保存
            if (output_path && typeof output_path === 'object' && output_path.createWritable) {
                // File System Access API
                try {
                    const writable = await output_path.createWritable();
                    await writable.write(pdfBlob);
                    await writable.close();
                    console.log('✅ 文件已保存到用户选择的位置');
                } catch (err) {
                    console.error('❌ 保存文件失败:', err);
                    downloadBlob(pdfBlob, 'merged.pdf');
                }
            } else if (output_path === 'BROWSER_DOWNLOAD') {
                // 自动下载
                downloadBlob(pdfBlob, 'merged.pdf');
                console.log('✅ 文件已下载');
            }
        }
        
        return result;
    };
    
    /**
     * 保存文件对话框
     */
    window.pywebview.api.save_file_dialog = async function() {
        if ('showSaveFilePicker' in window) {
            try {
                const handle = await window.showSaveFilePicker({
                    suggestedName: 'merged.pdf',
                    types: [{
                        description: 'PDF Files',
                        accept: {'application/pdf': ['.pdf']}
                    }]
                });
                return handle;
            } catch (err) {
                if (err.name === 'AbortError') {
                    return null; // 用户取消
                }
                return 'BROWSER_DOWNLOAD'; // 降级
            }
        }
        return 'BROWSER_DOWNLOAD';
    };
    
    /**
     * CSV 保存对话框
     */
    window.pywebview.api.save_csv_dialog = async function() {
        return 'BROWSER_DOWNLOAD';
    };
    
    /**
     * 保存报销单 CSV
     */
    window.pywebview.api.save_reimbursement_csv = async function(path, rows) {
        const response = await fetch('/api/save_reimbursement_csv', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ path, rows })
        });
        
        const result = await response.json();
        
        if (result.success && result.content) {
            downloadBase64(result.content, result.filename);
        }
        
        return result;
    };
    
    /**
     * 保存统计 CSV
     */
    window.pywebview.api.save_statistics_csv = async function(path, amounts) {
        const response = await fetch('/api/save_statistics_csv', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ path, amounts })
        });
        
        const result = await response.json();
        
        if (result.success && result.content) {
            downloadBase64(result.content, result.filename);
        }
        
        return result;
    };
    
    /**
     * 保存 CSV 数据（通用）
     */
    window.pywebview.api.save_csv_data = async function(path, rows) {
        return await window.pywebview.api.save_reimbursement_csv(path, rows);
    };
    
    console.log('✅ 垫片层加载完成');
    console.log('📝 所有 API 调用将通过 HTTP 转发到服务器');
    
})();
