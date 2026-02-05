/**
 * 垫片层 (Shim Layer)
 * 自动适配 pywebview 本地模式和浏览器模式
 * 最小侵入式设计：前端代码无需修改
 */

(function() {
    'use strict';
    
    console.log('🔧 Shim Layer 初始化...');
    
    // 检测运行环境
    const isLocalMode = typeof window.pywebview !== 'undefined';
    
    if (isLocalMode) {
        console.log('✅ 本地模式 (pywebview)');
        // 本地模式：直接使用 pywebview.api
        // 不需要做任何事情，保持原样
        return;
    }
    
    console.log('🌐 浏览器模式 (HTTP API)');
    
    // ==================== 浏览器模式垫片 ====================
    
    // 创建虚拟的 pywebview 对象
    window.pywebview = {
        api: {}
    };
    
    // ==================== 前端缓存管理 ====================
    
    // 合并文件缓存：{file_id: {blob: Blob, base64: string, name: string}}
    window.mergedFilesCache = window.mergedFilesCache || {};
    
    // 辅助函数：Blob 转 Base64
    function blobToBase64(blob) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onloadend = () => resolve(reader.result);
            reader.onerror = reject;
            reader.readAsDataURL(blob);
        });
    }
    
    // 辅助函数：Base64 转 Blob
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
    
    // 文件上传辅助函数
    async function uploadFiles(files) {
        const formData = new FormData();
        for (let file of files) {
            formData.append('files', file);
        }
        
        const response = await fetch('/api/upload_files', {
            method: 'POST',
            body: formData
        });
        
        return await response.json();
    }
    
    // 下载文件辅助函数
    function downloadFile(url, filename) {
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
    }
    
    // 下载 Blob 辅助函数
    function downloadBlob(blob, filename) {
        const url = URL.createObjectURL(blob);
        downloadFile(url, filename);
        setTimeout(() => URL.revokeObjectURL(url), 1000);
    }
    
    // 下载Base64内容为文件
    function downloadBase64(base64Content, filename) {
        const byteCharacters = atob(base64Content);
        const byteNumbers = new Array(byteCharacters.length);
        for (let i = 0; i < byteCharacters.length; i++) {
            byteNumbers[i] = byteCharacters.charCodeAt(i);
        }
        const byteArray = new Uint8Array(byteNumbers);
        const blob = new Blob([byteArray], { type: 'text/csv;charset=utf-8;' });
        const url = URL.createObjectURL(blob);
        downloadFile(url, filename);
        URL.revokeObjectURL(url);
    }
    
    // ==================== API 方法适配 ====================
    
    /**
     * 选择 PDF 文件（使用 HTML5 文件选择器）
     */
    window.pywebview.api.select_pdfs = function() {
        return new Promise((resolve) => {
            const input = document.createElement('input');
            input.type = 'file';
            input.accept = '.pdf';
            input.multiple = true;
            
            input.onchange = async (e) => {
                const files = Array.from(e.target.files);
                if (files.length === 0) {
                    resolve([]);
                    return;
                }
                
                try {
                    const result = await uploadFiles(files);
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
     * 获取页面图片
     */
    window.pywebview.api.get_page_image = async function(file_path, page_index, quality = 1.0) {
        const response = await fetch('/api/get_page_image', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ file_path, page_index, quality })
        });
        return await response.json();
    };
    
    /**
     * 合并页面
     */
    window.pywebview.api.merge_pages = async function(page_list, output_path, mode = 'normal') {
        const response = await fetch('/api/merge_pages', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ page_list, mode })
        });
        
        const result = await response.json();
        
        if (result.success && result.pdf_content) {
            const fileId = result.output_path;
            const pdfBase64 = `data:application/pdf;base64,${result.pdf_content}`;
            
            // 转换为 Blob
            const pdfBlob = base64ToBlob(pdfBase64, 'application/pdf');
            
            // 缓存到前端
            window.mergedFilesCache[fileId] = {
                blob: pdfBlob,
                base64: pdfBase64,
                name: 'merged.pdf',
                timestamp: Date.now(),
                blobUrl: null  // 延迟创建
            };
            
            console.log(`✅ 文件已缓存: ${fileId}, 大小: ${(result.file_size / 1024).toFixed(2)} KB`);
            
            // 处理文件保存
            if (output_path && typeof output_path === 'object' && output_path.createWritable) {
                // 用户选择了保存位置（File System Access API）
                try {
                    const writable = await output_path.createWritable();
                    await writable.write(pdfBlob);
                    await writable.close();
                    console.log('✅ 文件已保存到用户选择的位置');
                } catch (err) {
                    console.error('❌ 保存文件失败:', err);
                    // 降级：自动下载
                    downloadBlob(pdfBlob, 'merged.pdf');
                }
            } else if (output_path === 'browser_download.pdf') {
                // 降级方案：自动下载到默认位置
                downloadBlob(pdfBlob, 'merged.pdf');
                console.log('✅ 文件已下载到浏览器默认位置');
            }
            // 如果 output_path 为 null，说明用户取消了保存
        }
        
        return result;
    };
    
    /**
     * 保存文件对话框（使用 File System Access API）
     */
    window.pywebview.api.save_file_dialog = async function() {
        // 检查浏览器是否支持 File System Access API
        if ('showSaveFilePicker' in window) {
            try {
                const handle = await window.showSaveFilePicker({
                    suggestedName: 'merged.pdf',
                    types: [{
                        description: 'PDF Files',
                        accept: {'application/pdf': ['.pdf']}
                    }]
                });
                
                console.log('✅ 用户选择了保存位置');
                return handle;  // 返回文件句柄
            } catch (err) {
                if (err.name === 'AbortError') {
                    console.log('⚠️ 用户取消了保存');
                    return null;  // 用户取消
                }
                console.error('❌ 文件保存对话框错误:', err);
                return 'browser_download.pdf';  // 降级：自动下载
            }
        } else {
            console.log('⚠️ 浏览器不支持 File System Access API，将自动下载');
            return 'browser_download.pdf';  // 降级：自动下载
        }
    };
    
    /**
     * 获取路线配置
     */
    window.pywebview.api.get_routes = async function() {
        const response = await fetch('/api/get_routes', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' }
        });
        return await response.json();
    };
    
    /**
     * 保存路线配置
     */
    window.pywebview.api.save_routes = async function(routes) {
        const response = await fetch('/api/save_routes', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ routes })
        });
        return await response.json();
    };
    
    /**
     * 生成报销单
     */
    window.pywebview.api.generate_reimbursement_form = async function(file_paths, date_range) {
        const response = await fetch('/api/generate_reimbursement_form', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ file_paths, date_range })
        });
        return await response.json();
    };
    
    /**
     * 计算发票金额
     */
    window.pywebview.api.calculate_invoice_amounts = async function(pages_info) {
        const response = await fetch('/api/calculate_invoice_amounts', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ pages_info })
        });
        return await response.json();
    };
    
    /**
     * 保存 CSV 对话框（浏览器模式直接下载）
     */
    window.pywebview.api.save_csv_dialog = async function() {
        return 'browser_download.csv';
    };
    
    /**
     * 保存报销单 CSV
     */
    window.pywebview.api.save_reimbursement_csv = async function(path, rows) {
        const response = await fetch('/api/save_reimbursement_csv', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ rows })
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
            body: JSON.stringify({ amounts })
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
        // 复用报销单CSV的逻辑
        return await window.pywebview.api.save_reimbursement_csv(path, rows);
    };
    
    /**
     * 打印 PDF
     */
    window.pywebview.api.print_pdf = async function(file_path) {
        const response = await fetch('/api/print_pdf', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ file_path })
        });
        return await response.json();
    };
    
    /**
     * 清空文件（浏览器模式不需要实现）
     */
    window.pywebview.api.clear_files = async function() {
        return true;
    };
    
    /**
     * 获取文件信息
     */
    window.pywebview.api.get_file_info = async function(file_path) {
        const response = await fetch('/api/get_file_info', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ file_path })
        });
        return await response.json();
    };
    
    console.log('✅ 垫片层加载完成');
    console.log('📝 所有 API 调用将通过 HTTP 转发到服务器');
    
})();
