        function moveUp() {
            if (selectedFileIndex > 0) {
                // 交换位置
                const temp = pdfFiles[selectedFileIndex];
                pdfFiles[selectedFileIndex] = pdfFiles[selectedFileIndex - 1];
                pdfFiles[selectedFileIndex - 1] = temp;
                selectedFileIndex--;
                updateFileList();
                updateButtonStates();
                updateStatus('文件已上移');

                // 同步到后端
                syncFileOrderToBackend();

                // 更新当前预览的文件路径
                if (currentPDFPath) {
                    currentPDFPath = pdfFiles[selectedFileIndex];
                    // 重新加载预览
                    previewFile(selectedFileIndex);
                }
            }
        }

        function moveDown() {
            if (selectedFileIndex >= 0 && selectedFileIndex < pdfFiles.length - 1) {
                // 交换位置
                const temp = pdfFiles[selectedFileIndex];
                pdfFiles[selectedFileIndex] = pdfFiles[selectedFileIndex + 1];
                pdfFiles[selectedFileIndex + 1] = temp;
                selectedFileIndex++;
                updateFileList();
                updateButtonStates();
                updateStatus('文件已下移');

                // 同步到后端
                syncFileOrderToBackend();

                // 更新当前预览的文件路径
                if (currentPDFPath) {
                    currentPDFPath = pdfFiles[selectedFileIndex];
                    // 重新加载预览
                    previewFile(selectedFileIndex);
                }
            }
        }

        // 同步文件顺序到后端
        async function syncFileOrderToBackend() {
            try {
                // 获取当前文件顺序的索引数组
                const currentOrder = [];
                for (let i = 0; i < pdfFiles.length; i++) {
                    // 找到每个文件在后端原始列表中的索引
                    const result = await pywebview.api.get_file_path(i);
                    if (result) {
                        // 通过文件路径找到原始索引
                        const originalIndex = pdfFiles.findIndex(f => f === result);
                        if (originalIndex !== -1) {
                            currentOrder.push(originalIndex);
                        }
                    }
                }

                // 如果无法确定顺序，直接发送当前顺序
                if (currentOrder.length !== pdfFiles.length) {
                    // 创建从0到length-1的数组，表示当前显示顺序
                    const displayOrder = Array.from({length: pdfFiles.length}, (_, i) => i);
                    const result = await pywebview.api.update_file_order(displayOrder);
                    if (!result.success) {
                        console.error('同步文件顺序失败：', result.error);
                    }
                } else {
                    // 发送重新排序后的索引数组
                    const result = await pywebview.api.update_file_order(currentOrder);
                    if (!result.success) {
                        console.error('同步文件顺序失败：', result.error);
                    }
                }
            } catch (error) {
                console.error('同步文件顺序时出错：', error);
            }
        }