// 导出差异报告功能

// 生成差异报告
function generateDiffReport() {
    if (!file1Workbook || !file2Workbook) {
        showAlert('请先上传两个Excel文件并进行比对', 'danger');
        return;
    }

    try {
        // 创建报告容器
        const reportContainer = document.createElement('div');
        reportContainer.className = 'report-container';

        // 添加报告标题
        const reportTitle = document.createElement('h1');
        reportTitle.textContent = 'Excel文件比对差异报告';
        reportContainer.appendChild(reportTitle);

        // 添加报告生成时间
        const reportTime = document.createElement('p');
        reportTime.className = 'report-time';
        reportTime.textContent = `报告生成时间: ${new Date().toLocaleString()}`;
        reportContainer.appendChild(reportTime);

        // 添加文件信息
        const fileInfo = document.createElement('div');
        fileInfo.className = 'report-file-info';
        fileInfo.innerHTML = `
            <div class="file-info-item">
                <h3>原始文件</h3>
                <p>${document.getElementById('file1-header').textContent}</p>
            </div>
            <div class="file-info-item">
                <h3>对比文件</h3>
                <p>${document.getElementById('file2-header').textContent}</p>
            </div>
        `;
        reportContainer.appendChild(fileInfo);

        // 添加比对摘要
        const summaryInfo = document.createElement('div');
        summaryInfo.className = 'report-summary';
        summaryInfo.innerHTML = `
            <h2>比对摘要</h2>
            <div class="summary-stats">
                <div class="summary-stat">
                    <span>总单元格数:</span>
                    <span>${document.getElementById('total-cells').textContent}</span>
                </div>
                <div class="summary-stat">
                    <span>差异单元格数:</span>
                    <span>${document.getElementById('diff-cells').textContent}</span>
                </div>
                <div class="summary-stat">
                    <span>匹配单元格数:</span>
                    <span>${document.getElementById('match-cells').textContent}</span>
                </div>
                <div class="summary-stat">
                    <span>匹配率:</span>
                    <span>${document.getElementById('match-rate').textContent}</span>
                </div>
            </div>
        `;
        reportContainer.appendChild(summaryInfo);

        // 获取当前工作表名称
        const sheetSelect = document.getElementById('sheet-select');
        const currentSheetName = sheetSelect.options[currentSheetIndex].text;

        // 添加工作表信息
        const sheetInfo = document.createElement('div');
        sheetInfo.className = 'report-sheet-info';
        sheetInfo.innerHTML = `
            <h2>工作表: ${currentSheetName}</h2>
            <p>对比模式: ${document.getElementById('compare-mode').value === 'row-by-row' ? '逐行对比' : '跨行对比'}</p>
            <p>忽略大小写: ${document.getElementById('ignore-case').checked ? '是' : '否'}</p>
            <p>仅比对数值: ${document.getElementById('only-values').checked ? '是' : '否'}</p>
        `;
        reportContainer.appendChild(sheetInfo);

        // 添加差异详情表格
        const diffDetails = document.createElement('div');
        diffDetails.className = 'report-diff-details';
        diffDetails.innerHTML = '<h2>差异详情</h2>';

        // 创建差异表格
        const diffTable = document.createElement('table');
        diffTable.className = 'report-diff-table';

        // 添加表头
        const thead = document.createElement('thead');
        const headerRow = document.createElement('tr');
        headerRow.innerHTML = `
            <th>行号</th>
            <th>列</th>
            <th>原始文件值</th>
            <th>对比文件值</th>
        `;
        thead.appendChild(headerRow);
        diffTable.appendChild(thead);

        // 添加表格内容
        const tbody = document.createElement('tbody');

        // 获取当前工作表数据
        const data1 = window.editableData1 || [];
        const data2 = window.editableData2 || [];

        // 查找并添加差异单元格
        for (let row = 0; row < Math.max(data1.length, data2.length); row++) {
            const rowData1 = data1[row] || [];
            const rowData2 = data2[row] || [];

            for (let col = 0; col < Math.max(rowData1.length, rowData2.length); col++) {
                const val1 = rowData1[col] !== undefined ? rowData1[col] : '';
                const val2 = rowData2[col] !== undefined ? rowData2[col] : '';

                // 根据比对选项判断是否有差异
                let hasDiff = false;
                
                if (document.getElementById('only-values').checked) {
                    // 仅比对数值
                    const num1 = parseFloat(val1);
                    const num2 = parseFloat(val2);
                    if (!isNaN(num1) && !isNaN(num2)) {
                        hasDiff = num1 !== num2;
                    }
                } else if (document.getElementById('ignore-case').checked) {
                    // 忽略大小写
                    hasDiff = String(val1).toLowerCase() !== String(val2).toLowerCase();
                } else {
                    // 默认比对
                    hasDiff = String(val1) !== String(val2);
                }

                if (hasDiff) {
                    const diffRow = document.createElement('tr');
                    diffRow.innerHTML = `
                        <td>${row + 1}</td>
                        <td>${String.fromCharCode(65 + col)}</td>
                        <td>${formatCellValue(val1)}</td>
                        <td>${formatCellValue(val2)}</td>
                    `;
                    tbody.appendChild(diffRow);
                }
            }
        }

        diffTable.appendChild(tbody);
        diffDetails.appendChild(diffTable);
        reportContainer.appendChild(diffDetails);

        // 生成HTML报告
        const reportHTML = `
        <!DOCTYPE html>
        <html lang="zh-CN">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Excel差异比对报告</title>
            <style>
                body {
                    font-family: 'Roboto', sans-serif;
                    line-height: 1.6;
                    color: #333;
                    max-width: 1200px;
                    margin: 0 auto;
                    padding: 20px;
                }
                h1, h2, h3 {
                    color: #4361ee;
                }
                .report-container {
                    background-color: #fff;
                    border-radius: 8px;
                    box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
                    padding: 25px;
                }
                .report-time {
                    color: #6c757d;
                    font-style: italic;
                }
                .report-file-info {
                    display: flex;
                    justify-content: space-between;
                    margin: 20px 0;
                    gap: 20px;
                }
                .file-info-item {
                    flex: 1;
                    padding: 15px;
                    background-color: #f8f9fa;
                    border-radius: 8px;
                }
                .report-summary {
                    margin: 30px 0;
                    padding: 20px;
                    background-color: #f0f4ff;
                    border-radius: 8px;
                }
                .summary-stats {
                    display: flex;
                    flex-wrap: wrap;
                    gap: 20px;
                }
                .summary-stat {
                    flex: 1;
                    min-width: 200px;
                    display: flex;
                    justify-content: space-between;
                    padding: 10px;
                    background-color: #fff;
                    border-radius: 4px;
                    box-shadow: 0 2px 4px rgba(0, 0, 0, 0.05);
                }
                .report-sheet-info {
                    margin: 20px 0;
                    padding: 15px;
                    background-color: #f8f9fa;
                    border-radius: 8px;
                }
                .report-diff-table {
                    width: 100%;
                    border-collapse: collapse;
                    margin: 20px 0;
                }
                .report-diff-table th, .report-diff-table td {
                    border: 1px solid #dee2e6;
                    padding: 10px;
                    text-align: left;
                }
                .report-diff-table th {
                    background-color: #4361ee;
                    color: #fff;
                }
                .report-diff-table tr:nth-child(even) {
                    background-color: #f8f9fa;
                }
                @media print {
                    body {
                        padding: 0;
                    }
                    .report-container {
                        box-shadow: none;
                    }
                }
            </style>
        </head>
        <body>
            ${reportContainer.outerHTML}
        </body>
        </html>
        `;

        // 创建Blob对象
        const blob = new Blob([reportHTML], { type: 'text/html' });

        // 创建下载链接
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `Excel差异报告_${new Date().toISOString().slice(0, 10)}.html`;
        document.body.appendChild(a);
        a.click();

        // 清理
        setTimeout(() => {
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
        }, 100);

        showAlert('差异报告已成功导出', 'success');
    } catch (error) {
        showAlert('导出报告时发生错误: ' + error.message, 'danger');
        console.error('Report generation error:', error);
    }
}

// 保存当前会话状态
function saveSessionState() {
    if (!file1Workbook || !file2Workbook) {
        showAlert('请先上传两个Excel文件', 'danger');
        return;
    }

    try {
        // 创建会话状态对象
        const sessionState = {
            timestamp: new Date().toISOString(),
            file1Name: document.getElementById('file1-header').textContent,
            file2Name: document.getElementById('file2-header').textContent,
            currentSheetIndex: currentSheetIndex,
            compareMode: document.getElementById('compare-mode').value,
            ignoreCase: document.getElementById('ignore-case').checked,
            onlyDiff: document.getElementById('only-diff').checked,
            onlyValues: document.getElementById('only-values').checked,
            savedSheets: window.savedSheets || {}
        };

        // 将当前工作表的编辑内容保存
        const currentSheetName = document.getElementById('sheet-select').options[currentSheetIndex].text;
        if (window.editableData1 && window.editableData2) {
            sessionState.savedSheets[currentSheetName] = {
                data1: window.editableData1,
                data2: window.editableData2
            };
        }

        // 将会话状态转换为JSON字符串
        const sessionJSON = JSON.stringify(sessionState);

        // 创建Blob对象
        const blob = new Blob([sessionJSON], { type: 'application/json' });

        // 创建下载链接
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `Excel比对会话_${new Date().toISOString().slice(0, 19).replace(/:/g, '-')}.json`;
        document.body.appendChild(a);
        a.click();

        // 清理
        setTimeout(() => {
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
        }, 100);

        showAlert('会话状态已成功保存', 'success');
    } catch (error) {
        showAlert('保存会话状态时发生错误: ' + error.message, 'danger');
        console.error('Session save error:', error);
    }
}

// 加载保存的会话状态
function loadSessionState() {
    // 创建文件输入元素
    const fileInput = document.createElement('input');
    fileInput.type = 'file';
    fileInput.accept = '.json';
    fileInput.style.display = 'none';
    document.body.appendChild(fileInput);

    // 监听文件选择
    fileInput.addEventListener('change', function(e) {
        const file = e.target.files[0];
        if (!file) {
            document.body.removeChild(fileInput);
            return;
        }

        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const sessionState = JSON.parse(e.target.result);
                
                // 恢复会话状态
                // 这里需要实现恢复逻辑，包括重新加载Excel文件等
                // 由于需要用户重新上传Excel文件，这里只提供一个提示
                showAlert('会话状态已加载，请重新上传相同的Excel文件', 'info');
                
                // 恢复比对选项
                document.getElementById('compare-mode').value = sessionState.compareMode || 'row-by-row';
                document.getElementById('ignore-case').checked = sessionState.ignoreCase || false;
                document.getElementById('only-diff').checked = sessionState.onlyDiff || false;
                document.getElementById('only-values').checked = sessionState.onlyValues || false;
                
                // 保存会话数据以便后续使用
                window.savedSession = sessionState;
            } catch (error) {
                showAlert('加载会话状态时发生错误: ' + error.message, 'danger');
                console.error('Session load error:', error);
            }
            document.body.removeChild(fileInput);
        };
        reader.onerror = function() {
            showAlert('文件读取失败', 'danger');
            document.body.removeChild(fileInput);
        };
        reader.readAsText(file);
    });

    // 触发文件选择对话框
    fileInput.click();
}