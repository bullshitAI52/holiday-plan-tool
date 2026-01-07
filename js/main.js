// 全局变量
let file1Workbook = null;
let file2Workbook = null;
let currentSheetIndex = 0;

// DOM加载完成后执行
document.addEventListener('DOMContentLoaded', function() {
    // 初始化事件监听器
    initEventListeners();
    
    // 初始化同步滚动
    initSyncScroll();
});

// 初始化所有事件监听器
function initEventListeners() {
    // 文件上传监听
    document.getElementById('file1').addEventListener('change', handleFileUpload);
    document.getElementById('file2').addEventListener('change', handleFileUpload);
    
    // 按钮监听
    document.getElementById('compare-btn').addEventListener('click', compareExcelFiles);
    document.getElementById('reset-btn').addEventListener('click', resetTool);
    document.getElementById('download-btn').addEventListener('click', downloadExcelFiles);
    document.getElementById('history-btn').addEventListener('click', showHistory);
    document.getElementById('sync-btn').addEventListener('click', syncChanges);
    document.getElementById('add-note-btn').addEventListener('click', function() {
        window.annotationMode = true;
        showAlert('请点击要添加批注的单元格', 'info');
    });
    
    // 工作表选择监听
    document.getElementById('sheet-select').addEventListener('change', function() {
        // 保存当前工作表的编辑内容
        if (window.editableData1 && window.editableData2) {
            const currentSheetName = document.getElementById('sheet-select').options[currentSheetIndex].text;
            window.savedSheets = window.savedSheets || {};
            window.savedSheets[currentSheetName] = {
                data1: window.editableData1,
                data2: window.editableData2
            };
        }
        
        currentSheetIndex = parseInt(this.value);
        compareExcelFiles();
    });
}

// 保存比对历史到本地存储
function saveToHistory(data1, data2, sheetName) {
    const history = JSON.parse(localStorage.getItem('excelCompareHistory') || '[]');
    const timestamp = new Date().toISOString();
    
    // 只保留最近60条记录
    if (history.length >= 60) {
        history.shift();
    }
    
    // 确保所有DOM元素存在
    const file1Header = document.getElementById('file1-header');
    const file2Header = document.getElementById('file2-header');
    let compareType = document.querySelector('input[name="compare-type"]:checked');
    const totalCells = document.getElementById('total-cells');
    const diffCells = document.getElementById('diff-cells');
    const matchRate = document.getElementById('match-rate');
    
    // 如果没有选中比较类型，使用默认值
    if (!compareType) {
        compareType = { value: 'row-by-row' };
    }
    
    if (!file1Header || !file2Header || !totalCells || !diffCells || !matchRate) {
        console.warn('无法保存历史记录：缺少必要的DOM元素');
        showAlert('无法保存历史记录：请先上传文件并完成比对', 'warning');
        return;
    }
    
    // 确保统计信息已计算
    if (!totalCells.textContent || !diffCells.textContent || !matchRate.textContent) {
        console.warn('无法保存历史记录：统计信息未计算');
        showAlert('无法保存历史记录：请先完成比对', 'warning');
        return;
    }
    
    history.push({
        timestamp,
        sheetName,
        file1Name: file1Header.textContent,
        file2Name: file2Header.textContent,
        compareType: compareType.value,
        stats: {
            totalCells: totalCells.textContent,
            diffCells: diffCells.textContent,
            matchRate: matchRate.textContent
        }
    });
    
    localStorage.setItem('excelCompareHistory', JSON.stringify(history));
}

// 显示历史记录
function showHistory() {
    const history = JSON.parse(localStorage.getItem('excelCompareHistory') || '[]').sort((a, b) => {
        return new Date(b.timestamp) - new Date(a.timestamp);
    });
    let currentPage = 1;
    const recordsPerPage = 6;
    
    // 创建模态弹窗
    const modal = document.createElement('div');
    modal.className = 'history-modal';
    modal.style.position = 'fixed';
    modal.style.top = '0';
    modal.style.left = '0';
    modal.style.width = '100%';
    modal.style.height = '100%';
    modal.style.backgroundColor = 'rgba(0,0,0,0.5)';
    modal.style.display = 'flex';
    modal.style.justifyContent = 'center';
    modal.style.alignItems = 'center';
    modal.style.zIndex = '1000';
    
    // 创建弹窗内容
    const modalContent = document.createElement('div');
    modalContent.className = 'history-modal-content';
    modalContent.style.backgroundColor = '#f8f9fa';
    modalContent.style.padding = '25px';
    modalContent.style.borderRadius = '8px';
    modalContent.style.maxWidth = '70%';
    modalContent.style.maxHeight = '85%';
    modalContent.style.overflow = 'auto';
    modalContent.style.boxShadow = '0 4px 20px rgba(0,0,0,0.15)';
    modalContent.style.border = '1px solid #dee2e6';
    

    
    function renderHistory(page) {
        const startIndex = (page - 1) * recordsPerPage;
        const endIndex = Math.min(startIndex + recordsPerPage, history.length);
        
        let html = '<div class="history-list" style="display:flex;flex-wrap:wrap;justify-content:space-between;">';
        html += '<h3 style="text-align:center;margin-bottom:15px;width:100%;">比对历史记录</h3>';
        
        for (let i = startIndex; i < endIndex; i++) {
            const record = history[i];
            html += `
            <div class="history-item" style="width:48%;margin-bottom:20px;padding:15px;border-radius:6px;background:#ffffff;box-shadow:0 2px 8px rgba(0,0,0,0.05);border-left:4px solid #4a6baf;">
                <h4 style="color:#4a6baf;margin-bottom:10px;">比对记录 #${i + 1}${i === 0 ? '（最新记录）' : ''}</h4>
                <p><strong>时间:</strong> ${new Date(record.timestamp).toLocaleString()}</p>
                <p><strong>文件1:</strong> ${record.file1Name}</p>
                <p><strong>文件2:</strong> ${record.file2Name}</p>
                <p><strong>工作表:</strong> ${record.sheetName}</p>
                <p><strong>比对统计:</strong> ${record.compareType === 'row-by-row' ? '逐行对比' : '跨行对比'} | 总单元格: ${record.stats.totalCells} | 差异单元格: ${record.stats.diffCells} | 匹配率: ${record.stats.matchRate}</p>
            </div>
            `;
        }
        
        // 添加分页控制
        if (history.length > recordsPerPage) {
            const totalPages = Math.ceil(history.length / recordsPerPage);
            html += '<div style="width:100%;display:flex;justify-content:center;margin-top:15px;">';
            if (page > 1) {
                html += `<button class="page-btn" style="margin:0 10px;padding:5px 15px;border:1px solid #4a6baf;background:white;border-radius:4px;cursor:pointer;" data-action="prev">上一页</button>`;
            }
            html += `<span style="margin:0 10px;line-height:30px;">第 ${page} 页 / 共 ${totalPages} 页</span>`;
            if (page < totalPages) {
                html += `<button class="page-btn" style="margin:0 10px;padding:5px 15px;border:1px solid #4a6baf;background:white;border-radius:4px;cursor:pointer;" data-action="next">下一页</button>`;
            }
            html += '</div>';
        }
        
        html += '</div>';
        modalContent.innerHTML = html;
    }
    
    // 添加分页按钮事件监听
    modalContent.addEventListener('click', function(e) {
        if (e.target.classList.contains('page-btn')) {
            const action = e.target.getAttribute('data-action');
            if (action === 'prev' && currentPage > 1) {
                currentPage--;
            } else if (action === 'next' && currentPage < Math.ceil(history.length / recordsPerPage)) {
                currentPage++;
            }
            renderHistory(currentPage);
            return;
        }
    });

    // 添加点击外部关闭功能
    modal.addEventListener('click', function(e) {
        if (e.target === modal) {
            document.body.removeChild(modal);
        }
    });
    
    // 阻止内容区域点击事件冒泡
    modalContent.addEventListener('click', function(e) {
        e.stopPropagation();
    });
    
    if (history.length === 0) {
        modalContent.innerHTML = '<p>暂无历史记录</p>';
        modal.appendChild(modalContent);
        document.body.appendChild(modal);
        return;
    }
    
    renderHistory(currentPage);
    modal.appendChild(modalContent);
    document.body.appendChild(modal);
}

// 同步修改到原文件
function syncChanges() {
    if (!file1Workbook || !file2Workbook) {
        showAlert('请先上传两个Excel文件', 'danger');
        return;
    }
    
    if (!confirm('确定要将所有修改同步回原始文件吗？此操作不可撤销。')) {
        return;
    }
    
    try {
        const sheetName = document.getElementById('sheet-select').options[currentSheetIndex].text;
        
        // 更新file1Workbook
        const ws1 = XLSX.utils.aoa_to_sheet(window.editableData1);
        file1Workbook.Sheets[sheetName] = ws1;
        
        // 更新file2Workbook
        const ws2 = XLSX.utils.aoa_to_sheet(window.editableData2);
        file2Workbook.Sheets[sheetName] = ws2;
        
        showAlert('修改已成功同步到原始文件', 'success');
    } catch (error) {
        showAlert('同步失败: ' + error.message, 'danger');
        console.error('Sync error:', error);
    }
}

// 初始化批注处理
function initAnnotationHandlers() {
    // 为所有单元格添加点击事件
    document.querySelectorAll('.diff-cell, .match-cell').forEach(cell => {
        cell.addEventListener('click', function() {
            if (window.annotationMode) {
                addAnnotation(this);
                window.annotationMode = false;
            }
        });
    });
}

// 添加批注
function addAnnotation(cellElement) {
    if (!window.annotationMode) {
        return;
    }
    
    const noteText = document.getElementById('note-input').value;
    if (!noteText) {
        showAlert('请输入批注内容', 'warning');
        return;
    }
    
    if (!cellElement) {
        showAlert('请先点击选择一个单元格进行批注', 'warning');
        return;
    }
    
    // 获取当前选中的单元格 - 优先使用传入的cellElement
    let activeCell = cellElement || window.lastSelectedCell;
    
    // 如果没有传入且没有通过window.lastSelectedCell获取到，尝试其他方式
    if (!activeCell) {
        // 优先获取有active类的单元格
        activeCell = document.querySelector('.cell-active, .diff-cell.active, .match-cell.active');
        
        // 如果没有active类，尝试获取当前焦点元素
        if (!activeCell && document.activeElement) {
            activeCell = document.activeElement.closest('.diff-cell, .match-cell, .cell-active');
        }
    }
    
    // 确保activeCell有正确的类名
    if (activeCell && (!activeCell.classList.contains('diff-cell') && 
        !activeCell.classList.contains('match-cell') && 
        !activeCell.classList.contains('cell-active'))) {
        activeCell = null;
    }
    
    if (!activeCell) {
        showAlert('请先点击选择一个单元格进行批注', 'warning');
        return;
    }
    
    // 确保单元格有正确的类名
    if (!activeCell.classList.contains('diff-cell') && 
        !activeCell.classList.contains('match-cell') && 
        !activeCell.classList.contains('cell-active')) {
        showAlert('无法添加批注：无效的单元格类型', 'danger');
        return;
    }
    
    // 移除旧的批注标记（如果存在）
    const oldMarkers = activeCell.querySelectorAll('.note-marker');
    oldMarkers.forEach(marker => {
        activeCell.removeChild(marker);
    });
    
    // 创建新的批注标记
    const noteMarker = document.createElement('div');
    noteMarker.className = 'note-marker';
    noteMarker.innerHTML = '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M12 2C6.48 2 2 6.48 2 12C2 17.52 6.48 22 12 22C17.52 22 22 17.52 22 12C22 6.48 17.52 2 12 2Z" fill="#4a6baf"/><path d="M12 17C11.45 17 11 16.55 11 16V12C11 11.45 11.45 11 12 11C12.55 11 13 11.45 13 12V16C13 16.55 12.55 17 12 17Z" fill="white"/><path d="M12 9C11.45 9 11 8.55 11 8C11 7.45 11.45 7 12 7C12.55 7 13 7.45 13 8C13 8.55 12.55 9 12 9Z" fill="white"/></svg>';
    noteMarker.title = noteText;
    noteMarker.style.position = 'absolute';
    noteMarker.style.top = '2px';
    noteMarker.style.right = '2px';
    noteMarker.style.width = '16px';
    noteMarker.style.height = '16px';
    noteMarker.style.cursor = 'pointer';
    noteMarker.style.transition = 'transform 0.2s';
    
    // 添加悬浮效果
    noteMarker.addEventListener('mouseenter', function() {
        this.style.transform = 'scale(1.2)';
        
        const tooltip = document.createElement('div');
        tooltip.className = 'note-tooltip';
        tooltip.textContent = noteText;
        tooltip.style.position = 'absolute';
        tooltip.style.bottom = '100%';
        tooltip.style.left = '50%';
        tooltip.style.transform = 'translateX(-50%)';
        tooltip.style.backgroundColor = '#4a6baf';
        tooltip.style.color = '#fff';
        tooltip.style.padding = '8px 12px';
        tooltip.style.borderRadius = '6px';
        tooltip.style.boxShadow = '0 2px 8px rgba(0,0,0,0.15)';
        tooltip.style.maxWidth = '200px';
        tooltip.style.whiteSpace = 'pre-wrap';
        tooltip.style.zIndex = '1000';
        tooltip.style.fontSize = '14px';
        tooltip.style.lineHeight = '1.4';
        this.appendChild(tooltip);
    });
    
    noteMarker.addEventListener('mouseleave', function() {
        this.style.transform = 'scale(1)';
        const tooltip = this.querySelector('.note-tooltip');
        if (tooltip) {
            tooltip.style.display = 'none';
            tooltip.style.visibility = 'hidden';
        }
    });
    
    noteMarker.addEventListener('click', function() {
        const tooltip = this.querySelector('.note-tooltip');
        if (tooltip) {
            tooltip.style.display = tooltip.style.display === 'none' ? 'block' : 'none';
            tooltip.style.visibility = tooltip.style.display === 'none' ? 'hidden' : 'visible';
        }
    });
    
    // 添加到单元格并持久化批注数据
    activeCell.style.position = 'relative';
    activeCell.appendChild(noteMarker);
    
    // 保存批注数据到单元格属性
    activeCell.setAttribute('data-note', noteText);
    
    // 清空输入框
    document.getElementById('note-input').value = '';
    showAlert('批注添加成功', 'success');
    
    // 标记单元格为已批注状态
    activeCell.classList.add('annotated');
    
    // 确保批注元素已添加到DOM
    setTimeout(() => {
        if (!activeCell.contains(noteMarker)) {
            activeCell.appendChild(noteMarker);
        }
        // 保持单元格焦点
        activeCell.focus();
    }, 0);
}

// 初始化同步滚动
function initSyncScroll() {
    const file1Scroll = document.getElementById('file1-scroll');
    const file2Scroll = document.getElementById('file2-scroll');
    
    file1Scroll.addEventListener('scroll', function() {
        file2Scroll.scrollTop = this.scrollTop;
        file2Scroll.scrollLeft = this.scrollLeft;
    });
    
    file2Scroll.addEventListener('scroll', function() {
        file1Scroll.scrollTop = this.scrollTop;
        file1Scroll.scrollLeft = this.scrollLeft;
    });
}

// 处理文件上传
function handleFileUpload(e) {
    const file = e.target.files[0];
    if (!file) return;
    
    const fileId = e.target.id;
    const fileInfoElement = document.getElementById(`${fileId}-info`);
    
    fileInfoElement.textContent = `已选择: ${file.name} (${formatFileSize(file.size)})`;
    fileInfoElement.style.color = "var(--primary-color)";
    
    readExcelFile(file, fileId);
}

// 格式化文件大小
function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

// 读取Excel文件
function readExcelFile(file, fileType) {
    const reader = new FileReader();
    
    // 验证文件类型
    if (!file.name.match(/\.(xlsx|xls|csv)$/i)) {
        showAlert('请上传有效的Excel或CSV文件(.xlsx, .xls或.csv)', 'danger');
        return;
    }
    
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            let workbook;
            if (file.name.match(/\.csv$/i)) {
                const csvData = e.target.result;
                workbook = XLSX.utils.book_new();
                const worksheet = XLSX.utils.csv_to_sheet(csvData);
                XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
            } else {
                workbook = XLSX.read(data, { type: 'array' });
            }
            
            if (fileType === 'file1') {
                file1Workbook = workbook;
                document.getElementById('file1-header').textContent = file.name;
            } else {
                file2Workbook = workbook;
                document.getElementById('file2-header').textContent = file.name;
            }
            
            // 检查是否两个文件都已加载
            if (file1Workbook && file2Workbook) {
                document.getElementById('compare-btn').disabled = false;
                prepareSheetSelector();
            }
        } catch (error) {
            showAlert('文件解析失败，请确保上传的是有效的Excel文件', 'danger');
            console.error('Error parsing Excel file:', error);
        }
    };
    
    reader.onerror = function() {
        showAlert('文件读取失败，请重试', 'danger');
    };
    
    reader.readAsArrayBuffer(file);
}

// 准备工作表选择器
function prepareSheetSelector() {
    const sheetSelect = document.getElementById('sheet-select');
    sheetSelect.innerHTML = '';
    
    // 获取两个工作簿的工作表名称交集
    const sheets1 = file1Workbook.SheetNames;
    const sheets2 = file2Workbook.SheetNames;
    const commonSheets = sheets1.filter(name => sheets2.includes(name));
    
    if (commonSheets.length === 0) {
        showAlert('两个Excel文件没有相同名称的工作表', 'danger');
        return;
    }
    
    // 添加选项
    commonSheets.forEach((sheet, index) => {
        const option = document.createElement('option');
        option.value = index;
        option.textContent = sheet;
        sheetSelect.appendChild(option);
    });
    
    // 显示工作表选择器
    document.getElementById('sheet-selector').style.display = 'block';
}

// 比对Excel文件
function compareExcelFiles() {
    if (!file1Workbook || !file2Workbook) {
        showAlert('请先上传两个Excel文件', 'danger');
        return;
    }
    
    const sheetSelect = document.getElementById('sheet-select');
    if (sheetSelect.options.length === 0) {
        showAlert('没有可用的工作表进行比对', 'danger');
        return;
    }
    
    const sheetName = sheetSelect.options[currentSheetIndex].text;
    const sheet1 = file1Workbook.Sheets[sheetName];
    const sheet2 = file2Workbook.Sheets[sheetName];
    
    if (!sheet1 || !sheet2) {
        showAlert('工作表不存在或无法访问', 'danger');
        return;
    }
    
    try {
        // 转换为JSON格式
        const json1 = XLSX.utils.sheet_to_json(sheet1, { header: 1 });
        const json2 = XLSX.utils.sheet_to_json(sheet2, { header: 1 });
        
        // 显示比对结果
        displayComparisonResults(json1, json2, sheetName);
        
        // 初始化批注功能
        initAnnotationHandlers();
        
        // 显示比对摘要
        displayComparisonSummary(json1, json2);
        
        // 显示比对区域
        document.getElementById('comparison-container').style.display = 'flex';
        document.getElementById('summary').style.display = 'block';
        
        // 保存到历史记录
        saveToHistory(json1, json2, sheetName);
        
        // 初始化批注功能
        initAnnotationHandlers();
        
        // 保存到历史记录
        saveToHistory(json1, json2, sheetName);
        
    } catch (error) {
        showAlert('比对过程中发生错误: ' + error.message, 'danger');
        console.error('Comparison error:', error);
    }
}

// 显示比对结果
function displayComparisonResults(data1, data2, sheetName) {
    const table1 = document.getElementById('file1-table');
    const table2 = document.getElementById('file2-table');
    
    // 清空表格
    table1.innerHTML = '';
    table2.innerHTML = '';
    
    // 存储原始数据以便编辑
    window.editableData1 = data1;
    window.editableData2 = data2;
    
    // 确定最大行数和列数
    const maxRows = Math.max(data1.length, data2.length);
    const maxCols = Math.max(
        data1.reduce((max, row) => Math.max(max, row.length), 
        data2.reduce((max, row) => Math.max(max, row.length), 0)
    )); // 修正语法错误，添加缺失的右括号
    
    // 创建表头
    const thead1 = document.createElement('thead');
    const thead2 = document.createElement('thead');
    const headerRow1 = document.createElement('tr');
    const headerRow2 = document.createElement('tr');
    
    // 添加空单元格作为左上角
    const emptyHeader = document.createElement('th');
    emptyHeader.textContent = '';
    headerRow1.appendChild(emptyHeader.cloneNode(true));
    headerRow2.appendChild(emptyHeader.cloneNode(true));
    
    // 添加列标题 (A, B, C...)
    for (let col = 0; col < maxCols; col++) {
        const th1 = document.createElement('th');
        const th2 = document.createElement('th');
        th1.textContent = String.fromCharCode(65 + col);
        th2.textContent = String.fromCharCode(65 + col);
        headerRow1.appendChild(th1);
        headerRow2.appendChild(th2);
    }
    
    thead1.appendChild(headerRow1);
    thead2.appendChild(headerRow2);
    table1.appendChild(thead1);
    table2.appendChild(thead2);
    
    // 创建表格内容
    const tbody1 = document.createElement('tbody');
    const tbody2 = document.createElement('tbody');
    
    let diffCount = 0;
    let matchCount = 0;
    
    for (let row = 0; row < maxRows; row++) {
        const tr1 = document.createElement('tr');
        const tr2 = document.createElement('tr');
        
        // 添加行号
        const rowHeader1 = document.createElement('th');
        const rowHeader2 = document.createElement('th');
        rowHeader1.textContent = row + 1;
        rowHeader2.textContent = row + 1;
        tr1.appendChild(rowHeader1);
        tr2.appendChild(rowHeader2);
        
        // 获取当前行数据 (处理可能不存在的行)
        const rowData1 = data1[row] || [];
        const rowData2 = data2[row] || [];
        
        for (let col = 0; col < maxCols; col++) {
            const td1 = document.createElement('td');
            const td2 = document.createElement('td');
            
            // 获取单元格值 (处理可能不存在的列)
            const val1 = rowData1[col] !== undefined ? rowData1[col] : '';
            const val2 = rowData2[col] !== undefined ? rowData2[col] : '';
            
            td1.textContent = formatCellValue(val1);
            td2.textContent = formatCellValue(val2);
            
            // 使单元格可编辑
            td1.contentEditable = true;
            td2.contentEditable = true;
            
            // 添加编辑事件监听
            td1.addEventListener('blur', function() {
                const newValue = this.textContent;
                window.editableData1[row][col] = newValue;
                // 重新计算并更新整个表格的差异标记
                const table1 = document.getElementById('file1-table');
                const table2 = document.getElementById('file2-table');
                displayComparisonResults(window.editableData1, window.editableData2, 
                    document.getElementById('sheet-select').options[currentSheetIndex].text);
                displayComparisonSummary(window.editableData1, window.editableData2);
            });
            
            td2.addEventListener('blur', function() {
                const newValue = this.textContent;
                window.editableData2[row][col] = newValue;
                // 重新计算并更新整个表格的差异标记
                const table1 = document.getElementById('file1-table');
                const table2 = document.getElementById('file2-table');
                displayComparisonResults(window.editableData1, window.editableData2, 
                    document.getElementById('sheet-select').options[currentSheetIndex].text);
                displayComparisonSummary(window.editableData1, window.editableData2);
            });
            
            // 比较单元格值
            const compareMode = document.getElementById('compare-mode').value;
            if (compareMode === 'row-by-row') {
                // 逐行对比模式
                if (String(val1) !== String(val2)) {
                    td1.classList.add('diff-cell');
                    td2.classList.add('diff-cell');
                    diffCount++;
                } else {
                    td1.classList.add('match-cell');
                    td2.classList.add('match-cell');
                    matchCount++;
                }
            } else {
                // 跨行对比模式
                // 假设电话号码在第2列（索引为1）
                const phoneColIndex = 2;
                
                // 查找匹配的行（基于姓名列）
                const matchingRow1 = data1.find(row => 
                    row.length > 1 && String(row[1]) === String(rowData2[1]));
                const matchingRow2 = data2.find(row => 
                    row.length > 1 && String(row[1]) === String(rowData1[1]));
                
                // 如果找到匹配的行，只比较电话号码列
                if (matchingRow1 || matchingRow2) {
                    const phone1 = matchingRow1 ? matchingRow1[phoneColIndex] : '';
                    const phone2 = matchingRow2 ? matchingRow2[phoneColIndex] : '';
                    
                    // 只标记电话号码列的差异
                    if (col === phoneColIndex && String(phone1) !== String(phone2)) {
                        td1.classList.add('diff-cell');
                        td2.classList.add('diff-cell');
                        diffCount++;
                    } else {
                        td1.classList.add('match-cell');
                        td2.classList.add('match-cell');
                        matchCount++;
                    }
                } else {
                    // 如果没有匹配的行，标记整行为差异
                    if (String(val1) !== String(val2)) {
                        td1.classList.add('diff-cell');
                        td2.classList.add('diff-cell');
                        diffCount++;
                    } else {
                        td1.classList.add('match-cell');
                        td2.classList.add('match-cell');
                        matchCount++;
                    }
                }
            }
            
            tr1.appendChild(td1);
            tr2.appendChild(td2);
        }
        
        tbody1.appendChild(tr1);
        tbody2.appendChild(tr2);
    }
    
    table1.appendChild(tbody1);
    table2.appendChild(tbody2);
}

// 格式化单元格值
function formatCellValue(value) {
    if (value === null || value === undefined) return '';
    if (typeof value === 'string') return value.trim();
    if (typeof value === 'number') {
        // 如果是日期数字 (Excel日期是1900年以来的天数)
        if (value > 10000 && value < 50000) {
            const date = new Date((value - 25569) * 86400 * 1000);
            return date.toLocaleDateString();
        }
        return value;
    }
    return String(value);
}

// 显示比对摘要
function displayComparisonSummary(data1, data2) {
    let totalCells = 0;
    let diffCount = 0;
    let matchCount = 0;
    
    // 计算总单元格数、差异单元格数和匹配单元格数
    const compareMode = document.getElementById('compare-mode').value;
    for (let row = 0; row < Math.max(data1.length, data2.length); row++) {
        const row1 = data1[row] || [];
        const row2 = data2[row] || [];
        
        for (let col = 0; col < Math.max(row1.length, row2.length); col++) {
            totalCells++;
            
            const val1 = row1[col] !== undefined ? row1[col] : '';
            const val2 = row2[col] !== undefined ? row2[col] : '';
            
            if (compareMode === 'row-by-row') {
                // 逐行对比模式
                if (String(val1) !== String(val2)) {
                    diffCount++;
                } else {
                    matchCount++;
                }
            } else {
                // 跨行对比模式
                const rowExistsInSheet1 = data1.some(rowData => 
                    rowData.length === row2.length && 
                    rowData.every((cell, i) => String(cell) === String(row2[i]))
                );
                const rowExistsInSheet2 = data2.some(rowData => 
                    rowData.length === row1.length && 
                    rowData.every((cell, i) => String(cell) === String(row1[i]))
                );
                
                if (rowExistsInSheet1 || rowExistsInSheet2) {
                    // 如果整行数据在另一侧存在，则标记为匹配
                    matchCount++;
                } else {
                    // 如果整行数据不存在，则标记为差异
                    if (String(val1) !== String(val2)) {
                        diffCount++;
                    } else {
                        matchCount++;
                    }
                }
            }
        }
    }
    
    // 更新摘要信息
    document.getElementById('total-cells').textContent = totalCells;
    document.getElementById('diff-cells').textContent = diffCount;
    document.getElementById('match-cells').textContent = matchCount;
    
    // 计算匹配率（百分比）
    const matchRate = totalCells > 0 ? (matchCount / totalCells * 100).toFixed(2) : 0;
    document.getElementById('match-rate').textContent = `${matchRate}%`;
    
    // 更新进度条
    const progressBar = document.getElementById('progress-bar');
    progressBar.style.width = `${matchRate}%`;
}

// 重置工具
function downloadExcelFiles() {
    if (!file1Workbook || !file2Workbook) {
        showAlert('请先上传两个Excel文件', 'danger');
        return;
    }
    
    try {
        // 创建新的工作簿来保存所有编辑后的工作表
        const editedWorkbook1 = XLSX.utils.book_new();
        const editedWorkbook2 = XLSX.utils.book_new();
        const sheetSelect = document.getElementById('sheet-select');
        
        // 获取所有工作表名称
        const allSheets = [];
        for (let i = 0; i < sheetSelect.options.length; i++) {
            allSheets.push(sheetSelect.options[i].text);
        }
        
        // 处理每个工作表
        allSheets.forEach(sheetName => {
            // 获取当前工作表
            const sheet1 = file1Workbook.Sheets[sheetName];
            const sheet2 = file2Workbook.Sheets[sheetName];
            
            // 使用编辑后的数据创建新工作表（如果有保存的编辑数据）
            const editableData1 = window.savedSheets && window.savedSheets[sheetName] ? 
                window.savedSheets[sheetName].data1 : XLSX.utils.sheet_to_json(sheet1, { header: 1 });
            const editableData2 = window.savedSheets && window.savedSheets[sheetName] ? 
                window.savedSheets[sheetName].data2 : XLSX.utils.sheet_to_json(sheet2, { header: 1 });
            
            const editedSheet1 = XLSX.utils.json_to_sheet(editableData1);
            const editedSheet2 = XLSX.utils.json_to_sheet(editableData2);
            
            // 复制原始工作表的全部格式属性到新工作表
            if(sheet1['!cols']) editedSheet1['!cols'] = sheet1['!cols'];
            if(sheet1['!rows']) editedSheet1['!rows'] = sheet1['!rows'];
            if(sheet1['!merges']) editedSheet1['!merges'] = sheet1['!merges'];
            if(sheet1['!margins']) editedSheet1['!margins'] = sheet1['!margins'];
            if(sheet1['!protect']) editedSheet1['!protect'] = sheet1['!protect'];
            if(sheet1['!autofilter']) editedSheet1['!autofilter'] = sheet1['!autofilter'];
            
            if(sheet2['!cols']) editedSheet2['!cols'] = sheet2['!cols'];
            if(sheet2['!rows']) editedSheet2['!rows'] = sheet2['!rows'];
            if(sheet2['!merges']) editedSheet2['!merges'] = sheet2['!merges'];
            if(sheet2['!margins']) editedSheet2['!margins'] = sheet2['!margins'];
            if(sheet2['!protect']) editedSheet2['!protect'] = sheet2['!protect'];
            if(sheet2['!autofilter']) editedSheet2['!autofilter'] = sheet2['!autofilter'];
            
            // 将工作表添加到工作簿
            XLSX.utils.book_append_sheet(editedWorkbook1, editedSheet1, sheetName);
            XLSX.utils.book_append_sheet(editedWorkbook2, editedSheet2, sheetName);
        });
        
        // 确保使用当前显示的表格状态数据
        const currentSheetName = document.getElementById('sheet-select').options[currentSheetIndex].text;
        if (window.editableData1 && window.editableData2) {
            window.savedSheets = window.savedSheets || {};
            window.savedSheets[currentSheetName] = {
                data1: window.editableData1,
                data2: window.editableData2
            };
            
            // 强制更新当前工作表的编辑数据
            const editedSheet1 = XLSX.utils.json_to_sheet(window.editableData1);
            const editedSheet2 = XLSX.utils.json_to_sheet(window.editableData2);
            
            // 复制格式属性
            const sheet1 = file1Workbook.Sheets[currentSheetName];
            const sheet2 = file2Workbook.Sheets[currentSheetName];
            
            if(sheet1['!cols']) editedSheet1['!cols'] = sheet1['!cols'];
            if(sheet1['!rows']) editedSheet1['!rows'] = sheet1['!rows'];
            if(sheet1['!merges']) editedSheet1['!merges'] = sheet1['!merges'];
            if(sheet1['!margins']) editedSheet1['!margins'] = sheet1['!margins'];
            if(sheet1['!protect']) editedSheet1['!protect'] = sheet1['!protect'];
            if(sheet1['!autofilter']) editedSheet1['!autofilter'] = sheet1['!autofilter'];
            
            if(sheet2['!cols']) editedSheet2['!cols'] = sheet2['!cols'];
            if(sheet2['!rows']) editedSheet2['!rows'] = sheet2['!rows'];
            if(sheet2['!merges']) editedSheet2['!merges'] = sheet2['!merges'];
            if(sheet2['!margins']) editedSheet2['!margins'] = sheet2['!margins'];
            if(sheet2['!protect']) editedSheet2['!protect'] = sheet2['!protect'];
            if(sheet2['!autofilter']) editedSheet2['!autofilter'] = sheet2['!autofilter'];
            
            // 更新工作簿中的当前工作表
            editedWorkbook1.Sheets[currentSheetName] = editedSheet1;
            editedWorkbook2.Sheets[currentSheetName] = editedSheet2;
        }
        
        // 创建ZIP文件并添加两个Excel文件
        const zip = new JSZip();
        zip.file('edited_file1.xlsx', XLSX.write(editedWorkbook1, { bookType: 'xlsx', type: 'array' }));
        zip.file('edited_file2.xlsx', XLSX.write(editedWorkbook2, { bookType: 'xlsx', type: 'array' }));
        
        // 生成ZIP文件并下载
        zip.generateAsync({ type: 'blob' }).then(function(content) {
            saveAs(content, 'excel_comparison_all_sheets.zip');
        });
        
    } catch (error) {
        showAlert('下载过程中发生错误: ' + error.message, 'danger');
        console.error('Download error:', error);
    }
}

function resetTool() {
    // 重置文件信息
    document.getElementById('file1').value = '';
    document.getElementById('file2').value = '';
    document.getElementById('file1-info').textContent = '未选择文件';
    document.getElementById('file2-info').textContent = '未选择文件';
    document.getElementById('file1-info').style.color = 'var(--gray-color)';
    document.getElementById('file2-info').style.color = 'var(--gray-color)';
    
    // 重置工作簿
    file1Workbook = null;
    file2Workbook = null;
    
    // 重置按钮状态
    document.getElementById('compare-btn').disabled = true;
    
    // 隐藏工作表选择器
    document.getElementById('sheet-selector').style.display = 'none';
    
    // 清空表格
    document.getElementById('file1-table').innerHTML = '';
    document.getElementById('file2-table').innerHTML = '';
    
    // 隐藏结果区域
    document.getElementById('comparison-container').style.display = 'none';
    document.getElementById('summary').style.display = 'none';
    
    // 重置摘要信息
    document.getElementById('total-cells').textContent = '0';
    document.getElementById('diff-cells').textContent = '0';
    document.getElementById('match-cells').textContent = '0';
    document.getElementById('match-rate').textContent = '0%';
    document.getElementById('progress-bar').style.width = '100%';
}

// 显示提示信息
function showAlert(message, type) {
    // 创建弹窗容器
    const modal = document.createElement('div');
    modal.className = 'modal';
    
    // 创建弹窗内容
    const modalContent = document.createElement('div');
    modalContent.className = `modal-content modal-${type}`;
    
    // 创建关闭按钮
    const closeBtn = document.createElement('span');
    closeBtn.className = 'close-btn';
    closeBtn.innerHTML = '&times;';
    closeBtn.onclick = () => modal.style.display = 'none';
    
    // 创建消息内容
    const messageEl = document.createElement('p');
    messageEl.textContent = message;
    
    // 组装弹窗
    modalContent.appendChild(closeBtn);
    modalContent.appendChild(messageEl);
    modal.appendChild(modalContent);
    
    // 添加到页面
    document.body.appendChild(modal);
    
    // 显示弹窗
    modal.style.display = 'block';
    
    // 3秒后自动消失
    setTimeout(() => {
        modal.style.opacity = '0';
        setTimeout(() => {
            modal.style.display = 'none';
            modal.remove();
        }, 300);
    }, 3000);
}

// 添加弹窗样式
const style = document.createElement('style');
style.textContent = `
.modal {
    display: none;
    position: fixed;
    z-index: 1000;
    left: 0;
    top: 0;
    width: 100%;
    height: 100%;
    background-color: rgba(0,0,0,0.4);
    transition: opacity 0.3s ease;
}

.modal-content {
    position: relative;
    background-color: #fefefe;
    margin: 15% auto;
    padding: 20px;
    border: 1px solid #888;
    width: 80%;
    max-width: 500px;
    border-radius: 8px;
    box-shadow: 0 4px 8px rgba(0,0,0,0.1);
    animation: modalFadeIn 0.3s;
}

@keyframes modalFadeIn {
    from {opacity: 0; transform: translateY(-20px);}
    to {opacity: 1; transform: translateY(0);}
}

.close-btn {
    position: absolute;
    right: 15px;
    top: 10px;
    color: #aaa;
    font-size: 28px;
    font-weight: bold;
    cursor: pointer;
}

.close-btn:hover {
    color: #333;
}

.modal-danger {
    border-left: 5px solid #dc3545;
}

.modal-success {
    border-left: 5px solid #28a745;
}

.modal-warning {
    border-left: 5px solid #ffc107;
}

.modal-info {
    border-left: 5px solid #17a2b8;
}
`;
document.head.appendChild(style);
