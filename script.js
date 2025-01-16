// 存储计算结果的全局变量
let calculationResults = { riverResults: [], areaResults: {} };

// 显示指定的section，隐藏其他section
function showSection(sectionId) {
    // 隐藏所有section
    document.querySelectorAll('.section').forEach(section => {
        section.style.display = 'none';
    });

    // 显示指定的section
    const sectionToDisplay = document.getElementById(sectionId);
    if (sectionToDisplay) {
        sectionToDisplay.style.display = 'block';
    } else {
        console.error(`Section with id ${sectionId} not found.`);
    }
}

// 计算单个水质参数的CWQI
function calculateParameterCWQI(concentration, parameter, riverType) {
    if (concentration === null || concentration === undefined || isNaN(concentration)) {
        return null;
    }

    concentration = parseFloat(concentration);

    switch (parameter) {
        case 'pH':
            return concentration <= 7.0 ? 
                   (7.0 - concentration) / (7.0 - 6.0) : 
                   (concentration - 7.0) / (9.0 - 7.0);
        case 'DO':
            return 5.0 / concentration;
        case 'COD':
            return concentration / 6.0;
        case 'NH3N':
            return concentration / 1.0;
        case 'TP':
            const standardLimit = riverType === 1 ? 0.05 : 0.2;
            return concentration / standardLimit;
        case 'TN':
            return concentration / 1.0;
        default:
            throw new Error(`未知的水质参数: ${parameter}`);
    }
}

// 处理Excel文件
function processFile() {
    const fileInput = document.getElementById('fileInput');
    const output = document.getElementById('output');

    if (!fileInput.files.length) {
        output.innerHTML = '<p style="color: red;">请先选择一个Excel文件</p>';
        return;
    }

    const file = fileInput.files[0];
    const reader = new FileReader();

    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

            // 清空之前的结果
            calculationResults.riverResults = [];
            calculationResults.areaResults = {};

            // 检查是否包含区域信息
            const hasAreaInfo = jsonData[0].length >= 9;

            // 处理每条河流数据
            for (let i = 1; i < jsonData.length; i++) {
                const row = jsonData[i];
                if (!row || row.length < 7) continue;

                const riverName = row[0];
                const concentrations = row.slice(1, 7);
                const riverType = row[7];
                const area = hasAreaInfo && row.length >= 9 ? row[8] : null;

                try {
                    // 计算各个参数的CWQI
                    const parameterCWQIs = {
                        pH: calculateParameterCWQI(concentrations[0], 'pH', riverType),
                        DO: calculateParameterCWQI(concentrations[1], 'DO', riverType),
                        COD: calculateParameterCWQI(concentrations[2], 'COD', riverType),
                        NH3N: calculateParameterCWQI(concentrations[3], 'NH3N', riverType),
                        TP: calculateParameterCWQI(concentrations[4], 'TP', riverType),
                        TN: calculateParameterCWQI(concentrations[5], 'TN', riverType)
                    };

                    // 检查是否所有参数都有有效值
                    const allValid = Object.values(parameterCWQIs).every(value => value !== null);
                    
                    if (allValid) {
                        const totalCWQI = Object.values(parameterCWQIs).reduce((sum, value) => sum + value, 0);
                        
                        // 保存河流结果
                        calculationResults.riverResults.push({
                            riverName,
                            totalCWQI: parseFloat(totalCWQI.toFixed(3)),
                            area
                        });

                        // 累计区域CWQI
                        if (hasAreaInfo && area) {
                            if (!calculationResults.areaResults[area]) {
                                calculationResults.areaResults[area] = 0;
                            }
                            calculationResults.areaResults[area] += totalCWQI;
                        }
                    }
                } catch (error) {
                    console.error(`处理河流 ${riverName} 时出错:`, error);
                }
            }

            // 显示结果
            displayResults(hasAreaInfo && calculationResults.areaResults && Object.keys(calculationResults.areaResults).length > 0);
            
            // 显示导出按钮
            document.getElementById('exportButton').style.display = 'inline';

        } catch (error) {
            output.innerHTML = `<p style="color: red;">处理文件时出错: ${error.message}</p>`;
        }
    };

    reader.onerror = function() {
        output.innerHTML = '<p style="color: red;">文件读取错误，请重试</p>';
    };

    reader.readAsArrayBuffer(file);
}

// 显示计算结果
function displayResults(hasAreaInfo) {
    const output = document.getElementById('output');

    // 创建结果HTML
    let html = '<h3>计算结果</h3>';

    // 创建结果容器
    html += '<div style="display: flex; gap: var(--table-gap, 20px);">';

    // 河流结果表格（保持原始顺序）
    html += '<div style="flex: 1;">';
    html += '<h4>河流水质指数</h4>';
    html += '<table><tr><th>排名</th><th>河流名称</th>' + (hasAreaInfo ? '<th>区域名称</th>' : '') + '<th>水质指数(CWQI)</th></tr>';

    // 计算并添加河流排名
    const sortedRivers = calculationResults.riverResults
        .map(result => ({ ...result, totalCWQI: parseFloat(result.totalCWQI.toFixed(3)) }))
        .sort((a, b) => a.totalCWQI - b.totalCWQI);

    sortedRivers.forEach((result, index) => {
        html += `<tr>
            <td>${index + 1}</td>
            <td style="text-align: left">${result.riverName}</td>` + (hasAreaInfo ? `<td>${result.area}</td>` : '') + `
            <td>${result.totalCWQI}</td>
        </tr>`;
    });

    html += '</table>';
    html += '</div>';

    // 区域结果表格
    if (hasAreaInfo) {
        html += '<div style="flex: 1;">';
        html += '<h4>区域水质指数</h4>';
        const sortedAreas = Object.entries(calculationResults.areaResults)
            .map(([area, totalCWQI]) => ({ area, totalCWQI: parseFloat(totalCWQI.toFixed(3)) }))
            .sort((a, b) => a.totalCWQI - b.totalCWQI);

        html += '<table><tr><th>排名</th><th>区域名称</th><th>CWQI</th></tr>';
        sortedAreas.forEach((result, index) => {
            html += `<tr>
                <td>${index + 1}</td>
                <td>${result.area}</td>
                <td>${result.totalCWQI}</td>
            </tr>`;
        });
        html += '</table>';
        html += '</div>';
    }

    html += '</div>'; // 结束flex容器

    output.innerHTML = html;
}

// 导出结果到Excel
function exportToExcel() {
    // 创建工作簿
    const wb = XLSX.utils.book_new();

    // 准备数据
    const data = [];

    // 判断是否有区域信息
    const hasAreaInfo = calculationResults.riverResults.some(result => result.area);

    // 添加表头
    data.push(["河流名称", "水质指数", "河流水质指数排名"]);
    if (hasAreaInfo) {
        data[0].push("区域名称", "区域水质指数", "区域排名");
    }

    // 计算河流排名
    const sortedRivers = calculationResults.riverResults
        .map(result => result.totalCWQI)
        .sort((a, b) => a - b);

    const riverRankings = {};
    calculationResults.riverResults.forEach(result => {
        const rank = sortedRivers.indexOf(result.totalCWQI) + 1;
        riverRankings[result.riverName] = rank;
    });

    // 如果有区域信息，计算区域排名
    let areaRankings = {};
    if (hasAreaInfo) {
        const sortedAreas = Object.entries(calculationResults.areaResults)
            .map(([area, totalCWQI]) => ({
                area,
                totalCWQI: parseFloat(totalCWQI.toFixed(3))
            }))
            .sort((a, b) => a.totalCWQI - b.totalCWQI);

        sortedAreas.forEach((result, index) => {
            areaRankings[result.area] = index + 1;
        });
    }

    // 添加河流数据
    calculationResults.riverResults.forEach(result => {
        const rowData = [result.riverName, result.totalCWQI, riverRankings[result.riverName]];
        // 只有当有区域信息时才添加区域相关列
        if (hasAreaInfo && result.area) {
            const areaTotal = calculationResults.areaResults[result.area];
            rowData.push(result.area, parseFloat(areaTotal.toFixed(3)), areaRankings[result.area]);
        }
        data.push(rowData);
    });

    // 创建工作表
    const ws = XLSX.utils.aoa_to_sheet(data);

    // 设置合并单元格
    const merges = [];

    // 遍历每一列
    for (let col = 0; col < data[0].length; col++) {
        let startRow = 1;
        let currentValue = data[startRow][col];
        
        for (let i = 2; i < data.length; i++) {
            if (data[i][col] === currentValue) {
                continue;
            }
            
            // 如果发现不同值，检查是否需要合并
            if (i - startRow > 1) {
                merges.push({
                    s: { r: startRow, c: col },
                    e: { r: i - 1, c: col }
                });
                
                // 清除重复值
                for (let j = startRow + 1; j < i; j++) {
                    data[j][col] = '';
                }
            }
            
            startRow = i;
            currentValue = data[i][col];
        }
        
        // 处理最后一组
        if (data.length - startRow > 1) {
            merges.push({
                s: { r: startRow, c: col },
                e: { r: data.length - 1, c: col }
            });
            
            // 清除重复值
            for (let j = startRow + 1; j < data.length; j++) {
                data[j][col] = '';
            }
        }
    }

    ws['!merges'] = merges;

    // 添加工作表
    XLSX.utils.book_append_sheet(wb, ws, "河流与区域CWQI");

    // 导出文件
    XLSX.writeFile(wb, "CWQI计算结果.xlsx");
}

// 图片转Excel功能相关函数
let imageData = null;
let workerPool = []; // Worker线程池
const MAX_WORKERS = 4; // 最大并发线程数

// 初始化Worker线程池
async function initWorkerPool() {
    for (let i = 0; i < MAX_WORKERS; i++) {
        const worker = await Tesseract.createWorker({
            logger: m => console.log(m),
            errorHandler: err => console.error(err)
        });
        await worker.loadLanguage('chi_sim+eng');
        await worker.initialize('chi_sim+eng');
        workerPool.push(worker);
    }
}

// 更新进度条
function updateProgress(progress) {
    const progressBar = document.getElementById('imageConversionProgress');
    const percentage = document.getElementById('progressPercentage');
    const progressValue = Math.round(progress * 100);
    
    progressBar.value = progressValue;
    percentage.textContent = `${progressValue}%`;
}

// 处理图片转换
async function convertImage() {
    const imageInput = document.getElementById('imageInput');
    const output = document.getElementById('imageOutput');
    const progressBar = document.getElementById('imageConversionProgress');
    const progressText = document.getElementById('progressPercentage');

    if (!imageInput.files.length) {
        output.innerHTML = '<p style="color: red;">请先选择一张图片</p>';
        return;
    }

    const file = imageInput.files[0];
    output.innerHTML = '<p>正在处理图片，请稍候...</p>';
    progressBar.value = 0;
    progressText.textContent = '0%';

    try {
        // 初始化Worker线程池
        if (workerPool.length === 0) {
            await initWorkerPool();
        }

        // 图片预处理
        updateProgress(0.1);
        const image = await loadImage(file);
        updateProgress(0.2);
        const processedImage = preprocessImage(image);
        updateProgress(0.3);

        // 分割图片为多个区域
        const regions = splitImageIntoRegions(processedImage);
        updateProgress(0.4);

        // 并行处理各个区域
        const totalRegions = regions.length;
        let completedRegions = 0;
        
        const results = await Promise.all(regions.map(async (region, index) => {
            const result = await recognizeRegion(region, index % workerPool.length);
            completedRegions++;
            updateProgress(0.4 + (completedRegions / totalRegions) * 0.5);
            return result;
        }));

        // 合并识别结果
        const mergedData = mergeRegionResults(results);
        updateProgress(0.9);

        // 保存识别结果
        imageData = mergedData;
        updateProgress(1.0);
        output.innerHTML = '<p style="color: green;">图片处理完成！</p>';
        
        // 显示导出按钮
        document.getElementById('exportImageButton').style.display = 'inline';

    } catch (error) {
        output.innerHTML = `<p style="color: red;">图片处理失败: ${error.message}</p>`;
        updateProgress(0);
    }
}

// 加载图片
function loadImage(file) {
    return new Promise((resolve, reject) => {
        const img = new Image();
        img.onload = () => resolve(img);
        img.onerror = reject;
        img.src = URL.createObjectURL(file);
    });
}

// 图片预处理
function preprocessImage(img) {
    const canvas = document.createElement('canvas');
    const ctx = canvas.getContext('2d');
    
    // 设置画布大小
    const maxSize = 1200; // 降低最大尺寸以加快处理速度
    const scale = Math.min(maxSize / img.width, maxSize / img.height);
    canvas.width = img.width * scale;
    canvas.height = img.height * scale;

    // 绘制图片
    ctx.drawImage(img, 0, 0, canvas.width, canvas.height);

    // 获取图像数据
    const imageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
    const data = imageData.data;

    // 快速二值化处理
    const threshold = 128; // 使用固定阈值加快处理速度
    for (let i = 0; i < data.length; i += 4) {
        const avg = (data[i] + data[i + 1] + data[i + 2]) / 3;
        const value = avg > threshold ? 255 : 0;
        data[i] = value; // red
        data[i + 1] = value; // green
        data[i + 2] = value; // blue
    }

    // 快速去噪处理
    const kernelSize = 3;
    const halfKernel = Math.floor(kernelSize / 2);
    const tempData = new Uint8ClampedArray(data.length);
    
    // 使用快速去噪算法
    for (let y = halfKernel; y < canvas.height - halfKernel; y += kernelSize) {
        for (let x = halfKernel; x < canvas.width - halfKernel; x += kernelSize) {
            const index = (y * canvas.width + x) * 4;
            let count = 0;
            
            // 只检查4个角点
            const offsets = [
                [-halfKernel, -halfKernel],
                [-halfKernel, halfKernel],
                [halfKernel, -halfKernel],
                [halfKernel, halfKernel]
            ];
            
            offsets.forEach(([ky, kx]) => {
                const kIndex = ((y + ky) * canvas.width + (x + kx)) * 4;
                if (data[kIndex] === 0) count++;
            });
            
            const newValue = count > 2 ? 0 : 255;
            // 填充整个kernel区域
            for (let ky = -halfKernel; ky <= halfKernel; ky++) {
                for (let kx = -halfKernel; kx <= halfKernel; kx++) {
                    const fillIndex = ((y + ky) * canvas.width + (x + kx)) * 4;
                    tempData[fillIndex] = newValue;
                    tempData[fillIndex + 1] = newValue;
                    tempData[fillIndex + 2] = newValue;
                    tempData[fillIndex + 3] = 255;
                }
            }
        }
    }

    // 应用去噪结果
    for (let i = 0; i < data.length; i++) {
        data[i] = tempData[i];
    }

    ctx.putImageData(imageData, 0, 0);
    return canvas;
}
