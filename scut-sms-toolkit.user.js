// ==UserScript==
// @name         SCUT 学工系统-综测数据工具箱(ZIP导出+加权均分+综测总分Excel)
// @namespace    https://github.com/Breeze1733/scut-sms-toolkit
// @version      2.4.0
// @description  SCUT 学工系统综测辅助工具：全班CSV打包ZIP导出、加权平均分Excel导出(双Sheet)、读取学生申请分数计算综测总成绩(X/C/S)导出Excel
// @author       Breeze1733
// @license      MIT
// @match        https://sms.scut.edu.cn/*
// @match        https://sms-443.webvpn.scut.edu.cn/*
// @icon         https://sms.scut.edu.cn/favicon.ico
// @require      https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js
// @require      https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js
// @homepageURL  https://github.com/Breeze1733/scut-sms-toolkit
// @supportURL   https://github.com/Breeze1733/scut-sms-toolkit/issues
// @updateURL    https://raw.githubusercontent.com/Breeze1733/scut-sms-toolkit/main/scut-sms-toolkit.user.js
// @downloadURL  https://raw.githubusercontent.com/Breeze1733/scut-sms-toolkit/main/scut-sms-toolkit.user.js
// @grant        none
// ==/UserScript==

(function() {
    'use strict';

    // 初始化操作按钮栏（保持简洁的两个按钮）
    function init() {
        const table = document.querySelector('table');
        if (!table || document.getElementById('scut-tools-container')) return;

        const container = document.createElement('div');
        container.id = 'scut-tools-container';
        container.style.cssText = `
            position: fixed;
            top: 15px;
            right: 20px;
            z-index: 999999;
            display: flex;
            gap: 10px;
        `;

        // 功能一按钮：ZIP导出
        const zipBtn = document.createElement('button');
        zipBtn.id = 'scut-export-zip-btn';
        zipBtn.innerText = '📦 打包全班CSV (ZIP)';
        zipBtn.style.cssText = getButtonStyle('#28a745');
        zipBtn.onclick = startBatchExportZIP;

        // 功能二按钮：加权均分Excel导出（一次性计算双模式输出两个Sheet）
        const gpaBtn = document.createElement('button');
        gpaBtn.id = 'scut-calc-gpa-btn';
        gpaBtn.innerText = '📊 计算加权平均分 (Excel)';
        gpaBtn.style.cssText = getButtonStyle('#007bff');
        gpaBtn.onclick = startCalcWeightedGPA;

        // 功能三按钮：读取学生申请分数计算综测总成绩并导出 Excel
        const evalBtn = document.createElement('button');
        evalBtn.id = 'scut-calc-eval-btn';
        evalBtn.innerText = '📈 计算综测总成绩 (Excel)';
        evalBtn.style.cssText = getButtonStyle('#fd7e14');
        evalBtn.onclick = startCalcComprehensiveEval;

        container.appendChild(zipBtn);
        container.appendChild(gpaBtn);
        container.appendChild(evalBtn);
        document.body.appendChild(container);
    }

    function getButtonStyle(bgColor) {
        return `
            padding: 9px 16px;
            background-color: ${bgColor};
            color: #fff;
            border: none;
            border-radius: 4px;
            font-size: 13px;
            font-weight: bold;
            cursor: pointer;
            box-shadow: 0 2px 8px rgba(0,0,0,0.2);
            transition: opacity 0.2s;
            outline: none;
        `;
    }

    // 提取成绩数字（支持纯数字及 "优秀（95.0）"、"良好(85.0)" 格式）
    function extractScore(text) {
        if (!text) return null;
        text = text.trim();
        const bracketMatch = text.match(/[（(]\s*([0-9]+(?:\.[0-9]+)?)\s*[）)]/);
        if (bracketMatch) {
            return parseFloat(bracketMatch[1]);
        }
        const numMatch = text.match(/^[0-9]+(?:\.[0-9]+)?$/);
        if (numMatch) {
            return parseFloat(numMatch[0]);
        }
        const anyNumMatch = text.match(/[0-9]+(?:\.[0-9]+)?/);
        if (anyNumMatch) {
            return parseFloat(anyNumMatch[0]);
        }
        return null;
    }

    // 获取列表有效数据行及详情页链接
    function getTargetListRows() {
        const trs = Array.from(document.querySelectorAll('table tbody tr, table tr')).filter(tr => {
            const links = Array.from(tr.querySelectorAll('a'));
            return links.some(a => a.innerText.includes('查看') || a.getAttribute('onclick')?.includes('detail') || a.href.includes('detail'));
        });

        return trs.map((tr, index) => {
            const cells = Array.from(tr.querySelectorAll('td')).map(td => td.innerText.trim());
            const studentId = cells[1] || `学号_${index + 1}`;
            const studentName = cells[2] || `姓名_${index + 1}`;

            const viewLink = Array.from(tr.querySelectorAll('a')).find(a =>
                a.innerText.includes('查看') || a.getAttribute('onclick')?.includes('detail') || a.href.includes('detail')
            );

            let url = null;
            if (viewLink) {
                const href = viewLink.getAttribute('href');
                if (href && !href.startsWith('javascript:') && href !== '#') {
                    url = new URL(href, document.baseURI).href;
                } else {
                    const match = (viewLink.getAttribute('onclick') || '').match(/['"]([^'"]*detail[^'"]*)['"]/);
                    if (match) url = new URL(match[1], document.baseURI).href;
                }
            }
            return { studentId, studentName, url };
        }).filter(item => item.url !== null);
    }

    // ================= 功能三：读取学生申请分数计算综测总成绩并导出 Excel =================
    async function startCalcComprehensiveEval() {
        if (typeof XLSX === 'undefined') {
            alert('SheetJS 库未加载完成，请稍后刷新重试！');
            return;
        }

        const items = getTargetListRows();
        if (items.length === 0) {
            alert('未检测到包含“查看”操作的数据行，请确认处于学工系统学生列表页面！');
            return;
        }

        const btn = document.getElementById('scut-calc-eval-btn');
        btn.disabled = true;
        btn.style.opacity = '0.6';

        // 表头字段
        const excelRows = [
            [
                '学号',
                '姓名',
                '学业成绩积分(X)',
                '操行评定基本分',
                '操行评定加分',
                '操行评定积分(C)',
                '德育积分',
                '智育积分',
                '体育积分',
                '美育积分',
                '劳育积分',
                '综合素养评定积分(S)',
                '综测总成绩'
            ]
        ];

        for (let i = 0; i < items.length; i++) {
            const { studentId, studentName, url } = items[i];
            btn.innerText = `⏳ 计算中 (${i + 1}/${items.length})：${studentName}`;

            try {
                const response = await fetch(url, {
                    method: 'GET',
                    credentials: 'include',
                    headers: { 'X-Requested-With': 'XMLHttpRequest' }
                });
                if (!response.ok) throw new Error(`HTTP ${response.status}`);

                const htmlText = await response.text();
                const parser = new DOMParser();
                const doc = parser.parseFromString(htmlText, 'text/html');

                // 提取学号与姓名
                const pageText = doc.body.innerText;
                const idMatch = pageText.match(/学号[：:]\s*([0-9A-Za-z]+)/);
                const nameMatch = pageText.match(/姓名[：:]\s*([^\s]+)/);
                const actualId = idMatch ? idMatch[1].trim() : studentId;
                const actualName = nameMatch ? nameMatch[1].trim() : studentName;

                const tables = Array.from(doc.querySelectorAll('table'));

                // 1. 计算学业成绩积分 X（复用功能二：必修课＋选修课加权平均分，排除通选课）
                let gradeTable = null;
                let scoreCol = -1;
                let creditCol = -1;
                let typeCol = -1;

                for (const t of tables) {
                    const trs = Array.from(t.querySelectorAll('tr'));
                    for (const r of trs) {
                        const texts = Array.from(r.querySelectorAll('th, td')).map(cell => cell.innerText.trim());
                        const sIdx = texts.findIndex(txt => txt === '成绩');
                        const cIdx = texts.findIndex(txt => txt === '学分');
                        const tIdx = texts.findIndex(txt => txt.includes('课程类型') || txt === '类型' || txt.includes('课程性质'));

                        if (sIdx !== -1 && cIdx !== -1 && texts.some(txt => txt.includes('课程名称'))) {
                            gradeTable = t;
                            scoreCol = sIdx;
                            creditCol = cIdx;
                            typeCol = tIdx;
                            break;
                        }
                    }
                    if (gradeTable) break;
                }

                let weightBoth = 0;
                let creditBoth = 0;
                if (gradeTable) {
                    const rows = Array.from(gradeTable.querySelectorAll('tr'));
                    for (const r of rows) {
                        const cells = Array.from(r.querySelectorAll('td'));
                        if (cells.length <= Math.max(scoreCol, creditCol)) continue;

                        let isBoth = true;
                        if (typeCol !== -1 && cells.length > typeCol) {
                            const courseType = cells[typeCol].innerText.trim();
                            const hasRequired = courseType.includes('必修') && !courseType.includes('通选');
                            const hasElective = courseType.includes('选修') && !courseType.includes('通选');
                            isBoth = hasRequired || hasElective;
                        }

                        const scoreVal = extractScore(cells[scoreCol].innerText);
                        const creditVal = parseFloat(cells[creditCol].innerText.trim());

                        if (scoreVal !== null && !isNaN(scoreVal) && !isNaN(creditVal) && creditVal > 0) {
                            if (isBoth) {
                                weightBoth += scoreVal * creditVal;
                                creditBoth += creditVal;
                            }
                        }
                    }
                }
                const X = creditBoth > 0 ? Number((weightBoth / creditBoth).toFixed(2)) : 0.00;

                // 2. 解析各评定模块的学生申请分数 / 学生评分
                const scores = {
                    conduct_base: 0,
                    conduct_add: 0,
                    moral: 0,
                    intellectual: 0,
                    sports: 0,
                    arts: 0,
                    labor: 0
                };

                for (const table of tables) {
                    if (table === gradeTable) continue;

                    // 识别当前模块标题
                    const container = table.closest('.panel, .panel-default, .card, [class*="panel"], div');
                    const heading = container?.querySelector('h1, h2, h3, h4, h5, h6, .panel-heading, .panel-title');
                    const divHeading = table.closest('div')?.querySelector('h4, h5, .panel-heading, .panel-title');
                    const prevText = table.previousElementSibling ? table.previousElementSibling.innerText : '';
                    const title = ((heading ? heading.innerText : '') + ' ' + (divHeading ? divHeading.innerText : '') + ' ' + prevText).trim();

                    let cat = null;
                    if (title.includes('0201') || title.includes('操行评定基本分') || title.includes('操行基本分')) {
                        cat = 'conduct_base';
                    } else if (title.includes('0202') || title.includes('操行评定加分') || title.includes('操行加分')) {
                        cat = 'conduct_add';
                    } else if (title.includes('0301') || title.includes('德育')) {
                        cat = 'moral';
                    } else if ((title.includes('0302') || title.includes('智育')) && !title.includes('学业')) {
                        cat = 'intellectual';
                    } else if (title.includes('0303') || title.includes('体育')) {
                        cat = 'sports';
                    } else if (title.includes('0304') || title.includes('美育')) {
                        cat = 'arts';
                    } else if (title.includes('0305') || title.includes('劳育')) {
                        cat = 'labor';
                    }

                    if (cat && scores.hasOwnProperty(cat)) {
                        const trs = Array.from(table.querySelectorAll('tr'));
                        let scoreColIdx = -1;

                        for (const tr of trs) {
                            const cells = Array.from(tr.querySelectorAll('th, td'));
                            const texts = cells.map(c => c.innerText.trim());

                            if (texts.includes('序号') || texts.some(t => t.includes('学生申请分数') || t.includes('学生评分'))) {
                                scoreColIdx = texts.findIndex(t => t.includes('学生申请分数') || t.includes('学生评分'));
                                continue;
                            }

                            if (scoreColIdx !== -1 && cells.length > scoreColIdx) {
                                const val = extractScore(cells[scoreColIdx].innerText);
                                if (val !== null && !isNaN(val)) {
                                    scores[cat] += val;
                                }
                            }
                        }
                    }
                }

                // 3. 计算操行评定积分 (C) = 操行评定基本分 + 操行评定加分
                const baseScore = Number(scores.conduct_base.toFixed(2));
                const addScore = Number(scores.conduct_add.toFixed(2));
                const C = Number((baseScore + addScore).toFixed(2));

                // 4. 计算综合素养评定积分 (S) = 德育积分 + 3 * 智育积分 + 体育积分 + 美育积分 + 劳育积分
                const moralScore = Number(scores.moral.toFixed(2));
                const intelScore = Number(scores.intellectual.toFixed(2));
                const sportsScore = Number(scores.sports.toFixed(2));
                const artsScore = Number(scores.arts.toFixed(2));
                const laborScore = Number(scores.labor.toFixed(2));
                const S = Number((moralScore + 3 * intelScore + sportsScore + artsScore + laborScore).toFixed(2));

                // 5. 计算综测成绩 = 0.7X + 0.1C + 0.2S
                const totalScore = Number((0.7 * X + 0.1 * C + 0.2 * S).toFixed(2));

                excelRows.push([
                    actualId,
                    actualName,
                    X,
                    baseScore,
                    addScore,
                    C,
                    moralScore,
                    intelScore,
                    sportsScore,
                    artsScore,
                    laborScore,
                    S,
                    totalScore
                ]);

                await new Promise(resolve => setTimeout(resolve, 300));
            } catch (err) {
                console.error(`[-] 获取 ${studentName} 综测数据失败:`, err);
                excelRows.push([
                    studentId,
                    studentName,
                    '计算失败',
                    '-', '-', '-', '-', '-', '-', '-', '-', '-', '-'
                ]);
            }
        }

        btn.innerText = '📑 正在导出 Excel...';

        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.aoa_to_sheet(excelRows);
        ws['!cols'] = [
            { wch: 16 }, // 学号
            { wch: 12 }, // 姓名
            { wch: 16 }, // 学业成绩积分(X)
            { wch: 15 }, // 操行评定基本分
            { wch: 15 }, // 操行评定加分
            { wch: 16 }, // 操行评定积分(C)
            { wch: 12 }, // 德育积分
            { wch: 12 }, // 智育积分
            { wch: 12 }, // 体育积分
            { wch: 12 }, // 美育积分
            { wch: 12 }, // 劳育积分
            { wch: 20 }, // 综合素养评定积分(S)
            { wch: 14 }  // 综测总成绩
        ];
        XLSX.utils.book_append_sheet(wb, ws, "综测成绩汇总");

        const today = new Date().toISOString().slice(0, 10);
        XLSX.writeFile(wb, `全班综测成绩汇总_${today}.xlsx`);

        btn.innerText = '✅ 综测成绩导出完成';
        btn.style.opacity = '1';
        setTimeout(() => {
            btn.disabled = false;
            btn.innerText = '📈 计算综测总成绩 (Excel)';
        }, 3000);
    }

    // ================= 功能二：遍历计算加权平均分并导出 Excel（双 Sheet） =================
    async function startCalcWeightedGPA() {
        if (typeof XLSX === 'undefined') {
            alert('SheetJS 库未加载完成，请稍后刷新重试！');
            return;
        }

        const items = getTargetListRows();
        if (items.length === 0) {
            alert('未检测到包含“查看”操作的数据行，请确认处于学工系统学生列表页面！');
            return;
        }

        const btn = document.getElementById('scut-calc-gpa-btn');
        btn.disabled = true;
        btn.style.opacity = '0.6';

        // 分别准备两个 Sheet 的数据行：表头三列均为【姓名、加权平均分、学分总数】
        const excelRowsBoth = [
            ['姓名', '加权平均分', '学分总数']
        ];
        const excelRowsRequiredOnly = [
            ['姓名', '加权平均分', '学分总数']
        ];

        for (let i = 0; i < items.length; i++) {
            const { studentName, url } = items[i];
            btn.innerText = `⏳ 计算中 (${i + 1}/${items.length})：${studentName}`;

            try {
                const response = await fetch(url, {
                    method: 'GET',
                    credentials: 'include',
                    headers: { 'X-Requested-With': 'XMLHttpRequest' }
                });
                if (!response.ok) throw new Error(`HTTP ${response.status}`);

                const htmlText = await response.text();
                const parser = new DOMParser();
                const doc = parser.parseFromString(htmlText, 'text/html');

                // 从详情页提取真实姓名
                const nameMatch = doc.body.innerText.match(/姓名[：:]\s*([^\s]+)/);
                const actualName = nameMatch ? nameMatch[1].trim() : studentName;

                // 寻找成绩表格并定位“成绩”、“学分”、“课程类型”所在列
                const tables = Array.from(doc.querySelectorAll('table'));
                let gradeTable = null;
                let scoreCol = -1;
                let creditCol = -1;
                let typeCol = -1;

                for (const t of tables) {
                    const trs = Array.from(t.querySelectorAll('tr'));
                    for (const r of trs) {
                        const texts = Array.from(r.querySelectorAll('th, td')).map(cell => cell.innerText.trim());
                        const sIdx = texts.findIndex(txt => txt === '成绩');
                        const cIdx = texts.findIndex(txt => txt === '学分');
                        const tIdx = texts.findIndex(txt => txt.includes('课程类型') || txt === '类型' || txt.includes('课程性质'));

                        if (sIdx !== -1 && cIdx !== -1 && texts.some(txt => txt.includes('课程名称'))) {
                            gradeTable = t;
                            scoreCol = sIdx;
                            creditCol = cIdx;
                            typeCol = tIdx;
                            break;
                        }
                    }
                    if (gradeTable) break;
                }

                // 范围1：必修课 + 选修课（排除通选课）
                let weightBoth = 0;
                let creditBoth = 0;

                // 范围2：仅必修课（排除选修及通选课）
                let weightRequired = 0;
                let creditRequired = 0;

                if (gradeTable) {
                    const rows = Array.from(gradeTable.querySelectorAll('tr'));
                    for (const r of rows) {
                        const cells = Array.from(r.querySelectorAll('td'));
                        if (cells.length <= Math.max(scoreCol, creditCol)) continue;

                        let isRequired = true;
                        let isBoth = true;

                        if (typeCol !== -1 && cells.length > typeCol) {
                            const courseType = cells[typeCol].innerText.trim();
                            const hasRequired = courseType.includes('必修') && !courseType.includes('通选');
                            const hasElective = courseType.includes('选修') && !courseType.includes('通选');

                            isRequired = hasRequired;
                            isBoth = hasRequired || hasElective;
                        }

                        const scoreVal = extractScore(cells[scoreCol].innerText);
                        const creditVal = parseFloat(cells[creditCol].innerText.trim());

                        if (scoreVal !== null && !isNaN(scoreVal) && !isNaN(creditVal) && creditVal > 0) {
                            if (isBoth) {
                                weightBoth += scoreVal * creditVal;
                                creditBoth += creditVal;
                            }
                            if (isRequired) {
                                weightRequired += scoreVal * creditVal;
                                creditRequired += creditVal;
                            }
                        }
                    }
                }

                // 记录【必修+选修】成绩
                const avgBoth = creditBoth > 0 ? Number((weightBoth / creditBoth).toFixed(2)) : 0.00;
                const roundedCreditBoth = Number(creditBoth.toFixed(2));
                excelRowsBoth.push([actualName, avgBoth, roundedCreditBoth]);

                // 记录【仅必修课】成绩
                const avgRequired = creditRequired > 0 ? Number((weightRequired / creditRequired).toFixed(2)) : 0.00;
                const roundedCreditRequired = Number(creditRequired.toFixed(2));
                excelRowsRequiredOnly.push([actualName, avgRequired, roundedCreditRequired]);

                // 延时 300ms 避免请求过频
                await new Promise(resolve => setTimeout(resolve, 300));
            } catch (err) {
                console.error(`[-] 获取 ${studentName} 成绩失败:`, err);
                excelRowsBoth.push([studentName, '计算失败', '-']);
                excelRowsRequiredOnly.push([studentName, '计算失败', '-']);
            }
        }

        btn.innerText = '📑 正在导出 Excel...';

        const wb = XLSX.utils.book_new();

        // Sheet 1: 必修课＋选修课
        const wsBoth = XLSX.utils.aoa_to_sheet(excelRowsBoth);
        wsBoth['!cols'] = [{ wch: 15 }, { wch: 15 }, { wch: 15 }];
        XLSX.utils.book_append_sheet(wb, wsBoth, "必修课＋选修课");

        // Sheet 2: 仅必修课
        const wsRequired = XLSX.utils.aoa_to_sheet(excelRowsRequiredOnly);
        wsRequired['!cols'] = [{ wch: 15 }, { wch: 15 }, { wch: 15 }];
        XLSX.utils.book_append_sheet(wb, wsRequired, "仅必修课");

        const today = new Date().toISOString().slice(0, 10);
        XLSX.writeFile(wb, `全班加权平均分汇总_${today}.xlsx`);

        btn.innerText = '✅ Excel 导出完成(双Sheet)';
        btn.style.opacity = '1';
        setTimeout(() => {
            btn.disabled = false;
            btn.innerText = '📊 计算加权平均分 (Excel)';
        }, 3000);
    }

    // ================= 功能一：打包全班 CSV 为 ZIP =================
    async function startBatchExportZIP() {
        if (typeof JSZip === 'undefined') {
            alert('JSZip 依赖库加载失败，请刷新重试！');
            return;
        }

        const items = getTargetListRows();
        if (items.length === 0) {
            alert('未检测到包含“查看”操作的数据行，请确认处于学工系统学生列表页面！');
            return;
        }

        const btn = document.getElementById('scut-export-zip-btn');
        btn.disabled = true;
        btn.style.opacity = '0.6';

        const zip = new JSZip();
        const folderName = `综测导出_${new Date().toISOString().slice(0, 10)}`;
        const folder = zip.folder(folderName);
        let successCount = 0;

        for (let i = 0; i < items.length; i++) {
            const { studentId, studentName, url } = items[i];
            btn.innerText = `⏳ 抓取中 (${i + 1}/${items.length})：${studentName}`;

            try {
                const response = await fetch(url, {
                    method: 'GET',
                    credentials: 'include',
                    headers: { 'X-Requested-With': 'XMLHttpRequest' }
                });
                if (!response.ok) throw new Error(`HTTP ${response.status}`);

                const htmlText = await response.text();
                const parser = new DOMParser();
                const doc = parser.parseFromString(htmlText, 'text/html');
                const csvSections = [];

                const pageText = doc.body.innerText;
                const idMatch = pageText.match(/学号[：:]\s*([0-9A-Za-z]+)/);
                const nameMatch = pageText.match(/姓名[：:]\s*([^\s]+)/);
                const actualId = idMatch ? idMatch[1].trim() : studentId;
                const actualName = nameMatch ? nameMatch[1].trim() : studentName;

                csvSections.push(`"学号","${actualId}","姓名","${actualName}"`);
                csvSections.push('');

                const tables = doc.querySelectorAll('table');
                tables.forEach((table, tIdx) => {
                    const title = table.closest('div')?.querySelector('h4, h5, .panel-heading')?.innerText.trim() || `表格_${tIdx + 1}`;
                    csvSections.push(`"【${title.replace(/"/g, '""')}】"`);
                    Array.from(table.querySelectorAll('tr')).forEach(tr => {
                        const cells = Array.from(tr.querySelectorAll('th, td')).map(td => `"${td.innerText.trim().replace(/\s+/g, ' ').replace(/"/g, '""')}"`);
                        if (cells.length > 0) csvSections.push(cells.join(','));
                    });
                    csvSections.push('');
                });

                folder.file(`${actualId}_${actualName}.csv`, '\uFEFF' + csvSections.join('\r\n'));
                successCount++;
                await new Promise(resolve => setTimeout(resolve, 300));
            } catch (err) {
                console.error(`[-] 抓取 ${studentName} CSV 失败:`, err);
            }
        }

        btn.innerText = '🗜️ 正在压缩打包...';
        const blob = await zip.generateAsync({ type: 'blob' });
        const downloadUrl = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = downloadUrl;
        a.download = `${folderName}_共${successCount}人.zip`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(downloadUrl);

        btn.innerText = '✅ 打包下载完成';
        btn.style.opacity = '1';
        setTimeout(() => {
            btn.disabled = false;
            btn.innerText = '📦 打包全班CSV (ZIP)';
        }, 3000);
    }

    // 页面加载完成后初始化，并监听 DOM 变动确保单页切换或异步加载表格时也能显示
    setTimeout(init, 1000);

    const observer = new MutationObserver(() => {
        if (!document.getElementById('scut-tools-container') && document.querySelector('table')) {
            init();
        }
    });
    if (document.body) {
        observer.observe(document.body, { childList: true, subtree: true });
    }
})();

