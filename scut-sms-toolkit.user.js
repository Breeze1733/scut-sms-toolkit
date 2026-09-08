// ==UserScript==
// @name         SCUT 学工系统-综测数据工具箱(ZIP导出+加权均分Excel)
// @namespace    https://github.com/Breeze1733/scut-sms-toolkit
// @version      2.3.1
// @description  SCUT 学工系统综测辅助工具：支持全班CSV打包ZIP导出、自动解析成绩计算加权平均分并导出Excel（单次计算自动生成“必修课＋选修课”与“仅必修课”两个Sheet）
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

        container.appendChild(zipBtn);
        container.appendChild(gpaBtn);
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

