// ==UserScript==
// @name         正方教务系统批量打印周次课表
// @version      1.1
// @description  在"学生课表查询（按周次）"(/kbcx/xskbcxZccx_cxXskbcxIndex.html)页面增加批量打印所有周次的功能
// @author       短臂猿-Short_Arm_Ape
// @homepage     https://github.com/Short-Arm-Ape/ZFsoft-TIISP-Print-Student-Schedule-by-Week
// @match        *://*/jwglxt/kbcx/xskbcxZccx_cxXskbcxIndex.html*
// @icon         https://cloud.zfsoft.com:6143/jwglxt/logo/favicon.ico
// @grant        none
// ==/UserScript==

(function () {
    'use strict';

    initBatchPrintOnTimetablePage();

    // 课表页面初始化：添加批量打印按钮，绑定事件
    function initBatchPrintOnTimetablePage() {
        if (document.readyState === 'loading') {
            document.addEventListener('DOMContentLoaded', setupBatchPrintUI);
        } else {
            setupBatchPrintUI();
        }
        window.addEventListener('load', setupBatchPrintUI);
    }

    function setupBatchPrintUI() {
        if (document.getElementById('batchPrintBtn')) return;

        const zsLi = document.getElementById('zsli');
        if (!zsLi) {
            setTimeout(setupBatchPrintUI, 500);
            return;
        }

        const batchBtn = document.createElement('button');
        batchBtn.type = 'button';  // 禁止提交表单行为，避免被导向至系统维护页
        batchBtn.id = 'batchPrintBtn';
        batchBtn.className = 'btn btn-primary btn-sm';
        batchBtn.style.marginLeft = '10px';
        batchBtn.style.verticalAlign = 'middle';
        batchBtn.innerHTML = '<span class="glyphicon glyphicon-print"></span> 批量打印所有周次';
        batchBtn.onclick = confirmRun;

        zsLi.appendChild(batchBtn);
    }

    // 确认开始
    function confirmRun() {
        let totalWeeks = $('#zs option').length;
        let estimatedDuration = totalWeeks * 3
        if (confirm(`将开始批量打印 ${totalWeeks} 个周次的课表，预计最长耗时 ${estimatedDuration} 秒，请确保网络通畅。\n点击确定后，脚本会自动遍历周次并收集数据，最后打开打印预览。\n在此过程中，请勿进行任何操作，包括调整页面缩放和窗口大小等。`)) {
            batchPrintAllWeeks();
        } else {
            return;
        }
    }

    async function batchPrintAllWeeks() {
        const btn = document.getElementById('batchPrintBtn');
        if (!btn) return;

        if (btn.disabled) return;
        btn.disabled = true;
        btn.innerHTML = '正在生成打印预览...';

        const $zs = $('#zs');
        const originalWeekVal = $zs.val();

        const weekOptions = $('#zs option').map(function () {
            return {
                value: $(this).val(),
                text: $(this).text().trim()
            };
        }).get();

        if (weekOptions.length === 0) {
            alert('未找到周次信息');
            btn.disabled = false;
            btn.innerHTML = '<span class="glyphicon glyphicon-print"></span> 批量打印所有周次';
            return;
        }

        const $printContainer = $('<div id="batchPrintContainer" style="display:none;"></div>');
        $('body').append($printContainer);

        let lastTableHtml = $('#myTab').html();

        for (let i = 0; i < weekOptions.length; i++) {
            const week = weekOptions[i];
            btn.innerHTML = `正在获取第 ${i + 1}/${weekOptions.length} 周...`;

            $zs.val(week.value).trigger('chosen:updated');

            if (typeof window.searchResult === 'function') {
                window.searchResult();
            } else {
                $('#to-search').first().click();
            }

            try {
                await waitForTableUpdate(lastTableHtml);
            } catch (e) {
                console.warn(`第${week.text}周加载超时，仍尝试采集当前表格`);
            }
            lastTableHtml = $('#myTab').html();

            const $tableClone = $('#myTab').clone(true);

            const $weekDiv = $('<div class="print-week-block"></div>');
            $weekDiv.append(`<h3 style="text-align:center; margin:20px 0;">第${week.text}周</h3>`);
            $weekDiv.append($tableClone);

            if (i > 0) {
                $weekDiv.css('page-break-before', 'always');
            }

            $printContainer.append($weekDiv);
        }

        $zs.val(originalWeekVal).trigger('chosen:updated');
        if (typeof window.searchResult === 'function') {
            window.searchResult();
        } else {
            $('#to-search').first().click();
        }

        btn.disabled = false;
        btn.innerHTML = '<span class="glyphicon glyphicon-print"></span> 批量打印所有周次';

        printCollectedWeeks($printContainer);
        $printContainer.empty();
    }

    function waitForTableUpdate(oldHtml) {
        return new Promise((resolve, reject) => {
            const maxChecks = 15;
            let checks = 0;
            const interval = setInterval(() => {
                const currentHtml = $('#myTab').html();
                if (currentHtml !== oldHtml) {
                    clearInterval(interval);
                    resolve();
                } else if (++checks >= maxChecks) {
                    clearInterval(interval);
                    reject('更新超时');
                }
            }, 200);// 即超过200*15=3000ms后判定超时
        });
    }

    function printCollectedWeeks($container) {
        const $iframe = $('<iframe style="position:absolute; width:0; height:0; border:0;"></iframe>');
        $('body').append($iframe);
        const iframe = $iframe[0];
        const iframeDoc = iframe.contentWindow.document;

        iframeDoc.open();
        // 获取学年和学期显示文本
        const xnmText = $('#xnm option:selected').text() || '未知';
        const xqmText = $('#xqm option:selected').text() || '未知';
        const pageTitle = xnmText + '学年' + '第' + xqmText + '学期课表';
        iframeDoc.write('<!DOCTYPE html><html><head><title>' + pageTitle + '</title>');

        copyStylesToIframe(iframeDoc);

        iframeDoc.write(`
            <style>
                /* 修复原页面样式导致打印高度限制的问题 */
                html, body {
                    height: auto !important;
                    overflow: visible !important;
                    width: auto !important;
                }
                table {
                    page-break-inside: auto !important;
                }
                @media print {
                    /* 全局强制保留背景色 */
                    * {
                        -webkit-print-color-adjust: exact !important;
                        print-color-adjust: exact !important;
                    }
                    body { background: white; }
                    .print-week-block { page-break-after: always; }
                    /* 隐藏固定列干扰元素 */
                    .fsttd, .fixleft, .fixtop { display: none !important; }
                    /* 表格边框优化 */
                    .weui-table { border-collapse: collapse; width: 100%; }
                    .weui-table td, .weui-table th { border: 1px solid #000; }
                }
            </style>
        `);
        iframeDoc.write('</head><body>');
        iframeDoc.write($container.html());
        iframeDoc.write('</body></html>');
        iframeDoc.close();
        $iframe.on('load', function () {
            iframe.contentWindow.focus();
            iframe.contentWindow.print();
            setTimeout(() => $iframe.remove(), 1000);
        });
    }

    function copyStylesToIframe(iframeDoc) {
        const links = document.querySelectorAll('link[rel="stylesheet"]');
        links.forEach(link => {
            try {
                const href = link.href;
                if (href) {
                    iframeDoc.write(`<link rel="stylesheet" type="text/css" href="${href}">`);
                }
            } catch (e) {
                console.warn('复制样式失败', e);
            }
        });

        const styles = document.querySelectorAll('style');
        styles.forEach(style => {
            iframeDoc.write(`<style>${style.innerHTML}</style>`);
        });
    }

    if (typeof $ === 'undefined') {
        console.error('此脚本依赖jQuery，但页面未加载。');
        return;
    }
})();
