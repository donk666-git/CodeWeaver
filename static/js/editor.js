// static/js/editor.js - code editor logic

import { generateHighlightHtml, normalizeIndentationText } from './highlight.js';
import { showStatus } from './state.js';

export function applyIndentationNormalization() {
    const code = $('#codeSource').val();
    if (!code) return showStatus("⚠️ No code", "error");
    const lang = $('#langSelect').val();
    const normalized = normalizeIndentationText(code, lang);
    $('#codeSource').val(normalized);
    showStatus("✅ Formatted");
}

// Insert highlighted code
export async function insertHighlight() {
    const code = $('#codeSource').val();
    const lang = $('#langSelect').val();
    const theme = $('#themeSelect').val();

    if (!code) return showStatus("❌ No code", "error");
    
    try {
        let newListingNumber = 1;
        
        await Word.run(async (ctx) => {
            const paragraphs = ctx.document.body.paragraphs;
            ctx.load(paragraphs, 'text');
            await ctx.sync();
            
            let maxNumberInDoc = 0;
            for (let i = 0; i < paragraphs.items.length; i++) {
                const paragraph = paragraphs.items[i];
                const match = paragraph.text.match(/Listing\s+(\d+):/);
                if (match) {
                    const number = parseInt(match[1]);
                    if (number > maxNumberInDoc) {
                        maxNumberInDoc = number;
                    }
                }
            }
            
            newListingNumber = maxNumberInDoc + 1;
            
            const selection = ctx.document.getSelection();
            const html = generateHighlightHtml(code, lang, theme, newListingNumber);
            selection.insertHtml(html, 'Replace');
            
            await ctx.sync();
        });
        
        showStatus(`✅ Inserted (Listing ${newListingNumber})`);
    } catch (e) {
        console.error(e);
        showStatus("❌ Insert failed: " + e.message, "error");
    }
}

// 修复：智能吸取模式
export async function getFromSelection() {
    try {
        await Word.run(async (ctx) => {
            // 1. 获取当前选区
            let range = ctx.document.getSelection();
            
            // 逻辑:检查光标是否在表格内
            const parentTable = range.parentTableOrNullObject;
            ctx.load(parentTable);
            await ctx.sync();

            // 如果在表格里，强制把“选区”扩展为“整个表格”
            // 这样哪怕只点了一下代码块，也能吸取全部代码！
            if (!parentTable.isNullObject) {
                range = parentTable.getRange();
            }
            
            // 2. 尝试 HTML 解析 (结构化数据)
            const htmlResult = range.getHtml();
            await ctx.sync();
            const html = htmlResult.value;

            let extractedHtmlCode = [];
            let htmlSuccess = false;

            if (html) {
                const parser = new DOMParser();
                const doc = parser.parseFromString(html, 'text/html');
                const rows = doc.querySelectorAll('tr');
                
                if (rows.length > 0) {
                    rows.forEach(row => {
                        const cells = row.querySelectorAll('td');
                        // 逻辑：如果有多个单元格，取最后一个；如果只有一个，就取那一个
                        let codeCell = null;
                        if (cells.length >= 2) codeCell = cells[cells.length - 1];
                        else if (cells.length === 1) codeCell = cells[0];

                        if (codeCell) {
                            let text = codeCell.textContent || codeCell.innerText;
                            text = text.replace(/\u00a0/g, ' '); 
                            extractedHtmlCode.push(text.replace(/[\r\n]+$/, ''));
                        }
                    });
                    if (extractedHtmlCode.length > 0) htmlSuccess = true;
                }
            }

            if (htmlSuccess) {
                $('#codeSource').val(normalizeIndentationText(extractedHtmlCode.join('\n')));
                return showStatus("✅ Extracted from table");
            }

            // Fallback: text parsing
            range.load("text");
            await ctx.sync();
            let rawText = range.text;
            
            if (rawText && rawText.trim()) {
                const lines = rawText.split(/\r\n|\r|\n/);
                const cleanedLines = lines.map(line => {
                    return line.replace(/^\s*\d+\s*/, '');
                });
                
                $('#codeSource').val(normalizeIndentationText(cleanedLines.join('\n')));
                showStatus("✅ Extracted (text)");
            } else {
                showStatus("⚠️ Nothing selected", "error");
            }
        });
    } catch(e){
        console.error(e);
        showStatus("❌ Extract failed", "error");
    }
}

// Smart locate in document
export async function locateInDoc(code) {
    if (!code) return;
    
    const lines = code.split('\n').map(l => l.trim()).filter(l => l.length > 0);
    if (lines.length === 0) return;

    let searchCandidates = [];

    let maxLine = "";
    for(let l of lines) {
        if(l.length > maxLine.length && l.length < 200) maxLine = l;
    }
    if (maxLine) searchCandidates.push(maxLine);

    if (lines[0].length > 5) searchCandidates.push(lines[0]);

    if (lines[lines.length-1].length > 5) searchCandidates.push(lines[lines.length-1]);

    searchCandidates = [...new Set(searchCandidates)];

    if (searchCandidates.length === 0) return showStatus("⚠️ Code too short", "error");

    try {
        await Word.run(async (ctx) => {
            let foundRange = null;

            for (let key of searchCandidates) {
                const results = ctx.document.body.search(key, { matchCase: true, ignoreSpace: true });
                ctx.load(results);
                await ctx.sync();

                if (results.items.length > 0) {
                    foundRange = results.items[0];
                    break;
                }
            }

            if (foundRange) {
                const parentTable = foundRange.parentTableOrNullObject;
                ctx.load(parentTable);
                await ctx.sync();

                if (!parentTable.isNullObject) {
                    parentTable.select();
                    showStatus("✅ 已定位 (整块)");
                } else {
                    foundRange.select();
                    showStatus("✅ 已定位 (单行)");
                }
                
                ctx.document.getSelection().context.sync();
            } else {
                showStatus("⚠️ 文档中未找到", "error");
            }
        });
    } catch(e){ console.error(e); }
}
