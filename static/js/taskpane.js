/* static/js/taskpane.js v4.7 - 智能表格全选吸取 */

// 全局变量
let deleteTarget = null;
let confirmModal = null;
let currentEditingId = null;
let searchTimer = null;
let hljsConfigured = false;
let listingCounter = 1;
let explanationCollapsed = false;
let lastExplanationContent = ''; // 存储原始解释内容用于复制

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
            $(document).ready(function () {
                console.log("✅ CodeWeaver v4.6 Ready");
            
            // 1. 初始化
            syncProjectName();
            buildLanguageDropdown();
            ensureHighlighter();
            if (typeof marked !== 'undefined') {
                marked.setOptions({ gfm: true, breaks: true });
            }
            loadSnippets();
            confirmModal = new bootstrap.Modal(document.getElementById('confirmModal'));

            // 2. 绑定静态按钮
            $('#btnSave').click(saveSnippet);
            $('#btnInsert').click(insertHighlight);
            $('#btnGetSelection').click(getFromSelection);
            $('#btnNormalize').click(applyIndentationNormalization);
            $('#btnExplain').click(requestExplanation);
             $('#btnRenumber').click(renumberListings);
            $('#toggleExplain').click(toggleExplainPanel);
            $('#btnCopyExplain').click(copyExplanation);
            setExplainVisibility(true);
            
            // 3. 绑定静态按钮 (项目库页)
            $('#btnRefresh').click(() => loadSnippets($('#searchBox').val()));
            $('#library-tab').click(() => loadSnippets($('#searchBox').val()));

            // 4. 事件委托
            $(document).on('click', '.action-load-editor', function() {
                const code = decodeURIComponent($(this).data('code'));
                const lang = $(this).data('lang');
                $('#codeSource').val(code);
                $('#langSelect').val(lang);
                clearEditingState();
                new bootstrap.Tab('#editor-tab').show();
            });

            $(document).on('click', '.action-edit', function() {
                const code = decodeURIComponent($(this).data('code'));
                const lang = $(this).data('lang');
                const title = $(this).data('title');
                const project = $(this).data('project');
                currentEditingId = $(this).data('id');

                $('#codeSource').val(code);
                $('#langSelect').val(lang);
                $('#inputTitle').val(title);
                $('#inputProject').val(project);
                updateEditingState(title, project);
                new bootstrap.Tab('#editor-tab').show();
            });

            $(document).on('click', '.action-locate', function() {
                const code = decodeURIComponent($(this).data('code'));
                locateInDoc(code);
            });

            $(document).on('click', '.action-del-snippet', function() {
                const id = $(this).data('id');
                const title = $(this).data('title');
                askDeleteSnippet(id, title);
            });

            $(document).on('click', '.action-del-project', function() {
                const name = $(this).data('name');
                askDeleteProject(name);
            });

            $('#btnConfirmDelete').click(performDelete);

            // 5. 搜索过滤
            $('#searchBox').on('keyup', function() {
                const val = $(this).val();
                if (searchTimer) clearTimeout(searchTimer);
                searchTimer = setTimeout(() => loadSnippets(val), 250);
            });
        });
    }
});

// --- 逻辑函数 ---

function showStatus(msg, type='info') {
    const color = type === 'error' ? 'text-danger' : 'text-success';
    $('#statusMsg').html(`<span class="${color}">${msg}</span>`);
    setTimeout(() => $('#statusMsg').empty(), 3000);
}

function setExplainVisibility(show) {
    explanationCollapsed = !show;
    const $result = $('#aiExplainResult');
    const $toggle = $('#toggleExplain');
    if (show) {
        $result.removeClass('d-none');
        $toggle.text('Hide');
    } else {
        $result.addClass('d-none');
        $toggle.text('Show');
    }
}

function toggleExplainPanel() {
    setExplainVisibility(explanationCollapsed);
}

function copyExplanation() {
    if (!lastExplanationContent) {
        showStatus("⚠️ Nothing to copy", "error");
        return;
    }
    navigator.clipboard.writeText(lastExplanationContent).then(() => {
        showStatus("✅ Copied!", "success");
    }).catch(() => {
        showStatus("❌ Copy failed", "error");
    });
}

function renderExplanation(content, isRaw = false) {
    const $result = $('#aiExplainResult');
    if (isRaw) {
        lastExplanationContent = content || '';
    }
    if (content && isRaw) {
        try {
            if (typeof marked !== 'undefined') {
                const html = typeof marked.parse === 'function' 
                    ? marked.parse(content) 
                    : marked(content);
                $result.html(html);
            } else {
                $result.html(content.replace(/\n/g, '<br>'));
            }
        } catch (e) {
            console.error('Markdown parse error:', e);
            $result.html(content.replace(/\n/g, '<br>'));
        }
    } else {
        $result.html(content || '');
    }
}

function normalizeIndentationText(raw, language = '') {
    if (!raw) return '';
    const tabSize = 4;
    let text = raw.replace(/\t/g, ' '.repeat(tabSize));
    let lines = text.split(/\r?\n/);

    while (lines.length && lines[lines.length - 1].trim() === '') {
        lines.pop();
    }
    while (lines.length && lines[0].trim() === '') {
        lines.shift();
    }

    let minIndent = null;
    lines.forEach(line => {
        if (!line.trim()) return;
        const match = line.match(/^(\s+)/);
        const indentLen = match ? match[1].length : 0;
        if (minIndent === null || indentLen < minIndent) minIndent = indentLen;
    });

    if (minIndent && minIndent > 0) {
        lines = lines.map(line => {
            if (!line.trim()) return '';
            return line.startsWith(' '.repeat(minIndent)) ? line.slice(minIndent) : line.replace(/^\s+/, '');
        });
    }

    lines = lines.map(line => line.replace(/\s+$/, ''));
    return lines.join('\n');
}

function applyIndentationNormalization() {
    const code = $('#codeSource').val();
    if (!code) return showStatus("⚠️ No code", "error");
    const lang = $('#langSelect').val();
    const normalized = normalizeIndentationText(code, lang);
    $('#codeSource').val(normalized);
    showStatus("✅ Formatted");
}

function ensureHighlighter() {
    if (typeof hljs === 'undefined') return;
    if (!hljsConfigured) {
        hljs.configure({ ignoreUnescapedHTML: true });
        hljsConfigured = true;
    }
}

function buildLanguageDropdown() {
    if (typeof hljs === 'undefined') return;
    const common = ['python', 'java', 'c', 'cpp', 'javascript', 'typescript', 'html', 'css', 'sql', 'bash', 'json', 'go', 'php', 'ruby', 'csharp', 'swift', 'kotlin', 'rust'];
    const rest = hljs.listLanguages ? hljs.listLanguages().slice() : [];
    const remaining = rest.filter(l => !common.includes(l)).sort();
    const merged = ['auto', 'label_common', ...common, 'label_rest', ...remaining];

    const $select = $('#langSelect');
    $select.empty();

    merged.forEach(lang => {
        if (lang === 'label_common') {
            $select.append('<option disabled>--Common--</option>');
            return;
        }
        if (lang === 'label_rest') {
            $select.append('<option disabled>--A–Z--</option>');
            return;
        }
        let label = lang;
        if (lang === 'auto') label = '✨ Auto';
        else {
            const map = { cpp: 'C++', c: 'C', csharp: 'C#', javascript: 'JavaScript', typescript: 'TypeScript', sql: 'SQL', html: 'HTML', css: 'CSS', json: 'JSON', php: 'PHP', go: 'Go', ruby: 'Ruby', bash: 'Bash', kotlin: 'Kotlin', swift: 'Swift', rust: 'Rust', python: 'Python', java: 'Java' };
            label = map[lang] || lang.charAt(0).toUpperCase() + lang.slice(1);
        }
        $select.append(`<option value="${lang}">${label}</option>`);
    });
    $select.val('auto');
}

function updateEditingState(title, project) {
    $('#editState').html(`✏️ 正在编辑：<strong>${title}</strong> <span class="text-muted">@ ${project}</span>`);
}

function clearEditingState() {
    currentEditingId = null;
    $('#editState').empty();
}

function syncProjectName() {
    try {
        const url = Office.context.document.url;
        if (url) {
            let filename = url.substring(url.lastIndexOf('/') + 1);
            if (filename.indexOf('.') > -1) filename = filename.substring(0, filename.lastIndexOf('.'));
            filename = decodeURIComponent(filename);
            if (filename) $('#inputProject').val(filename);
        } else {
            const last = localStorage.getItem("last_project");
            if(last) $('#inputProject').val(last);
        }
    } catch (e) {}
}

async function saveSnippet() {
    const code = $('#codeSource').val();
    const project = $('#inputProject').val() || "Default";
    const title = $('#inputTitle').val();
    if (!code || !title) return showStatus("❌ Code & title required", "error");
    localStorage.setItem("last_project", project);

    try {
        showStatus("⏳ Saving...");
        const payload = { project, title, code, language: $('#langSelect').val() };
        if (currentEditingId) payload.id = currentEditingId;
        const res = await fetch('/api/snippets', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
        });
        if ((await res.json()).status === 'success') {
            showStatus("✅ Saved", "success");
            if (!currentEditingId) $('#inputTitle').val('');
            clearEditingState();
            loadSnippets($('#searchBox').val());
        } else showStatus("❌ Failed", "error");
    } catch (e) { showStatus("❌ Error", "error"); }
}

async function requestExplanation() {
    const code = $('#codeSource').val();
    if (!code) return showStatus("⚠️ No code", "error");
    const lang = $('#langSelect').val();

    const $result = $('#aiExplainResult');
    lastExplanationContent = ''; // 清空之前的内容
    setExplainVisibility(true);
    $result.removeClass('ai-error ai-ready').addClass('ai-loading');
    renderExplanation('⏳ AI 解读中...', false);
    try {
        const res = await fetch('/api/explain', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ code, language: lang })
        });
        const data = await res.json();
        if (data.status === 'success') {
            $result.removeClass('ai-loading ai-error').addClass('ai-ready');
            renderExplanation(data.explanation || '暂无解释', true);
        } else {
            $result.removeClass('ai-loading ai-ready').addClass('ai-error');
            renderExplanation(data.message || '解释失败', false);
        }
    } catch (e) {
        console.error(e);
        $result.removeClass('ai-ready ai-loading').addClass('ai-error');
        renderExplanation('网络异常', false);
    }
}

async function loadSnippets(keyword = '') {
    try {
        const params = new URLSearchParams({ t: Date.now() });
        if (keyword) params.append('q', keyword);
        const res = await fetch('/api/snippets?' + params.toString());
        const grouped = await res.json();
        const $cont = $('#gistContainer');
        $cont.empty();

        if (Object.keys(grouped).length === 0) {
            const msg = keyword ? '未找到匹配的代码' : '暂无代码';
            $cont.html(`<div class="text-center text-muted mt-4">${msg}</div>`);
            return;
        }

        for (const [projName, items] of Object.entries(grouped)) {
            let html = `
                <div class="project-card">
                    <div class="project-header">
                        <span>📂 ${projName}</span>
                        <button class="btn-del-proj action-del-project" data-name="${projName}">删除文件夹</button>
                    </div>
                    <div>
            `;
            items.forEach(item => {
                const safeCode = encodeURIComponent(item.code);
                html += `
                    <div class="snippet-item">
                        <div class="d-flex align-items-center text-truncate" style="flex:1;">
                            <span class="snippet-title text-truncate action-load-editor" 
                                  data-code="${safeCode}" 
                                  data-lang="${item.language}"
                                  title="点击编辑">
                                ${item.title}
                            </span>
                            <span class="badge-lang">${item.language}</span>
                        </div>
                        <div>
                            <button class="btn-action action-edit"
                                    data-id="${item.id}"
                                    data-code="${safeCode}"
                                    data-lang="${item.language}"
                                    data-title="${item.title}"
                                    data-project="${projName}"
                                    title="编辑">✏️</button>
                            <button class="btn-action btn-locate action-locate"
                                    data-code="${safeCode}"
                                    title="在文档中查找">🔍</button>
                                    
                            <button class="btn-action btn-delete action-del-snippet" 
                                    data-id="${item.id}" 
                                    data-title="${item.title}" 
                                    title="删除">🗑️</button>
                        </div>
                    </div>
                `;
            });
            html += `</div></div>`;
            $cont.append(html);
        }
    } catch (e) { console.error(e); }
}

function askDeleteSnippet(id, title) {
    deleteTarget = { type: 'snippet', id: id };
    $('#confirmMsg').text(`确认删除代码 "${title}" 吗？`);
    confirmModal.show();
}

function askDeleteProject(name) {
    deleteTarget = { type: 'project', name: name };
    $('#confirmMsg').html(`确认删除文件夹 <b>"${name}"</b> 吗？<br><small class="text-danger">这将删除里面的所有代码！</small>`);
    confirmModal.show();
}

async function performDelete() {
    if (!deleteTarget) return;
    confirmModal.hide();

    let url = '', method = '';
    let body = null;

    if (deleteTarget.type === 'snippet') {
        url = '/api/snippets/' + deleteTarget.id;
        method = 'DELETE';
    } else if (deleteTarget.type === 'project') {
        url = '/api/projects/delete';
        method = 'POST';
        body = JSON.stringify({ name: deleteTarget.name });
    }

    try {
        const opts = { method: method, headers: {'Content-Type': 'application/json'} };
        if(body) opts.body = body;
        
        const res = await fetch(url, opts);
        if ((await res.json()).status === 'success') {
            loadSnippets($('#searchBox').val());
        } else { alert("Delete failed"); }
    } catch (e) { alert("Network error"); }
}

// Renumber listings
async function renumberListings() {
    try {
        showStatus("⏳ Renumbering...");
        
        await Word.run(async (ctx) => {
            const paragraphs = ctx.document.body.paragraphs;
            ctx.load(paragraphs, 'text');
            await ctx.sync();
            
            const listingParagraphs = [];
            for (let i = 0; i < paragraphs.items.length; i++) {
                const paragraph = paragraphs.items[i];
                if (paragraph.text.match(/Listing\s+\d+:/)) {
                    listingParagraphs.push(paragraph);
                }
            }
            
            for (let i = 0; i < listingParagraphs.length; i++) {
                const paragraph = listingParagraphs[i];
                const oldText = paragraph.text;
                const match = oldText.match(/Listing\s+\d+:(.*)/);
                const description = match ? match[1] : '';
                const newText = `Listing ${i + 1}:${description}`;
                paragraph.insertText(newText, 'Replace');
            }
            
            await ctx.sync();
            listingCounter = listingParagraphs.length + 1;
        });
        
        showStatus(`✅ Renumbered`);
    } catch (e) {
        console.error(e);
        showStatus("❌ Renumber failed: " + e.message, "error");
    }
}

// Insert highlighted code
async function insertHighlight() {
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

// Generate highlighted HTML
function generateHighlightHtml(code, lang, theme, listingNo) {
    const normalizedCode = normalizeIndentationText(code, lang);
    if (!normalizedCode) return '';

    const syntaxThemes = {
        light: {
            'keyword': 'color:#d73a49; font-weight:bold;',
            'built_in': 'color:#005cc5;',
            'type': 'color:#005cc5;',
            'literal': 'color:#005cc5;',
            'number': 'color:#005cc5;',
            'string': 'color:#032f62;',
            'title': 'color:#6f42c1; font-weight:bold;',
            'attr': 'color:#22863a;',
            'comment': 'color:#6a737d; font-style:italic;',
            'variable': 'color:#24292f;',
            'symbol': 'color:#005cc5;',
            'function': 'color:#6f42c1;',
            'default': 'color:#24292f;'
        },
        dark: {
            'keyword': 'color:#f92672; font-weight:bold;',
            'built_in': 'color:#66d9ef;',
            'type': 'color:#66d9ef;',
            'literal': 'color:#ae81ff;',
            'number': 'color:#ae81ff;',
            'string': 'color:#e6db74;',
            'title': 'color:#a6e22e; font-weight:bold;',
            'attr': 'color:#a6e22e;',
            'comment': 'color:#75715e; font-style:italic;',
            'variable': 'color:#f8f8f2;',
            'symbol': 'color:#ae81ff;',
            'function': 'color:#a6e22e;',
            'default': 'color:#f8f8f2;'
        }
    };

    const currentSyntax = (theme === 'dark') ? syntaxThemes.dark : syntaxThemes.light;

    let bg_code = '#f6f8fa'; let bg_num = '#fff'; let color_code = '#24292f'; let color_num = '#6e7781'; let border = '#d0d7de';

    if (theme === 'dark') {
        bg_code = '#272822'; bg_num = '#fff'; color_code = '#f8f8f2'; border = '#272822';
    } else if (theme === 'green') {
        bg_code = '#e9f5e9'; border = '#e9f5e9';
    }

    const escapeHtml = (txt) => txt.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

    const style_common = "padding:0; margin:0; border:none; line-height:100%; vertical-align:middle;";
    //const style_num = `width:30px; background-color:${bg_num}; color:${color_num}; text-align:right; padding-right:5px; user-select:none; font-family:'Times New Roman'; font-size:6pt; ${style_common}`;
    const style_code = `width:100%; background-color:${bg_code}; color:${color_code}; padding-left:10px; font-family:'Courier New', monospace; font-size:10pt; white-space:pre; mso-no-proof:yes; ${style_common}`;
    const border_style = "1.5pt solid " + border;
    
    ensureHighlighter();

    let highlightedBlock = '';
    try {
        if (typeof hljs !== 'undefined' && hljs.highlight) {
            const hasLanguage = lang && lang !== 'auto' && hljs.getLanguage && hljs.getLanguage(lang);
            const res = hasLanguage
                ? hljs.highlight(normalizedCode, { language: lang, ignoreIllegals: true })
                : hljs.highlightAuto(normalizedCode);
            highlightedBlock = res.value || '';
        }
    } catch(e) { console.warn('highlight error', e); }

    if (!highlightedBlock) highlightedBlock = escapeHtml(normalizedCode);

    highlightedBlock = highlightedBlock.replace(/<span class="hljs-([^"]+)">/g, (match, cls) => {
        const key = cls.split(' ')[0];
        const style = currentSyntax[key] || '';
        return style ? `<span style="${style}">` : '<span>';
    });

    let lines = highlightedBlock.split(/\r?\n/);
    while (lines.length && lines[lines.length - 1] === '') lines.pop();

    let html = `<table style="width:100%; border-collapse:collapse; border-spacing:0; margin-bottom:10px; background-color:#fff;">`;
    lines.forEach((line, i) => {
        const lineHtml = line === '' ? '&nbsp;' : line;

        let cellBorder = `border-left:${border_style}; border-right:${border_style};`;
        if (i === 0) cellBorder += `border-top:${border_style};`;
        if (i === lines.length - 1) cellBorder += `border-bottom:${border_style};`;

        html += `<tr><td style="${style_code} ${cellBorder}">${lineHtml}</td></tr>`;
    });

    html += "</table>";
   const captionText = listingNo
  ? `Listing ${listingNo}:<span>&nbsp;</span>`
  : 'Listing:<span>&nbsp;</span>';

html += `<table style="width:100%; border-collapse:collapse; border-spacing:0; margin-top:4px;">
    <tr>
        <td style="text-align:center; font-family:'Times New Roman'; font-size:10.5pt; padding:0; border:none;">${captionText}</td>
    </tr>
</table>`;
    return html;
}
// 【关键修复：智能吸取模式】
async function getFromSelection() {
    try {
        await Word.run(async (ctx) => {
            // 1. 获取当前选区
            let range = ctx.document.getSelection();
            
            // 【核心逻辑】检查光标是否在表格内
            const parentTable = range.parentTableOrNullObject;
            ctx.load(parentTable);
            await ctx.sync();

            // 如果在表格里，强制把“选区”扩展为“整个表格”
            // 这样哪怕你只点了一下代码块，也能吸取全部代码！
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
async function locateInDoc(code) {
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