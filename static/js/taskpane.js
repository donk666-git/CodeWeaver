/* static/js/taskpane.js v4.7 - 智能表格全选吸取 */

// 全局变量
let deleteTarget = null;
let confirmModal = null;
let currentEditingId = null;
let searchTimer = null;
let hljsConfigured = false;
let listingCounter = 1;

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
            $(document).ready(function () {
                console.log("✅ CodeWeaver v4.6 Ready");
            
            // 1. 初始化
            syncProjectName();
            buildLanguageDropdown();
            ensureHighlighter();
            loadSnippets();
            renumberListings();
            confirmModal = new bootstrap.Modal(document.getElementById('confirmModal'));

            // 2. 绑定静态按钮
            $('#btnSave').click(saveSnippet);
            $('#btnInsert').click(insertHighlight);
            $('#btnGetSelection').click(getFromSelection);
            $('#btnNormalize').click(applyIndentationNormalization);
            $('#btnExplain').click(requestExplanation);
            
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

function normalizeIndentationText(raw, language = '') {
    if (!raw) return '';
    
    // 1. 预处理：统一换行符，移除首尾空行
    let text = raw.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
    let lines = text.split('\n');
    
    // 移除首尾空行
    while (lines.length && lines[lines.length - 1].trim() === '') lines.pop();
    while (lines.length && lines[0].trim() === '') lines.shift();
    
    if (lines.length === 0) return '';
    
    const lang = (language || '').toLowerCase();
    
    // 2. 特殊处理：如果是Python，使用Python专用逻辑
    if (lang.startsWith('python')) {
        return normalizePythonIndentation(lines);
    }
    
    // 3. 其他语言：先处理多语句，再规范化缩进
    lines = expandMultiStatements(lines, lang);
    const indentUnit = detectIndentUnit(lines);
    
    // 规范化缩进
    let depth = 0;
    const normalized = [];
    
    lines.forEach(line => {
        const trimmed = line.trim();
        if (!trimmed) {
            normalized.push('');
            return;
        }
        
        // 计算缩进调整
        const adjust = calculateIndentAdjust(trimmed, lang);
        const baseDepth = Math.max(depth - adjust.decreaseBefore, 0);
        const rebuilt = ' '.repeat(baseDepth * indentUnit) + trimmed;
        normalized.push(rebuilt);
        depth = Math.max(baseDepth + adjust.increaseAfter, 0);
    });
    
    return normalized.join('\n');
}

// 新增：Python专用缩进规范化
function normalizePythonIndentation(lines) {
    const normalized = [];
    const indentUnit = 4; // Python标准缩进为4空格
    let depth = 0;
    
    lines.forEach(line => {
        const trimmed = line.trim();
        if (!trimmed) {
            normalized.push('');
            return;
        }
        
        // 计算当前行的实际缩进级别
        let currentDepth = 0;
        const leadingSpaces = line.length - line.ltrimStart().length;
        if (leadingSpaces > 0) {
            currentDepth = Math.round(leadingSpaces / indentUnit);
        }
        
        // 处理特殊行
        let targetDepth = depth;
        
        // 减少缩进的情况
        if (/^(elif|else|except|finally)\b/.test(trimmed)) {
            targetDepth = Math.max(depth - 1, 0);
        } else if (/^[}\]\)]/.test(trimmed)) {
            // 虽然Python不用大括号，但为了兼容性保留
            targetDepth = Math.max(depth - 1, 0);
        }
        
        // 生成规范化行
        normalized.push(' '.repeat(targetDepth * indentUnit) + trimmed);
        
        // 计算下一行的深度
        if (/^def\s+|^class\s+|^if\s+|^elif\s+|^else\s*:\s*$|^for\s+|^while\s+|^try\s*:\s*$|^except\s+|^finally\s*:\s*$|^with\s+/.test(trimmed)) {
            if (/:\s*$/.test(trimmed)) {
                depth = targetDepth + 1;
            } else {
                depth = targetDepth;
            }
        } else if (/^(elif|else|except|finally)\b/.test(trimmed)) {
            depth = targetDepth + 1;
        } else {
            depth = targetDepth;
        }
    });
    
    return normalized.join('\n');
}

// 新增：规范化现有缩进
function normalizeExistingIndentation(lines, indentUnit) {
    return lines.map(line => {
        const trimmed = line.trimEnd();
        const content = trimmed.trim();
        if (!content) return '';
        
        // 计算当前缩进空格数
        const leadingSpaces = trimmed.length - trimmed.ltrimStart().length;
        // 规范化为指定单位的倍数
        const normalizedIndent = Math.round(leadingSpaces / indentUnit) * indentUnit;
        
        return ' '.repeat(normalizedIndent) + content;
    });
}

function detectIndentUnit(lines) {
    const counts = [];
    lines.forEach(line => {
        const match = line.match(/^(\s+)/);
        if (match) {
            const spaces = match[1].length;
            if (spaces > 0 && spaces < 20) {
                counts.push(spaces);
            }
        }
    });
    
    if (counts.length === 0) return 4;
    
    // 找出最常见的缩进单位
    const freq = {};
    counts.forEach(n => {
        const unit = n % 4 === 0 ? 4 : n % 2 === 0 ? 2 : n;
        freq[unit] = (freq[unit] || 0) + 1;
    });
    
    let best = 4, bestCount = 0;
    Object.entries(freq).forEach(([unit, cnt]) => {
        if (cnt > bestCount) { 
            bestCount = cnt; 
            best = parseInt(unit, 10); 
        }
    });
    
    return best || 4;
}

// String polyfill
if (!String.prototype.trimEnd) {
    String.prototype.trimEnd = function() {
        return this.replace(/\s+$/, '');
    };
}

if (!String.prototype.ltrimStart) {
    String.prototype.ltrimStart = function() {
        return this.replace(/^\s+/, '');
    };
}

function expandMultiStatements(lines, language) {
    const targetLangs = [
        'javascript', 'js', 'typescript', 'ts', 
        'java', 'c', 'cpp', 'csharp', 'cs',
        'php', 'swift', 'kotlin', 'go', 'rust'
    ];
    const applicable = targetLangs.includes(language);
    if (!applicable) return lines;

    const splitSafe = (line) => {
        const segments = [];
        let buf = '';
        let inStr = false;
        let strChar = '';
        let parenDepth = 0;
        let braceDepth = 0;
        
        const pushBuf = () => {
            const val = buf.trim();
            if (val) segments.push(val);
            buf = '';
        };

        const trimmed = line.trim();
        
        // 不拆分的情况
        if (/^(for|while)\s*\([^)]*\)/i.test(trimmed)) return [line];
        if (/^if\s*\([^)]*\)\s*[^{]/.test(trimmed)) return [line];
        if (/^}\s*else\s*/.test(trimmed)) return [line];
        if (/^}\s*catch\s*\(/.test(trimmed)) return [line];
        if (/^}\s*finally/.test(trimmed)) return [line];

        for (let i = 0; i < line.length; i++) {
            const ch = line[i];
            const prev = i > 0 ? line[i - 1] : '';
            
            if (inStr) {
                buf += ch;
                if (ch === strChar && prev !== '\\') {
                    inStr = false;
                    strChar = '';
                }
                continue;
            }
            
            if (ch === '"' || ch === '\'' || ch === '`') {
                inStr = true; 
                strChar = ch; 
                buf += ch; 
                continue;
            }
            
            if (ch === '(') parenDepth += 1;
            if (ch === ')' && parenDepth > 0) parenDepth -= 1;
            if (ch === '{') braceDepth += 1;
            if (ch === '}' && braceDepth > 0) braceDepth -= 1;
            
            // 在括号深度为0且不在字符串中时，按分号拆分
            if (ch === ';' && parenDepth === 0 && braceDepth === 0) {
                pushBuf();
                continue;
            }
            buf += ch;
        }
        pushBuf();
        return segments.length ? segments : [line.trimEnd()];
    };

    return lines.flatMap(splitSafe);
}
function calculateIndentAdjust(content, language) {
    let decreaseBefore = 0;
    let increaseAfter = 0;
    const lang = (language || '').toLowerCase();

    // 处理结束符号
    if (/^[}\]\)]/.test(content)) {
        const closing = content.match(/^[}\]\)]+/);
        decreaseBefore = closing ? closing[0].length : 0;
    }

    // 大括号语言的处理
    const tokens = countBraceChanges(content);
    decreaseBefore = Math.max(decreaseBefore, tokens.close);
    const net = tokens.open - tokens.close;
    if (net > 0) increaseAfter += net;
    
    // 处理 else, catch, finally 等关键字
    if (/\b(else|catch|finally)\b/.test(content) && !/\{/.test(content)) {
        decreaseBefore = Math.max(decreaseBefore, 1);
        increaseAfter += 1;
    }
    
    // 处理 case 语句
    if (/^(case\s+\w+|default)\s*:\s*$/.test(content)) {
        // case 通常与 switch 同级
    }
    
    // 处理标签
    if (/^\w+\s*:\s*$/.test(content) && !lang.includes('javascript') && !lang.includes('typescript')) {
        // 标签不缩进
    }

    return { decreaseBefore, increaseAfter };
}

// 改进的大括号计数
function countBraceChanges(content) {
    let open = 0, close = 0;
    let inStr = false;
    let strChar = '';
    let inComment = false;
    
    for (let i = 0; i < content.length; i++) {
        const ch = content[i];
        const prev = i > 0 ? content[i - 1] : '';
        const next = i < content.length - 1 ? content[i + 1] : '';
        
        // 处理注释
        if (!inStr && !inComment) {
            if (ch === '/' && next === '/') {
                break; // 单行注释
            }
            if (ch === '/' && next === '*') {
                inComment = true;
                i++;
                continue;
            }
        }
        
        if (inComment) {
            if (ch === '*' && next === '/') {
                inComment = false;
                i++;
            }
            continue;
        }
        
        // 处理字符串
        if (inStr) {
            if (ch === strChar && prev !== '\\') {
                inStr = false;
                strChar = '';
            }
            continue;
        }
        
        if (ch === '"' || ch === '\'' || ch === '`') {
            inStr = true; 
            strChar = ch; 
            continue;
        }
        
        // 计数大括号
        if (ch === '{') open += 1;
        else if (ch === '}') close += 1;
    }
    
    return { open, close };
}
function applyIndentationNormalization() {
    const code = $('#codeSource').val();
    if (!code) return showStatus("⚠️ 当前无代码", "error");
    const lang = $('#langSelect').val();
    const normalized = normalizeIndentationText(code, lang);
    $('#codeSource').val(normalized);
    showStatus("✅ 缩进已整理");
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
            $select.append('<option disabled>常用</option>');
            return;
        }
        if (lang === 'label_rest') {
            $select.append('<option disabled>A–Z</option>');
            return;
        }
        let label = lang;
        if (lang === 'auto') label = '✨ 自动检测';
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
    const project = $('#inputProject').val() || "默认";
    const title = $('#inputTitle').val();
    if (!code || !title) return showStatus("❌ 请填写代码和标题", "error");
    localStorage.setItem("last_project", project);

    try {
        showStatus("⏳ 保存中...");
        const payload = { project, title, code, language: $('#langSelect').val() };
        if (currentEditingId) payload.id = currentEditingId;
        const res = await fetch('/api/snippets', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
        });
        if ((await res.json()).status === 'success') {
            showStatus("✅ 成功", "success");
            if (!currentEditingId) $('#inputTitle').val('');
            clearEditingState();
            loadSnippets($('#searchBox').val());
        } else showStatus("❌ 失败", "error");
    } catch (e) { showStatus("❌ 错误", "error"); }
}

async function requestExplanation() {
    const code = $('#codeSource').val();
    if (!code) return showStatus("⚠️ 当前无代码", "error");
    const lang = $('#langSelect').val();

    $('#aiExplainResult').text('⏳ AI 解读中...');
    try {
        const res = await fetch('/api/explain', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ code, language: lang })
        });
        const data = await res.json();
        if (data.status === 'success') {
            $('#aiExplainResult').text(data.explanation || '暂无解释');
        } else {
            $('#aiExplainResult').text(data.message || '解释失败');
        }
    } catch (e) {
        console.error(e);
        $('#aiExplainResult').text('网络异常');
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
        } else { alert("删除失败"); }
    } catch (e) { alert("网络错误"); }
}

async function insertHighlight() {
    const code = $('#codeSource').val();
    const lang = $('#langSelect').val();
    const theme = $('#themeSelect').val();

    if (!code) return showStatus("❌ 代码为空", "error");
    try {
        const renumberedNext = await renumberListings();
        const html = generateHighlightHtml(code, lang, theme, renumberedNext || null);
        await Word.run(async (ctx)=>{
            ctx.document.getSelection().insertHtml(html, 'Replace');
            await ctx.sync();
        });
        const recalculated = await renumberListings();
        if (recalculated !== null) listingCounter = recalculated;
        showStatus("✅ 成功插入");
    } catch (e) {
        console.error(e);
        showStatus("❌ 插入失败:"+ e.message, "error");
    }
}

/**
 * 本地生成高亮 HTML (基于 highlight.js)
 * 复刻原 Python 后端逻辑，保留表格样式和 Word 兼容性
 */
/**
 * 本地生成高亮 HTML (终极版：修复行距 + 内联颜色样式)
 */
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
    // 注释掉行号样式
    //const style_num = `width:30px; background-color:${bg_num}; color:${color_num}; text-align:right; padding-right:5px; user-select:none; font-family:'Times New Roman'; font-size:6pt; ${style_common}`;
    const style_code = `width:100%; background-color:${bg_code}; color:${color_code}; padding-left:10px; font-family:'Courier New', monospace; font-size:10pt; white-space:pre; mso-no-proof:yes; ${style_common}`;
    const border_style = "1.5pt solid " + border;
    // 不再需要偏移，因为我们移除了行号列
    const table_width = `100%`;
    const table_margin_left = `0`;

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

    let html = `<table style="width:${table_width}; border-collapse:collapse; border-spacing:0; margin-bottom:10px; margin-left:${table_margin_left}; background-color:#fff;">`;
    lines.forEach((line, i) => {
        const lineHtml = line === '' ? '&nbsp;' : line;

        // 恢复原来的边框逻辑：只给第一行添加上边框，给最后一行添加下边框
        let cellBorder = `border-left:${border_style}; border-right:${border_style};`;
        if (i === 0) cellBorder += `border-top:${border_style};`;
        if (i === lines.length - 1) cellBorder += `border-bottom:${border_style};`;

        // 移除行号列，只保留代码列
        html += `<tr><td style="${style_code} ${cellBorder}">${lineHtml}</td></tr>`;
    });

    html += "</table>";
    const captionText = listingNo ? `Listing ${listingNo}: ` : 'Listing: ';
    html += `<div style="text-align:center; font-family:'Times New Roman'; font-size:10.5pt; margin-top:4px;">${captionText}</div>`;
    return html;
}

async function renumberListings() {
    let next = null;
    try {
        await Word.run(async (ctx) => {
            const results = ctx.document.body.search('Listing', { matchCase: false });
            results.load('items');
            await ctx.sync();
            results.items.forEach(r => r.load('text'));
            await ctx.sync();

            let counter = 1;
            results.items.forEach(range => {
                const raw = (range.text || '').replace(/\s+/g, ' ').trim();
                if (/^Listing\s*(\d+)?\s*:\s*$/i.test(raw) || /^Listing:\s*$/i.test(raw)) {
                    range.insertText(`Listing ${counter}: `, 'Replace');
                    counter += 1;
                }
            });
            await ctx.sync();
            next = counter;
        });
    } catch (e) {
        console.warn('renumber listings failed', e);
    }
    if (next !== null) listingCounter = next;
    return next;
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
                return showStatus("✅ 已从表格吸取");
            }

            // 3. 尝试文本强力解析 (备用)
            range.load("text");
            await ctx.sync();
            let rawText = range.text;
            
            if (rawText && rawText.trim()) {
                const lines = rawText.split(/\r\n|\r|\n/);
                const cleanedLines = lines.map(line => {
                    // 正则增强：移除行首的数字和空白
                    return line.replace(/^\s*\d+\s*/, '');
                });
                
                $('#codeSource').val(normalizeIndentationText(cleanedLines.join('\n')));
                showStatus("✅ 已吸取 (文本模式)");
            } else {
                showStatus("⚠️ 未选中内容", "error");
            }
        });
    } catch(e){
        console.error(e);
        showStatus("❌ 吸取失败", "error");
    }
}

// 【智能定位】
async function locateInDoc(code) {
    if (!code) return;
    
    const lines = code.split('\n').map(l => l.trim()).filter(l => l.length > 0);
    if (lines.length === 0) return;

    let searchCandidates = [];

    // 1. 最长的一行 (最独特，首选)
    let maxLine = "";
    for(let l of lines) {
        if(l.length > maxLine.length && l.length < 200) maxLine = l;
    }
    if (maxLine) searchCandidates.push(maxLine);

    // 2. 第一行 (如果不短的话)
    if (lines[0].length > 5) searchCandidates.push(lines[0]);

    // 3. 最后一行 (如果不短的话)
    if (lines[lines.length-1].length > 5) searchCandidates.push(lines[lines.length-1]);

    searchCandidates = [...new Set(searchCandidates)];

    if (searchCandidates.length === 0) return showStatus("⚠️ 代码太短无法定位", "error");

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
