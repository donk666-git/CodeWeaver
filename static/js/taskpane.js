/* static/js/taskpane.js v4.5 - 智能表格全选吸取 */

// 全局变量
let deleteTarget = null;
let confirmModal = null;
let currentSnippetId = null;
let searchTimer = null;
let explainModal = null;
let lastExplainText = '';

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        $(document).ready(function () {
            console.log("✅ CodeWeaver v4.5 Ready");

            if (window.hljs) {
                hljs.configure({ ignoreUnescapedHTML: true });
            }

            buildLanguageOptions();
            updateSaveButtonLabel();

            // 1. 初始化
            syncProjectName();
            loadSnippets();
            confirmModal = new bootstrap.Modal(document.getElementById('confirmModal'));
            explainModal = new bootstrap.Modal(document.getElementById('explainModal'));

            // 2. 绑定静态按钮
            $('#btnSave').click(saveSnippet);
            $('#btnNew').click(() => { resetEditorState(true); showStatus('🆕 新建空白'); });
            $('#btnInsert').click(insertHighlight);
            $('#btnExplain').click(explainCurrentCode);
            $('#btnGetSelection').click(getFromSelection);
            $('#btnCopyExplain').click(copyExplainText);
            
            // 3. 绑定静态按钮 (项目库页)
            $('#btnRefresh').click(() => loadSnippets($('#searchBox').val()));
            $('#library-tab').click(() => loadSnippets($('#searchBox').val()));

            $('#langSelect').on('change', function() {
                // 用户手动选择语言后仍保留列表顺序，不需要额外逻辑
            });

            // 4. 事件委托
            $(document).on('click', '.action-load-editor', function() {
                const code = decodeURIComponent($(this).data('code'));
                const lang = $(this).data('lang');
                const sid = $(this).data('id');
                const proj = $(this).data('project');
                const title = $(this).data('title');
                $('#codeSource').val(code);
                $('#langSelect').val(lang);
                if (proj) $('#inputProject').val(proj);
                if (title) $('#inputTitle').val(title);
                currentSnippetId = sid || null;
                updateSaveButtonLabel();
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

            // 5. 搜索过滤（后端模糊匹配：标题 / 代码 / 项目名）
            $('#searchBox').on('keyup', function() {
                const val = $(this).val();
                if (searchTimer) clearTimeout(searchTimer);
                searchTimer = setTimeout(() => loadSnippets(val), 220);
            });

            $('#aiProviderHint').text('AI 讲解由 DeepSeek 驱动，失败回落到本地快速总结');
        });
    }
});

// --- 逻辑函数 ---

function showStatus(msg, type='info') {
    const color = type === 'error' ? 'text-danger' : 'text-success';
    $('#statusMsg').html(`<span class="${color}">${msg}</span>`);
    setTimeout(() => $('#statusMsg').empty(), 3000);
}

const LANGUAGE_LABELS = {
    'bash': 'Bash / Shell',
    'c': 'C',
    'cpp': 'C++',
    'csharp': 'C#',
    'css': 'CSS',
    'go': 'Go',
    'html': 'HTML / XML',
    'java': 'Java',
    'javascript': 'JavaScript',
    'json': 'JSON',
    'kotlin': 'Kotlin',
    'lua': 'Lua',
    'matlab': 'MATLAB',
    'objectivec': 'Objective-C',
    'perl': 'Perl',
    'php': 'PHP',
    'python': 'Python',
    'r': 'R',
    'ruby': 'Ruby',
    'rust': 'Rust',
    'scala': 'Scala',
    'sql': 'SQL',
    'swift': 'Swift',
    'typescript': 'TypeScript',
    'yaml': 'YAML'
};

const COMMON_LANGS = [
    'python','javascript','java','c','cpp','csharp','go','rust','php','typescript','sql','bash','html','css','json','yaml','kotlin','swift','matlab'
];

function buildLanguageOptions() {
    const $select = $('#langSelect');
    if (!$select.length) return;

    const available = (window.hljs && typeof hljs.listLanguages === 'function') ? hljs.listLanguages() : [];
    const availableSet = available.length ? new Set(available) : null;
    const labelFor = (lang) => LANGUAGE_LABELS[lang] || lang.toUpperCase();

    const allCandidates = available.length ? available : Array.from(new Set([...COMMON_LANGS, ...Object.keys(LANGUAGE_LABELS)]));

    $select.empty();
    $select.append('<option value="auto">✨ 自动检测</option>');

    const commonOptions = [];
    COMMON_LANGS.forEach(lang => {
        if (!availableSet || availableSet.has(lang)) {
            commonOptions.push(`<option value="${lang}">${labelFor(lang)}</option>`);
        }
    });
    if (commonOptions.length) {
        $select.append(`<optgroup label="常用语言">${commonOptions.join('')}</optgroup>`);
    }

    const others = allCandidates
        .filter(lang => COMMON_LANGS.indexOf(lang) === -1)
        .filter(lang => !availableSet || availableSet.has(lang))
        .sort((a, b) => a.localeCompare(b));

    if (others.length) {
        const otherOpts = others.map(lang => `<option value="${lang}">${labelFor(lang)}</option>`);
        $select.append(`<optgroup label="全部 (A-Z)">${otherOpts.join('')}</optgroup>`);
    }
}

function updateSaveButtonLabel() {
    $('#btnSave').text(currentSnippetId ? '💾 更新' : '💾 保存');
}

function resetEditorState(clearFields = false) {
    currentSnippetId = null;
    if (clearFields) {
        $('#codeSource').val('');
        $('#inputTitle').val('');
        $('#langSelect').val('auto');
    }
    updateSaveButtonLabel();
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
        const res = await fetch('/api/snippets', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ project, title, code, language: $('#langSelect').val(), id: currentSnippetId })
        });
        const payload = await res.json();
        if (payload.status === 'success') {
            currentSnippetId = payload.id;
            updateSaveButtonLabel();
            showStatus(payload.mode === 'update' ? "✅ 已更新" : "✅ 已保存", "success");
            loadSnippets($('#searchBox').val());
        } else showStatus("❌ 失败", "error");
    } catch (e) { showStatus("❌ 错误", "error"); }
}

async function explainCurrentCode() {
    const code = $('#codeSource').val();
    const language = $('#langSelect').val() || 'auto';
    if (!code.trim()) return showStatus("⚠️ 没有可讲解的代码", "error");

    const $btn = $('#btnExplain');
    const prevText = $btn.text();
    lastExplainText = '';

    try {
        $('#explainContent').text('⏳ 正在调用 DeepSeek...');
        setExplainBadge('pending');
        explainModal.show();
        $btn.prop('disabled', true).text('🤖 讲解中...');
        const res = await fetch('/api/explain', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ code, language })
        });
        if (!res.ok) {
            const text = await res.text();
            throw new Error(text || `HTTP ${res.status}`);
        }
        const payload = await res.json();
        if (payload.status === 'success') {
            lastExplainText = payload.explanation || '';
            $('#explainContent').text(lastExplainText || '暂无结果');
            setExplainBadge(payload.provider === 'deepseek' ? 'deepseek' : 'local');
            $('#aiProviderHint').text(payload.provider === 'deepseek' ? '讲解来源：DeepSeek' : '讲解来源：本地摘要（未调用外部接口）');
        } else {
            $('#explainContent').text('❌ 解析失败：' + (payload.message || '未知错误'));
            setExplainBadge('failed');
        }
    } catch (e) {
        $('#explainContent').text('❌ 解析失败：' + e.message);
        setExplainBadge('failed');
    } finally {
        $btn.prop('disabled', false).text(prevText);
    }
}

function setExplainBadge(provider) {
    const $badge = $('#aiProviderBadge');
    const $meta = $('#aiExplainMeta');
    if (provider === 'deepseek') {
        $badge.text('DeepSeek').removeClass('bg-secondary').addClass('bg-gradient-blue');
        $meta.text('由 DeepSeek 生成的详细讲解');
    } else if (provider === 'local') {
        $badge.text('本地摘要').removeClass('bg-gradient-blue').addClass('bg-secondary');
        $meta.text('外部调用失败，使用快速本地总结');
    } else if (provider === 'failed') {
        $badge.text('出错').removeClass('bg-gradient-blue').addClass('bg-secondary');
        $meta.text('调用失败，请稍后重试');
    } else {
        $badge.text('准备中').removeClass('bg-secondary').addClass('bg-gradient-blue');
        $meta.text('DeepSeek 优先 · 支持自动降级');
    }
}

async function copyExplainText() {
    const text = lastExplainText || $('#explainContent').text();
    if (!text.trim()) return showStatus('⚠️ 暂无可复制的讲解', 'error');

    try {
        if (navigator.clipboard && window.isSecureContext) {
            await navigator.clipboard.writeText(text);
        } else {
            const tmp = document.createElement('textarea');
            tmp.value = text;
            document.body.appendChild(tmp);
            tmp.select();
            document.execCommand('copy');
            document.body.removeChild(tmp);
        }
        showStatus('✅ 已复制讲解');
    } catch (e) {
        showStatus('❌ 复制失败', 'error');
    }
}

async function loadSnippets(keyword = '') {
    try {
        const searchParam = keyword ? `&q=${encodeURIComponent(keyword)}` : '';
        const res = await fetch(`/api/snippets?t=${Date.now()}${searchParam}`);
        const grouped = await res.json();
        const $cont = $('#gistContainer');
        $cont.empty();

        if (Object.keys(grouped).length === 0) {
            $cont.html('<div class="text-center text-muted mt-4">暂无代码</div>');
            return;
        }

        for (const [projName, items] of Object.entries(grouped)) {
            const safeProj = projName.replace(/"/g, '&quot;');
            const displayProj = projName.replace(/</g, '&lt;').replace(/>/g, '&gt;');
            let html = `
                <div class="project-card">
                    <div class="project-header">
                        <span>📂 ${displayProj}</span>
                        <button class="btn-del-proj action-del-project" data-name="${safeProj}">删除文件夹</button>
                    </div>
                    <div>
            `;
            items.forEach(item => {
                const safeCode = encodeURIComponent(item.code);
                const safeTitle = (item.title || '').replace(/"/g, '&quot;');
                const displayTitle = (item.title || '').replace(/</g, '&lt;').replace(/>/g, '&gt;');
                html += `
                    <div class="snippet-item">
                        <div class="d-flex align-items-center text-truncate" style="flex:1;">
                            <span class="snippet-title text-truncate action-load-editor"
                                  data-id="${item.id}"
                                  data-title="${safeTitle}"
                                  data-project="${safeProj}"
                                  data-code="${safeCode}"
                                  data-lang="${item.language}"
                                  title="点击编辑">
                                ${displayTitle}
                            </span>
                            <span class="badge-lang">${item.language}</span>
                        </div>
                        <div>
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
        const html = generateHighlightHtml(code, lang, theme)
        await Word.run(async (ctx)=>{
            ctx.document.getSelection().insertHtml(html, 'Replace');
            await ctx.sync();
        });
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
function generateHighlightHtml(code, lang, theme) {
    if (!code) return '';

    // --- 1. 定义语法高亮颜色方案 (内联样式映射) ---
    // 分为 'light' (用于 gray/green 主题) 和 'dark' (用于 dark 主题)
    const syntaxThemes = {
        light: {
            'keyword': 'color:#d73a49; font-weight:bold;',       // 关键字 (红)
            'built_in': 'color:#005cc5;',                         // 内置函数 (蓝)
            'type': 'color:#005cc5;',                             // 类型
            'literal': 'color:#005cc5;',                          // 字面量
            'number': 'color:#005cc5;',                           // 数字
            'string': 'color:#032f62;',                           // 字符串 (深蓝)
            'title': 'color:#6f42c1; font-weight:bold;',          // 函数名 (紫)
            'attr': 'color:#22863a;',                             // 属性 (绿)
            'comment': 'color:#6a737d; font-style:italic;',       // 注释 (灰斜体)
            'variable': 'color:#24292f;',                         // 变量
            'symbol': 'color:#005cc5;',                           // 符号
            'function': 'color:#6f42c1;',                         // 函数调用
            'default': 'color:#24292f;'                           // 默认文本
        },
        dark: {
            'keyword': 'color:#f92672; font-weight:bold;',        // Monokai 风格
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

    // 根据用户选择的主题决定使用哪套语法颜色
    const currentSyntax = (theme === 'dark') ? syntaxThemes.dark : syntaxThemes.light;

    // --- 2. 主题参数 (背景 + 文字颜色，统一灰白基调) ---
    const themeMeta = {
        gray: { bg: '#f6f8fa', text: '#1f2933', border: '#d0d7de', shadow: '0 2px 8px rgba(17,24,39,0.08)', syntax: 'light' },
        green: { bg: '#f4f8f3', text: '#1f2a33', border: '#d6e4d1', shadow: '0 2px 8px rgba(15,118,110,0.08)', syntax: 'light' },
        dark: { bg: '#f3f4f6', text: '#111827', border: '#d1d5db', shadow: '0 3px 10px rgba(0,0,0,0.10)', syntax: 'light' }
    };
    const chosen = themeMeta[theme] || themeMeta.gray;

    // --- 3. 整块高亮，无行号 ---
    let highlighted = '';
    try {
        if (typeof hljs !== 'undefined') {
            const res = (lang && lang !== 'auto')
                ? hljs.highlight(code, {language: lang, ignoreIllegals:true})
                : hljs.highlightAuto(code);
            highlighted = res.value;
            highlighted = highlighted.replace(/<span class="hljs-([^"]+)">/g, (match, cls) => {
                const key = cls.split(' ')[0];
                const style = currentSyntax[key] || '';
                return style ? `<span style="${style}">` : '<span>';
            });
        } else {
            highlighted = code.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
        }
    } catch(e) {
        highlighted = code.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
    }

    // 处理缩进与换行：把每行的前导空格 / 制表符变成 &nbsp;，并显式用 <br> 断行，避免 Word 插入时丢失缩进或最后一行掉出框外
    const htmlLines = highlighted
        .split(/\r?\n/)
        .map(line => {
            if (!line.length) return '&nbsp;';
            return line.replace(/^([\t ]+)/, (m) => m
                .replace(/\t/g, '&nbsp;&nbsp;&nbsp;&nbsp;')
                .replace(/ /g, '&nbsp;')
            );
        })
        .join('<br/>');

    const preStyle = [
        'margin:0;',
        'padding:12px 14px;',
        `background:${chosen.bg};`,
        `border:1px solid ${chosen.border};`,
        'border-radius:10px;',
        `box-shadow:${chosen.shadow};`,
        "font-family:'Courier New', monospace;",
        'font-size:10pt;',
        'line-height:1.5;',
        'white-space:pre-wrap;',
        'word-break:break-word;',
        'tab-size:4;',
        'width:100%;',
        'box-sizing:border-box;',
        `color:${chosen.text};`
    ].join(' ');

    return `<div style="width:100%;"><pre style="${preStyle}">${htmlLines}</pre></div>`;
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
                } else {
                    const pre = doc.querySelector('pre');
                    if (pre) {
                        const text = (pre.textContent || '').replace(/\u00a0/g, ' ');
                        if (text.trim()) {
                            extractedHtmlCode.push(text);
                            htmlSuccess = true;
                        }
                    }
                }
            }

            if (htmlSuccess) {
                $('#codeSource').val(extractedHtmlCode.join('\n'));
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
                
                $('#codeSource').val(cleanedLines.join('\n'));
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
