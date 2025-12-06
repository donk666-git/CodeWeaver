/* static/js/taskpane.js v4.5 - 智能表格全选吸取 */

// 全局变量
let deleteTarget = null; 
let confirmModal = null;

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        $(document).ready(function () {
            console.log("✅ CodeWeaver v4.5 Ready");
            
            // 1. 初始化
            syncProjectName();
            loadSnippets();
            confirmModal = new bootstrap.Modal(document.getElementById('confirmModal'));

            // 2. 绑定静态按钮
            $('#btnSave').click(saveSnippet);
            $('#btnInsert').click(insertHighlight);
            $('#btnGetSelection').click(getFromSelection);
            
            // 3. 绑定静态按钮 (项目库页)
            $('#btnRefresh').click(loadSnippets);
            $('#library-tab').click(loadSnippets);

            // 4. 事件委托
            $(document).on('click', '.action-load-editor', function() {
                const code = decodeURIComponent($(this).data('code'));
                const lang = $(this).data('lang');
                $('#codeSource').val(code);
                $('#langSelect').val(lang);
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
                var val = $(this).val().toLowerCase();
                $(".snippet-item").each(function() {
                    $(this).toggle($(this).text().toLowerCase().indexOf(val) > -1);
                });
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
            body: JSON.stringify({ project, title, code, language: $('#langSelect').val() })
        });
        if ((await res.json()).status === 'success') {
            showStatus("✅ 成功", "success");
            $('#inputTitle').val('');
            loadSnippets();
        } else showStatus("❌ 失败", "error");
    } catch (e) { showStatus("❌ 错误", "error"); }
}

async function loadSnippets() {
    try {
        const res = await fetch('/api/snippets?t=' + Date.now());
        const grouped = await res.json();
        const $cont = $('#gistContainer');
        $cont.empty();

        if (Object.keys(grouped).length === 0) {
            $cont.html('<div class="text-center text-muted mt-4">暂无代码</div>');
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
            loadSnippets(); 
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

    // --- 2. 容器样式配置 ---
    let bg_code = '#f6f8fa'; let bg_num = '#fff'; let color_code = '#24292f'; let color_num = '#6e7781'; let border = '#d0d7de';
    
    if (theme === 'dark') { 
        bg_code = '#272822'; bg_num = '#fff'; color_code = '#f8f8f2'; border = '#272822'; 
    } else if (theme === 'green') {
        bg_code = '#e9f5e9'; border = '#e9f5e9'; // 护眼绿
    }
    
    // padding:0; margin:0; line-height:100% 是防止 Word 默认段落间距干扰的关键
    const style_common = "padding:0; margin:0; border:none; line-height:100%; vertical-align:middle;";
    const style_num = `width:30px; background-color:${bg_num}; color:${color_num}; text-align:right; padding-right:5px; user-select:none; font-family:'Times New Roman'; font-size:6pt; ${style_common}`;
    const style_code = `width:100%; background-color:${bg_code}; color:${color_code}; padding-left:10px; font-family:'Courier New', monospace; font-size:10pt; white-space:pre; mso-no-proof:yes; ${style_common}`;
    const border_style = "1.5pt solid " + border;

    // --- 3. 生成 HTML ---
    let html = `<table style="width:100%; border-collapse:collapse; border-spacing:0; margin-bottom:10px; background-color:#fff;">`;

    const lines = code.split(/\r?\n/);
    lines.forEach((line, i) => {
        let lineHtml = '';
        try {
            if (!line) {
                lineHtml = '&nbsp;';
            } else if (typeof hljs !== 'undefined') {
                // A. 调用 highlight.js 生成带有 class 的 HTML
                const res = (lang && lang !== 'auto') 
                    ? hljs.highlight(line, {language: lang, ignoreIllegals:true}) 
                    : hljs.highlightAuto(line);
                let rawHtml = res.value;

                // B. 【核心步骤】正则替换：把 class="hljs-xxx" 变成 style="..."
                lineHtml = rawHtml.replace(/<span class="hljs-([^"]+)">/g, (match, cls) => {
                    // cls 可能是 "keyword" 或 "keyword language-python" 等，只取第一个词
                    const key = cls.split(' ')[0]; 
                    const style = currentSyntax[key] || '';
                    return style ? `<span style="${style}">` : `<span>`; // 如果有对应颜色就替换，否则保持原样
                });

            } else {
                // 降级处理
                lineHtml = line.replace(/&/g, "&amp;").replace(/</g, "&lt;");
            }
        } catch(e) { 
            lineHtml = line.replace(/&/g, "&amp;").replace(/</g, "&lt;"); 
        }

        // 边框逻辑
        let cellBorder = `border-left:${border_style}; border-right:${border_style};`;
        if (i === 0) cellBorder += `border-top:${border_style};`;
        if (i === lines.length - 1) cellBorder += `border-bottom:${border_style};`;

        // 拼接 (紧凑模式)
        html += `<tr><td style="${style_num}">${i + 1}</td><td style="${style_code} ${cellBorder}">${lineHtml}</td></tr>`;
    });

    html += "</table>";
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