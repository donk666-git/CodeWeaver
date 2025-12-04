/* static/js/taskpane.js v4.0 */

// 全局变量
let deleteTarget = null; 
let confirmModal = null;

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        $(document).ready(function () {
            console.log("✅ CodeWeaver v4.0 (Event Delegation) Ready");
            
            // 1. 初始化
            syncProjectName();
            loadSnippets();
            confirmModal = new bootstrap.Modal(document.getElementById('confirmModal'));

            // 2. 绑定静态按钮 (编辑器页)
            $('#btnSave').click(saveSnippet);
            $('#btnInsert').click(insertHighlight);
            $('#btnGetSelection').click(getFromSelection);
            
            // 3. 绑定静态按钮 (项目库页)
            $('#btnRefresh').click(loadSnippets);
            $('#library-tab').click(loadSnippets); // 点击 Tab 也刷新

            // 4. 【核心】绑定动态列表按钮 (事件委托)
            // 这种写法确保即使是新加载出来的 HTML，点击也没问题
            
            // A. 点击标题 -> 加载到编辑器
            $(document).on('click', '.action-load-editor', function() {
                const code = decodeURIComponent($(this).data('code'));
                const lang = $(this).data('lang');
                $('#codeSource').val(code);
                $('#langSelect').val(lang);
                new bootstrap.Tab('#editor-tab').show();
            });

            // B. 点击定位 -> 在文档中搜索
            $(document).on('click', '.action-locate', function() {
                const code = decodeURIComponent($(this).data('code'));
                locateInDoc(code);
            });

            // C. 点击删除代码 -> 弹窗
            $(document).on('click', '.action-del-snippet', function() {
                const id = $(this).data('id');
                const title = $(this).data('title');
                askDeleteSnippet(id, title);
            });

            // D. 点击删除项目 -> 弹窗
            $(document).on('click', '.action-del-project', function() {
                const name = $(this).data('name'); // 注意这里取的是 data-name
                askDeleteProject(name);
            });

            // E. 确认删除按钮
            $('#btnConfirmDelete').click(performDelete);

            // 5. 搜索功能
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

// 1. 保存
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

// 2. 加载列表 (生成 data-* 属性)
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
            // HTML 安全处理：把项目名放到 data-name 里
            // 注意：这里我们不需要自己拼 onclick 字符串了，所以引号问题好解决多了
            
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

// 3. 删除逻辑 (弹窗)
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
            loadSnippets(); // 刷新
        } else { alert("删除失败"); }
    } catch (e) { alert("网络错误"); }
}

// 4. 其他逻辑
async function insertHighlight() {
    const code = $('#codeSource').val();
    const lang = $('#langSelect').val();
    if (!code) return showStatus("❌ 代码为空", "error");
    try {
        const res = await fetch('/api/render', {
            method: 'POST',
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify({code, language: lang})
        });
        const data = await res.json();
        if(data.status === 'success') {
            await Word.run(async (ctx) => {
                ctx.document.getSelection().insertHtml(data.html, "Replace");
                await ctx.sync();
            });
        } else showStatus("❌ 渲染失败", "error");
    } catch(e) {}
}

async function getFromSelection() {
    try {
        await Word.run(async (ctx) => {
            const r = ctx.document.getSelection();
            r.load("text");
            await ctx.sync();
            if(r.text) $('#codeSource').val(r.text);
        });
    } catch(e){}
}

async function locateInDoc(code) {
    const searchKey = code.substring(0, 50).trim();
    try {
        await Word.run(async (ctx) => {
            const r = ctx.document.body.search(searchKey, { matchCase: true });
            ctx.load(r);
            await ctx.sync();
            if (r.items.length > 0) {
                r.items[0].select();
                showStatus("✅ 已定位");
            } else {
                showStatus("⚠️ 未找到", "error");
            }
        });
    } catch(e){}
}