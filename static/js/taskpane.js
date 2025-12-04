/* static/js/taskpane.js v=8888 */

// ==========================================
// 1. 初始化区域
// ==========================================
Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        $(document).ready(function () {
            console.log("✅ CodeWeaver v8888 已加载"); // 看控制台有没有这行

            // 加载列表
            loadSnippets();

            // 【双保险】: 万一 onclick 没反应，这里的绑定会生效
            $('#btnSave').off('click').on('click', function(e) {
                console.log("JQuery click triggered");
                // 如果 HTML onclick 已经触发了，这里可能会触发第二次，但总比不触发好
                // 我们可以检查一下 event
            });
            
            // 绑定获取选中
            $('#btnGetSelection').click(getFromSelection);

            // 搜索框逻辑
            $('#searchBox').on('keyup', function() {
                var value = $(this).val().toLowerCase();
                $("#snippetList > button").filter(function() {
                    $(this).toggle($(this).text().toLowerCase().indexOf(value) > -1)
                });
            });
        });
    }
});

// ==========================================
// 2. 核心功能函数 (挂载到 window 确保全局可见)
// ==========================================

// 保存函数
window.saveSnippet = async function() {
    console.log("🚀 saveSnippet 被调用了！");
    alert("1. 按钮点击成功！开始保存...");

    const code = $('#codeSource').val();
    if (!code) {
        alert("⚠️ 代码框是空的");
        return;
    }

    const title = "自动保存-" + new Date().toLocaleTimeString();
    const lang = $('#langSelect').val() || 'auto';

    try {
        const res = await fetch('/api/snippets', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ title, code, language: lang })
        });
        
        const data = await res.json();
        if (data.status === 'success') {
            alert("✅ 数据库保存成功！");
            // 刷新列表
            window.loadSnippets();
            // 尝试切换 Tab
            try {
                var triggerEl = document.querySelector('#library-tab')
                var tab = new bootstrap.Tab(triggerEl)
                tab.show()
            } catch(e) { console.log(e); }
        } else {
            alert("❌ 保存失败: " + JSON.stringify(data));
        }
    } catch (error) {
        alert("❌ 网络请求错误: " + error.message);
    }
};

// 插入高亮函数
window.insertHighlight = async function() {
    const code = $('#codeSource').val();
    const lang = $('#langSelect').val();
    if (!code) return alert("请输入代码");

    $('#btnInsert').prop('disabled', true).text('处理中...');

    try {
        const response = await fetch('/api/render', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ code: code, language: lang })
        });
        const data = await response.json();

        if (data.status === 'success') {
            await Word.run(async (context) => {
                const range = context.document.getSelection();
                range.insertHtml(data.html, Word.InsertLocation.Replace);
                await context.sync();
            });
        } else {
            alert("渲染失败: " + data.message);
        }
    } catch (error) {
        alert("错误: " + error.message);
    } finally {
        $('#btnInsert').prop('disabled', false).text('⚡ 插入高亮代码');
    }
};

// 加载列表函数
window.loadSnippets = async function() {
    try {
        // 时间戳防缓存
        const response = await fetch('/api/snippets?t=' + new Date().getTime());
        const snippets = await response.json();
        
        const $list = $('#snippetList');
        $list.empty();

        if (!snippets || snippets.length === 0) {
            $list.append('<div class="text-center text-muted mt-3">暂无代码</div>');
            return;
        }

        snippets.forEach(item => {
            const $item = $(`
                <button type="button" class="list-group-item list-group-item-action text-start">
                    <div class="d-flex w-100 justify-content-between">
                        <strong>${item.title}</strong>
                        <small>${item.language}</small>
                    </div>
                </button>
            `);
            $item.click(() => {
                $('#codeSource').val(item.code);
                // 切回编辑器
                var triggerEl = document.querySelector('#editor-tab')
                var tab = new bootstrap.Tab(triggerEl)
                tab.show()
            });
            $list.append($item);
        });
    } catch (error) {
        console.error(error);
        $('#snippetList').html('<div class="text-danger text-center">加载失败</div>');
    }
};

// 获取选中
window.getFromSelection = async function() {
    try {
        await Word.run(async (context) => {
            const range = context.document.getSelection();
            range.load("text");
            await context.sync();
            if (range.text) $('#codeSource').val(range.text);
        });
    } catch (e) { console.error(e); }
};