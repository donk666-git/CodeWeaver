/* static/js/taskpane.js */

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        $(document).ready(function () {
            // 绑定按钮事件
            $('#btnInsert').click(insertHighlight);
            $('#btnSave').click(saveSnippet);
            $('#btnGetSelection').click(getFromSelection);
            
            // 搜索框简单过滤
            $('#searchBox').on('keyup', function() {
                var value = $(this).val().toLowerCase();
                $("#snippetList div").filter(function() {
                    $(this).toggle($(this).text().toLowerCase().indexOf(value) > -1)
                });
            });
        });
    }
});

async function getFromSelection() {
    try {
        await Word.run(async (context) => {
            const range = context.document.getSelection();
            range.load("text");
            await context.sync();
            if (range.text) $('#codeSource').val(range.text); // 不去除前后空格，保留原样
        });
    } catch (e) { console.error(e); }
}

async function insertHighlight() {
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
            alert("渲染错误: " + data.message);
        }
    } catch (error) {
        alert("连接错误: " + error.message);
    } finally {
        $('#btnInsert').prop('disabled', false).text('⚡ 插入高亮代码');
    }
}

// 【修复】保存功能
async function saveSnippet() {
    // 【新增调试代码】点击按钮时先弹个窗，证明 JS 跑通了
    console.log("正在尝试保存..."); 
    alert("我是弹窗：你点击了保存按钮！");
    
    const code = $('#codeSource').val();
    if (!code) return alert("代码是空的");

    const title = prompt("请输入标题:");
    if (!title) return;

    try {
        const res = await fetch('/api/snippets', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ 
                title: title, 
                code: code, 
                language: $('#langSelect').val() 
            })
        });
        const data = await res.json();
        
        if (data.status === 'success') {
            alert("✅ 保存成功！");
            // 自动跳转到代码库 Tab 并刷新
            $('#library-tab').tab('show'); 
            loadSnippets();
        } else {
            alert("保存失败");
        }
    } catch (error) {
        alert("保存错误: " + error.message);
    }
}

async function loadSnippets() {
    try {
        $('#snippetList').html('<div class="text-center mt-3"><div class="spinner-border text-primary spinner-border-sm"></div></div>');
        
        const response = await fetch('/api/snippets?t=' + new Date().getTime());
        const snippets = await response.json();
        
        console.log("从后台获取到的数据:", snippets); // 打开 F12 控制台可以看到数据

        const $list = $('#snippetList');
        $list.empty();

        if (snippets.length === 0) {
            $list.append('<div class="text-center text-muted mt-3">暂无代码，快去保存一条吧</div>');
            return;
        }

        snippets.forEach(item => {
            const $item = $(`
                <div class="list-group-item list-group-item-action p-2">
                    <div class="d-flex w-100 justify-content-between">
                        <strong class="mb-1 text-truncate" style="max-width: 150px;">${item.title}</strong>
                        <small class="text-primary">${item.language}</small>
                    </div>
                    <small class="text-muted d-block text-truncate" style="font-family:monospace;">${item.code.substring(0, 30)}...</small>
                </div>
            `);

            // 点击加载回编辑器
            $item.click(() => {
                $('#codeSource').val(item.code);
                $('#langSelect').val(item.language);
                $('#editor-tab').tab('show'); // 跳回编辑器 Tab
                var firstTabEl = document.querySelector('#myTab button[data-bs-target="#editor"]')
                var firstTab = new bootstrap.Tab(firstTabEl)
                firstTab.show()
            });

            $list.append($item);
        });

    } catch (error) {
       console.error("加载失败详情:", error);
        $('#snippetList').html('<div class="text-danger p-3">无法加载代码库<br><small>' + error.message + '</small></div>');
    }
}