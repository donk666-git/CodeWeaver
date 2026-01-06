// static/js/api.js - backend & AI calls

export async function loadSnippets(keyword = '') {
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
