// static/js/ui.js - DOM, buttons, layout logic

import { loadSnippets } from './api.js';
import { applyIndentationNormalization, insertHighlight, getFromSelection, locateInDoc } from './editor.js';
import { requestExplanation, toggleExplainPanel, copyExplanation, setExplainVisibility } from './explain.js';
import { buildLanguageDropdown, ensureHighlighter } from './highlight.js';
import { renumberListings, showStatus } from './state.js';

let deleteTarget = null;
let confirmModal = null;
let currentEditingId = null;
let searchTimer = null;

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

export function initializeTaskpane() {
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
