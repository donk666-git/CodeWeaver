// static/js/explain.js - AI explanation UI

import { showStatus } from './state.js';

let explanationCollapsed = false;
let lastExplanationContent = ''; // 存储原始解释内容用于复制

export function setExplainVisibility(show) {
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

export function toggleExplainPanel() {
    setExplainVisibility(explanationCollapsed);
}

export function copyExplanation() {
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

export function renderExplanation(content, isRaw = false) {
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

export async function requestExplanation() {
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
