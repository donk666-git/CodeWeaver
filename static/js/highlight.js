// static/js/highlight.js - syntax highlighting logic

let hljsConfigured = false;

export function ensureHighlighter() {
    if (typeof hljs === 'undefined') return;
    if (!hljsConfigured) {
        hljs.configure({ ignoreUnescapedHTML: true });
        hljsConfigured = true;
    }
}

export function buildLanguageDropdown() {
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

export function normalizeIndentationText(raw, language = '') {
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

// Generate highlighted HTML
export function generateHighlightHtml(code, lang, theme, listingNo) {
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
