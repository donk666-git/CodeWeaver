from flask import Flask, render_template, request, jsonify, send_from_directory
import db
import os
import time

app = Flask(__name__)

@app.route('/taskpane.html')
def taskpane():
    return render_template('taskpane.html')


@app.route('/highlight/<path:filename>')
def highlight_assets(filename):
    """Serve locally bundled highlight.js assets for offline stability."""
    return send_from_directory('highlight', filename)


def _summarize_code(code: str, language: str) -> str:
    """Generate a lightweight, local explanation for the provided code."""
    lines = [ln.rstrip() for ln in code.splitlines() if ln.strip()]
    line_count = len(lines)
    max_preview = 3
    preview = '\n'.join(lines[:max_preview]) if lines else ''

    keywords = {
        'python': ['def ', 'class ', 'import ', 'async ', 'yield '],
        'javascript': ['function ', 'const ', 'let ', 'class ', '=>'],
        'java': ['class ', 'public ', 'private ', 'static ', 'void '],
        'c': ['#include', 'int main', 'printf', 'struct '],
        'cpp': ['#include', 'std::', 'template', 'namespace'],
        'csharp': ['using ', 'namespace ', 'class ', 'public ', 'Task<'],
        'go': ['package ', 'func ', 'import ', 'defer '],
        'rust': ['fn ', 'let ', 'impl ', 'pub ', 'trait '],
    }

    detected = []
    lang_key = language.lower() if language else 'auto'
    for key, markers in keywords.items():
        if key == lang_key or lang_key == 'auto':
            for marker in markers:
                if any(marker in ln for ln in lines):
                    detected.append(key)
                    break

    lang_hint = lang_key if lang_key != 'auto' else (detected[0] if detected else '未知语言')
    headline = f"共 {line_count} 行代码，推测语言：{lang_hint}。"
    details = []

    def count_contains(substr: str) -> int:
        return sum(1 for ln in lines if substr in ln)

    function_like = count_contains('def ') + count_contains('function ') + count_contains('func ')
    if function_like:
        details.append(f"检测到 {function_like} 个函数/方法定义。")

    class_like = count_contains('class ')
    if class_like:
        details.append(f"包含 {class_like} 个类定义。")

    import_like = count_contains('import ') + count_contains('#include') + count_contains('using ')
    if import_like:
        details.append(f"有 {import_like} 处依赖导入。")

    if preview:
        details.append(f"前几行示例：\n{preview}")

    return headline + (" " + ' '.join(details) if details else '')


@app.route('/api/snippets', methods=['POST'])
def save_snippet():
    try:
        data = request.json
        project = data.get('project', 'Default')
        title = data.get('title', 'Untitled')
        code = data.get('code', '')
        language = data.get('language', 'auto')
        snippet_id = data.get('id')

        if snippet_id:
            success, new_id = db.update_snippet(snippet_id, project, title, code, language)
            mode = 'update'
        else:
            success, new_id = db.save_snippet_v2(project, title, code, language)
            mode = 'create'

        if success:
            return jsonify({'status': 'success', 'mode': mode, 'id': new_id})
        return jsonify({'status': 'error'}), 500
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500


@app.route('/api/explain', methods=['POST'])
def explain_snippet():
    try:
        data = request.json or {}
        code = data.get('code', '')
        language = data.get('language', 'auto')
        if not code.strip():
            return jsonify({'status': 'error', 'message': 'empty code'}), 400
        explanation = _summarize_code(code, language)
        return jsonify({'status': 'success', 'explanation': explanation})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

@app.route('/api/snippets', methods=['GET'])
def get_snippets():
    keyword = request.args.get('q')
    return jsonify(db.get_all_grouped(keyword))

@app.route('/api/snippets/<int:id>', methods=['DELETE'])
def delete_snippet(id):
    db.delete_snippet(id)
    return jsonify({'status': 'success'})

@app.route('/api/projects/delete', methods=['POST'])
def delete_project():
    try:
        name = request.json.get('name')
        if db.delete_project(name):
            return jsonify({'status': 'success'})
        return jsonify({'status': 'error', 'message': 'Not found'}), 404
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True, port=5000, ssl_context=('ssl/server.crt', 'ssl/server.key'))
