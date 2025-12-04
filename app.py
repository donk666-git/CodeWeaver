from flask import Flask, render_template, request, jsonify
from pygments import highlight
from pygments.lexers import get_lexer_by_name, guess_lexer
from pygments.formatters import HtmlFormatter
import db
import os

app = Flask(__name__)

@app.route('/taskpane.html')
def taskpane():
    return render_template('taskpane.html')

@app.route('/api/render', methods=['POST'])
def render_code():
    try:
        data = request.json
        code = data.get('code', '')
        lang = data.get('language', 'auto')
        if not code: return jsonify({'status': 'error', 'message': 'Empty'}), 400

        try:
            lexer = guess_lexer(code) if lang == 'auto' else get_lexer_by_name(lang)
        except:
            lexer = get_lexer_by_name('text')

        formatter = HtmlFormatter(style='vs', nowrap=True, noclasses=True)
        lines = code.splitlines()
        font_style = "font-family: Consolas, Monaco, 'Courier New', monospace; mso-font-alt: 'Courier New';"
        
        table_html = f'<table style="width:100%; border-collapse:collapse; border:1px solid #d0d7de; background-color:#f6f8fa; font-size:10pt; mso-no-proof:yes; {font_style}">'
        for index, line in enumerate(lines):
            line_num = index + 1
            leading_spaces = len(line) - len(line.lstrip(' '))
            content = line.strip()
            hl_line = '&nbsp;'
            if content:
                hl_line = highlight(content, lexer, formatter).replace('<span style="', f'<span style="{font_style} ')
            final_html = ('&nbsp;' * leading_spaces) + hl_line
            table_html += f'<tr style="height:auto;"><td style="width:35px; background:#eef1f5; border-right:1px solid #d0d7de; text-align:right; padding:2px 8px 2px 0; color:#999; font-size:9pt; vertical-align:top; user-select:none; {font_style}">{line_num}</td><td style="padding:2px 0 2px 10px; vertical-align:top; background:#f6f8fa; color:#24292f; {font_style}">{final_html}</td></tr>'
        table_html += "</table>"
        
        return jsonify({'status': 'success', 'html': table_html})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

@app.route('/api/snippets', methods=['POST'])
def save_snippet():
    try:
        data = request.json
        success = db.save_snippet_v2(
            data.get('project', 'Default'), 
            data.get('title', 'Untitled'), 
            data.get('code', ''), 
            data.get('language', 'auto')
        )
        return jsonify({'status': 'success'}) if success else jsonify({'status': 'error'}), 500
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

@app.route('/api/snippets', methods=['GET'])
def get_snippets():
    return jsonify(db.get_all_grouped())

@app.route('/api/snippets/<int:id>', methods=['DELETE'])
def delete_snippet(id):
    db.delete_snippet(id)
    return jsonify({'status': 'success'})

# 【新增】删除项目接口
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