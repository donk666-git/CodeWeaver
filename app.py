from flask import Flask, render_template, request, jsonify, send_from_directory
from pygments import highlight
from pygments.lexers import get_lexer_by_name, guess_lexer
from pygments.formatters import HtmlFormatter
import db
import os
import time

app = Flask(__name__)

@app.route('/taskpane.html')
def taskpane():
    return render_template('taskpane.html')


@app.route('/highlight/<path:filename>')
def highlight_assets(filename):
    return send_from_directory('highlight', filename)


@app.route('/api/snippets', methods=['POST'])
def save_snippet():
    try:
        data = request.json
        success = db.save_snippet_v2(
            data.get('project', 'Default'),
            data.get('title', 'Untitled'),
            data.get('code', ''),
            data.get('language', 'auto'),
            data.get('style_config'),
            data.get('id')
        )
        return jsonify({'status': 'success'}) if success else jsonify({'status': 'error'}), 500
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

@app.route('/api/snippets', methods=['GET'])
def get_snippets():
    keyword = request.args.get('q', '').strip()
    return jsonify(db.get_all_grouped(keyword if keyword else None))

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
