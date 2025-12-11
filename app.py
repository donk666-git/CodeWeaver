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
