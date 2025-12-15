from flask import Flask, render_template, request, jsonify, send_from_directory
from pygments import highlight
from pygments.lexers import get_lexer_by_name, guess_lexer
from pygments.formatters import HtmlFormatter
import db
import os
import time
import requests
import json

app = Flask(__name__)
app.config['DEEPSEEK_API_KEY'] = os.environ.get('DEEPSEEK_API_KEY')


def load_deepseek_api_key():
    env_key = (os.environ.get('DEEPSEEK_API_KEY') or app.config.get('DEEPSEEK_API_KEY') or '').strip()
    if env_key:
        return env_key

    config_path = os.path.join(os.path.dirname(__file__), 'config.json')
    if os.path.exists(config_path):
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                cfg = json.load(f) or {}
                file_key = cfg.get('DEEPSEEK_API_KEY') or cfg.get('deepseek_api_key') or cfg.get('deepseekApiKey')
                if file_key:
                    return str(file_key).strip()
        except Exception:
            pass
    return None

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


@app.route('/api/explain', methods=['POST'])
def explain_code():
    try:
        data = request.json or {}
        code = (data.get('code') or '').strip()
        language = data.get('language') or ''
        if not code:
            return jsonify({'status': 'error', 'message': '缺少代码内容'}), 400

        api_key = load_deepseek_api_key()
        if not api_key:
            return jsonify({'status': 'error', 'message': 'API key 未配置 (环境变量/配置文件 DEEPSEEK_API_KEY)'}), 400

        prompt = f"请用中文解释这段{language or '代码'}，突出核心逻辑：\n```\n{code}\n```"
        payload = {
            "model": "deepseek-chat",
            "messages": [
                {"role": "system", "content": "You are a senior engineer providing concise, accurate explanations."},
                {"role": "user", "content": prompt}
            ],
            "temperature": 0.2
        }

        resp = requests.post(
            'https://api.siliconflow.cn/v1/chat/completions',
            headers={
                'Authorization': f'Bearer {api_key}',
                'Content-Type': 'application/json'
            },
            json=payload,
            timeout=20
        )

        if resp.status_code >= 400:
            return jsonify({'status': 'error', 'message': '外部接口错误'}), 502

        result = resp.json()
        explanation = ''
        try:
            explanation = result.get('choices', [{}])[0].get('message', {}).get('content', '').strip()
        except Exception:
            explanation = ''

        if not explanation:
            explanation = '未获取到有效解读'

        return jsonify({'status': 'success', 'explanation': explanation})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True, port=5000, ssl_context=('ssl/server.crt', 'ssl/server.key'))
