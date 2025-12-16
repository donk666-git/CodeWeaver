from flask import Flask, render_template, request, jsonify, send_from_directory
from pygments import highlight
from pygments.lexers import get_lexer_by_name, guess_lexer
from pygments.formatters import HtmlFormatter
import db
import os
import time
import requests

app = Flask(__name__)

SILCON_API_KEY = os.getenv("SILCON_API_KEY")

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

        api_key = os.environ.get('SILCON_API_KEY')
        if not api_key:
            return jsonify({'status': 'error', 'message': 'API key 未配置 (process.env.SILCON_API_KEY)'}), 400

        prompt = f"""
请作为资深工程师，用中文输出结构化 Markdown 教程来解读这段{language or '代码'}。请严格按照以下章节组织内容，保持简洁且可复用：
# 标题
- 用一句话概括代码目的。

## Overview
- 用2~3句话说明整体思路与关键组件。

## Step-by-step explanation
1. 按执行顺序逐步说明核心逻辑。
2. 点出关键函数/变量的职责与数据流。
3. 强调边界条件、错误处理或性能考量。

## Key points
- 列出易错点、复杂度/性能、健壮性或安全注意事项。

## Example
- 如适用，提供简单示例、调用方式或运行提示。

待讲解的代码如下：
```
{code}
```
""".strip()
        payload = {
            "model": "deepseek-ai/DeepSeek-V3.2",
            "messages": [
                {"role": "system", "content": "You are a senior engineer who writes concise, accurate, structured Chinese Markdown explanations with sections: Title, Overview, Step-by-step explanation, Key points, Example."},
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
            print("SiliconFlow status:", resp.status_code)
            print("SiliconFlow response:", resp.text)
            return jsonify({
                'status': 'error',
                'message': resp.text
            }), 502


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
