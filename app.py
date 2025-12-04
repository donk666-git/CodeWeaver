# app.py
from flask import Flask, render_template, request, jsonify
from pygments import highlight
from pygments.lexers import get_lexer_by_name, guess_lexer
from pygments.formatters import HtmlFormatter
import db # 引入我们要写的数据库模块
import os

app = Flask(__name__)

# 1. 路由：插件的主页 (Frontend)
@app.route('/taskpane.html')
def taskpane():
    return render_template('taskpane.html')

# 2. API：代码高亮渲染 (核心功能)
# app.py 中的 render_code 函数替换为：

@app.route('/api/render', methods=['POST'])
def render_code():
    data = request.json
    code = data.get('code', '')
    lang = data.get('language', 'auto')
    
    try:
        # 1. 确定语言 (优化自动检测)
        if lang == 'auto':
            try:
                lexer = guess_lexer(code)
            except:
                lexer = get_lexer_by_name('text') # 猜不到就用纯文本
        else:
            lexer = get_lexer_by_name(lang)
            
        # 2. 获取 Formatter
        # 使用 'default' 风格，颜色更鲜艳
        formatter = HtmlFormatter(style='default', nowrap=True, noclasses=True)
        
        lines = code.splitlines()
        
        # 3. 定义 Word 专用样式
        # mso-no-proof: yes -> 禁止 Word 对代码进行拼写检查（去掉红色波浪线）
        font_style = "font-family: Consolas, Monaco, 'Courier New', monospace; mso-font-alt: 'Courier New';"
        
        # 构建表格
        table_html = f'''
        <table style="width:100%; border-collapse:collapse; border:1px solid #d0d7de; background-color:#f6f8fa; font-size:10pt; mso-no-proof: yes; {font_style}">
        '''
        
        for index, line in enumerate(lines):
            line_num = index + 1
            
            # 【核心修复 1】解决缩进消失
            # 逻辑：先计算前面的空格数，然后把它们变成 HTML 的 &nbsp;
            # Word 只有这样才能完美保留缩进
            leading_spaces = len(line) - len(line.lstrip(' '))
            content = line.strip()
            
            if not content: # 空行处理
                highlighted_line = '&nbsp;'
            else:
                # 高亮代码内容
                highlighted_line = highlight(content, lexer, formatter)
                # 将原来的 span 样式再次加强，防止被 Word 覆盖
                highlighted_line = highlighted_line.replace('<span style="', f'<span style="{font_style} ')
            
            # 补回缩进 (用 &nbsp; 替换空格)
            indent_html = '&nbsp;' * leading_spaces
            final_html = indent_html + highlighted_line
            
            table_html += f'''
            <tr style="height:auto;">
                <td style="
                    width: 35px; 
                    background-color: #eef1f5; 
                    border-right: 1px solid #d0d7de; 
                    text-align: right; 
                    padding: 2px 8px 2px 0; 
                    color: #999999; 
                    font-size: 9pt; 
                    vertical-align: top; 
                    user-select: none;
                    {font_style}">
                    {line_num}
                </td>
                
                <td style="
                    padding: 2px 0 2px 10px; 
                    vertical-align: top; 
                    background-color: #f6f8fa;
                    color: #24292f;
                    {font_style}">
                    {final_html}
                </td>
            </tr>
            '''
            
        table_html += "</table>"
        return jsonify({'status': 'success', 'html': table_html})
        
    except Exception as e:
        print(f"Render Error: {e}")
        return jsonify({'status': 'error', 'message': str(e)}), 400
# 3. API：保存代码片段 (CRUD - Create)
@app.route('/api/snippets', methods=['POST'])
def save_snippet():
    data = request.json
    db.add_snippet(data['title'], data['code'], data['language'])
    return jsonify({'status': 'success'})

# 4. API：获取所有片段 (CRUD - Read)
@app.route('/api/snippets', methods=['GET'])
def get_snippets():
    snippets = db.get_all_snippets()
    return jsonify(snippets)

if __name__ == '__main__':
    # 修改前：ssl_context='adhoc'
    # 修改后：指定刚才生成的固定证书路径
    app.run(debug=True, port=5000, ssl_context=('ssl/server.crt', 'ssl/server.key'))