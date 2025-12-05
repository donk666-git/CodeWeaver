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
        theme = data.get('theme', 'gray') 
        
        if not code: return jsonify({'status': 'error', 'message': 'Empty'}), 400

        try:
            lexer = guess_lexer(code) if lang == 'auto' else get_lexer_by_name(lang)
        except:
            lexer = get_lexer_by_name('text')

        # --- 1. 全局布局计算 ---
        offset_px = 34
        table_margin_left = f'-{offset_px}px'
        table_width_style = f'calc(100% + {offset_px}px)'
        num_col_padding_right = '4px'

        # --- 2. 主题配色系统 ---
        pygments_style = 'vs'
        bg_table_wrapper = '#ffffff' 
        bg_code_cell = '#f6f8fa'     
        bg_num_cell = '#ffffff'      
        text_color_code = '#24292f'  
        text_color_num = '#6e7781'   
        border_color = '#d0d7de'     
        enable_radius = True
        
        if theme == 'dark':
            pygments_style = 'monokai'
            bg_table_wrapper = '#ffffff' 
            bg_code_cell = '#272822'     
            bg_num_cell = '#ffffff'      
            text_color_code = '#f8f8f2'
            text_color_num = '#6e7781'   
            border_color = '#272822' 
            enable_radius = True
            
        elif theme == 'green':
            pygments_style = 'default'
            bg_table_wrapper = '#ffffff'
            bg_code_cell = '#e9f5e9'
            bg_num_cell = '#ffffff'
            text_color_code = '#000000'
            text_color_num = '#999999'
            border_color = '#e9f5e9'
            enable_radius = False

        # --- 3. 代码高亮处理 ---
        formatter = HtmlFormatter(style=pygments_style, nowrap=True, noclasses=True)
        lines = code.splitlines()
        
        # 字体定义
        font_family = "Consolas, Monaco, 'Courier New', monospace, 'SimSun'"
        font_style_code = f"font-family: 'Courier New', Consolas, monospace; font-size: 10pt;"
        font_style_num = f"font-family: 'Times New Roman', serif; font-size: 6pt;" 
        
        # --- 4. 构造 HTML 表格 ---
        table_html = (
            f'<table style="width:{table_width_style}; table-layout:fixed; border-collapse:collapse; border-spacing:0; '
            f'margin-bottom: 10px; margin-left: {table_margin_left}; '
            f'background-color:{bg_table_wrapper};">'
        )
        
        for index, line in enumerate(lines):
            line_num = index + 1
            leading_spaces = len(line) - len(line.lstrip(' '))
            content = line.strip()
            
            # 高亮生成
            hl_line = ''
            if content:
                hl_line = highlight(content, lexer, formatter)
                hl_line = hl_line.replace('<span style="', f'<span style="{font_style_code} ')
            else:
                hl_line = '&nbsp;'
            
            final_code = ('&nbsp;' * leading_spaces) + hl_line
            
            # --- 构造边框和圆角 ---
            border_width = "1.5pt"
            
            cell_style_extra = (
                f"border-left: {border_width} solid {border_color}; "
                f"border-right: {border_width} solid {border_color}; "
            )
            
            if index == 0: 
                cell_style_extra += f"border-top: {border_width} solid {border_color}; "
                if enable_radius:
                    cell_style_extra += "border-top-left-radius: 6px; border-top-right-radius: 6px;"
            
            if index == len(lines) - 1:
                cell_style_extra += f"border-bottom: {border_width} solid {border_color}; "
                if enable_radius:
                    cell_style_extra += "border-bottom-left-radius: 6px; border-bottom-right-radius: 6px;"

            # ==========================================================
            # 【重要修复】行间距设置
            # 移除了 mso-line-height-rule: exactly; 
            # 改用百分比 (120%)。
            # 想要更紧凑？ 改成 110%
            # 想要更宽松？ 改成 130% 或 140%
            # ==========================================================
            cell_spacing_style = "padding-top: 0; padding-bottom: 0; line-height: 110%;" 

            table_html += (
                f'<tr>'
                # --- 左侧：行号列 ---
                f'<td style="width:30px; background-color:{bg_num_cell}; color:{text_color_num}; '
                f'text-align:right; padding-right:{num_col_padding_right}; border:none; '
                f'vertical-align:baseline; user-select:none; {cell_spacing_style} {font_style_num}">'
                f'{line_num}'
                f'</td>'
                
                # --- 右侧：代码列 ---
                f'<td style="width:100%; background-color:{bg_code_cell}; color:{text_color_code}; '
                f'padding-left:10px; padding-right:10px; '
                f'vertical-align:baseline; {cell_style_extra} {cell_spacing_style} {font_style_code}">'
                f'{final_code}'
                f'</td>'
                f'</tr>'
            )
            
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