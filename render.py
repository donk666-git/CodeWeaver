@app.route('/api/render', methods=['POST'])
def render_code():
    try:
        data = request.json
        code = data.get('code', '')
        lang = data.get('language', 'auto').lower()
        theme = data.get('theme', 'gray') 
        
        if not code: return jsonify({'status': 'error', 'message': 'Empty'}), 400

        lexer_map = {
            'c': 'c', 'cpp': 'cpp', 'python': 'python', 'java': 'java',
            'html': 'html', 'sql': 'sql', 'matlab': 'matlab',
            'json': 'json', 'bash': 'bash', 'javascript': 'js'
        }

        try:
            if lang == 'auto':
                lexer = guess_lexer(code)
            else:
                alias = lexer_map.get(lang, lang)
                lexer = get_lexer_by_name(alias)
        except:
            lexer = get_lexer_by_name('text')

        # --- 布局参数（完全不动） ---
        offset_px = 34
        table_margin_left = f'-{offset_px}px'
        table_width_style = f'calc(100% + {offset_px}px)'
        num_col_padding_right = '4px'

        # --- 主题逻辑（完全不动） ---
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

        formatter = HtmlFormatter(style=pygments_style, nowrap=True, noclasses=True)
        lines = code.splitlines()
        
        # 【微改动 A】样式里加 white-space: pre 和 mso-no-proof: yes
        # mso-no-proof: yes 负责去红线
        # white-space: pre 负责保留缩进
        font_style_code = f"font-family: 'Courier New', Consolas, monospace; font-size: 10pt; white-space: pre; mso-no-proof: yes;"
        font_style_num = f"font-family: 'Times New Roman', serif; font-size: 6pt;" 
        
        table_html = (
            f'<table style="width:{table_width_style}; table-layout:fixed; border-collapse:collapse; border-spacing:0; '
            f'margin-bottom: 10px; margin-left: {table_margin_left}; '
            f'background-color:{bg_table_wrapper};">'
        )
        
        # 恢复原来的循环结构，确保边框样式和你原版一模一样
        for index, line in enumerate(lines):
            line_num = index + 1
            
            # 【微改动 B】
            # 1. 删掉 line.strip()，直接保留原始缩进
            # 2. 删掉 leading_spaces 计算，交给 CSS 处理
            
            content = line  # 不去空格
            
            hl_line = ''
            if content:
                # rstrip('\n\r') 防止 Pygments 加上多余换行
                hl_line = highlight(content, lexer, formatter).rstrip('\n\r')
            else:
                hl_line = '&nbsp;' # 空行保留占位符
            
            final_code = hl_line
            
            # --- 边框逻辑（完全不动，原样保留） ---
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

            cell_spacing_style = "padding-top: 0; padding-bottom: 0; line-height: 110%;" 

            table_html += (
                f'<tr>'
                f'<td style="width:30px; background-color:{bg_num_cell}; color:{text_color_num}; '
                f'text-align:right; padding-right:{num_col_padding_right}; border:none; '
                f'vertical-align:baseline; user-select:none; {cell_spacing_style} {font_style_num}">'
                f'{line_num}'
                f'</td>'
                
                # lang="zxx" 辅助屏蔽拼写检查
                f'<td lang="zxx" style="width:100%; background-color:{bg_code_cell}; color:{text_color_code}; '
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