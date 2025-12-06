# debug_test.py
from pygments import highlight
from pygments.lexers import get_lexer_by_name
from pygments.formatters import HtmlFormatter

# 1. 模拟一段 JSON 代码
code = '{\n  "name": "CodeWeaver",\n  "version": 1\n}'

try:
    # 2. 模拟你在 app.py 里的选择逻辑
    print("--- 正在加载 Lexer: json ---")
    lexer = get_lexer_by_name("json") 
    print(f"成功加载: {lexer.name}")

    # 3. 模拟你的 Formatter 配置
    # style='vs' 是你默认的主题
    formatter = HtmlFormatter(style='vs', nowrap=True, noclasses=True)

    # 4. 执行高亮
    result = highlight(code, lexer, formatter)
    
    print("\n--- 生成的 HTML 结果 ---")
    print(result)
    
    # 5. 检查关键点
    if 'style="color' in result:
        print("\n✅ 成功：结果中包含颜色样式 (style='color...')")
    else:
        print("\n❌ 失败：结果中没有颜色样式 (全黑)")
        
except Exception as e:
    print(f"\n❌ 报错了: {e}")