# db.py
import sqlite3
import os

# 【核心修复】获取当前 db.py 文件所在的绝对文件夹路径
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
# 拼接出数据库的绝对路径，确保不管你在哪里运行 python，都读写同一个文件
DB_FILE = os.path.join(BASE_DIR, 'code_weaver.db')

def get_connection():
    # check_same_thread=False 允许 Flask 多线程访问
    return sqlite3.connect(DB_FILE, check_same_thread=False)

def init_db():
    conn = get_connection()
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS snippets (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT,
            code TEXT,
            language TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    conn.commit()
    conn.close()
    print(f"✅ 数据库已连接: {DB_FILE}") # 打印路径方便你检查

def add_snippet(title, code, language):
    try:
        conn = get_connection()
        c = conn.cursor()
        c.execute('INSERT INTO snippets (title, code, language) VALUES (?, ?, ?)', 
                  (title, code, language))
        conn.commit()
        conn.close()
        print(f"💾 成功写入数据库: {title}") # 后台打印日志
        return True
    except Exception as e:
        print(f"❌ 写入失败: {e}")
        return False

def get_all_snippets():
    conn = get_connection()
    conn.row_factory = sqlite3.Row
    c = conn.cursor()
    c.execute('SELECT * FROM snippets ORDER BY id DESC')
    rows = c.fetchall()
    conn.close()
    # 将 row 对象转为字典，方便 Flask 序列化
    return [dict(row) for row in rows]

# 初始化
init_db()