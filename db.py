# db.py
import sqlite3
import os
import json

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_FILE = os.path.join(BASE_DIR, 'code_weaver.db')

def get_connection():
    return sqlite3.connect(DB_FILE, check_same_thread=False)

def init_db():
    conn = get_connection()
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS projects (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT UNIQUE,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    c.execute('''
        CREATE TABLE IF NOT EXISTS snippets (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id INTEGER,
            title TEXT,
            code TEXT,
            language TEXT,
            style_config TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(project_id) REFERENCES projects(id)
        )
    ''')
    conn.commit()
    conn.close()

def save_snippet_v2(project_name, title, code, language, style_config=None):
    conn = get_connection()
    c = conn.cursor()
    try:
        # 1. 查找或创建项目
        c.execute('SELECT id FROM projects WHERE name = ?', (project_name,))
        row = c.fetchone()
        if row:
            project_id = row[0]
        else:
            c.execute('INSERT INTO projects (name) VALUES (?)', (project_name,))
            project_id = c.lastrowid
            
        # 2. 插入代码
        config_str = json.dumps(style_config) if style_config else "{}"
        c.execute('''
            INSERT INTO snippets (project_id, title, code, language, style_config) 
            VALUES (?, ?, ?, ?, ?)
        ''', (project_id, title, code, language, config_str))
        conn.commit()
        return True
    except Exception as e:
        print(f"❌ DB Error: {e}")
        return False
    finally:
        conn.close()

def delete_snippet(snippet_id):
    conn = get_connection()
    c = conn.cursor()
    c.execute('DELETE FROM snippets WHERE id = ?', (snippet_id,))
    conn.commit()
    conn.close()

# 【新增】删除项目 (连带删除下面的代码)
def delete_project(project_name):
    conn = get_connection()
    c = conn.cursor()
    try:
        # 1. 先找 ID
        c.execute('SELECT id FROM projects WHERE name = ?', (project_name,))
        row = c.fetchone()
        if not row: return False
        pid = row[0]
        
        # 2. 删除该项目下的所有 snippet
        c.execute('DELETE FROM snippets WHERE project_id = ?', (pid,))
        
        # 3. 删除项目本身
        c.execute('DELETE FROM projects WHERE id = ?', (pid,))
        conn.commit()
        return True
    except Exception as e:
        print(e)
        return False
    finally:
        conn.close()

def get_all_grouped():
    conn = get_connection()
    conn.row_factory = sqlite3.Row
    c = conn.cursor()
    # 联表查询
    c.execute('''
        SELECT s.id, s.title, s.code, s.language, s.created_at, p.name as project_name
        FROM snippets s
        JOIN projects p ON s.project_id = p.id
        ORDER BY p.created_at DESC, s.created_at DESC
    ''')
    rows = c.fetchall()
    conn.close()
    
    result = {}
    for row in rows:
        item = dict(row)
        p_name = item['project_name']
        if p_name not in result:
            result[p_name] = []
        result[p_name].append(item)
    return result

init_db()