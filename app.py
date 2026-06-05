import os
import sqlite3
from datetime import datetime, timedelta
from functools import wraps
from flask import Flask, request, jsonify, session, send_file, send_from_directory
from flask_cors import CORS
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
from openpyxl import Workbook
from io import BytesIO

app = Flask(__name__)
app.secret_key = "super_secret_reimbursement_system_key_12345"
CORS(app, supports_credentials=True)

DB_PATH = 'expense_system.db'
UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB limit

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

# ============ 數據庫初始化 ============
def init_db():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    # 創建角色表
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS Roles (
        role_id INTEGER PRIMARY KEY AUTOINCREMENT,
        role_name TEXT UNIQUE,
        default_quota REAL
    )
    """)
    
    # 創建使用者表
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS Users (
        user_id INTEGER PRIMARY KEY AUTOINCREMENT,
        last_name TEXT,
        first_name TEXT,
        gender TEXT,
        email TEXT UNIQUE,
        password_hash TEXT,
        role_id INTEGER,
        current_quota REAL,
        login_failures INTEGER DEFAULT 0,
        login_count INTEGER DEFAULT 0,
        status TEXT,
        FOREIGN KEY(role_id) REFERENCES Roles(role_id)
    )
    """)
    
    # 創建 IP 登入紀錄表
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS IP_Log (
        log_id INTEGER PRIMARY KEY AUTOINCREMENT,
        ip_address TEXT,
        user_id INTEGER,
        is_success INTEGER,
        failed_reason TEXT,
        login_time TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY(user_id) REFERENCES Users(user_id)
    )
    """)
    
    # 創建部門表
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS Departments (
        dept_id INTEGER PRIMARY KEY AUTOINCREMENT,
        dept_name TEXT UNIQUE
    )
    """)
    
    # 創建費用類別表
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS Categories (
        category_id INTEGER PRIMARY KEY AUTOINCREMENT,
        category_name TEXT UNIQUE
    )
    """)
    
    # 創建報帳申請主檔
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS Expense_Requests (
        request_id INTEGER PRIMARY KEY AUTOINCREMENT,
        applicant_id INTEGER,
        amount REAL,
        expense_date TEXT,
        category_id INTEGER,
        dept_id INTEGER,
        status TEXT,
        description TEXT,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY(applicant_id) REFERENCES Users(user_id),
        FOREIGN KEY(category_id) REFERENCES Categories(category_id),
        FOREIGN KEY(dept_id) REFERENCES Departments(dept_id)
    )
    """)
    
    # 創建證明文件附件表
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS Attachments (
        file_id INTEGER PRIMARY KEY AUTOINCREMENT,
        request_id INTEGER,
        file_name TEXT,
        file_url TEXT,
        uploaded_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY(request_id) REFERENCES Expense_Requests(request_id)
    )
    """)
    
    # 創建審核歷史紀錄表
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS Approval_Logs (
        approval_id INTEGER PRIMARY KEY AUTOINCREMENT,
        request_id INTEGER,
        approver_id INTEGER,
        action TEXT,
        comments TEXT,
        quota_deducted REAL,
        reviewed_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY(request_id) REFERENCES Expense_Requests(request_id),
        FOREIGN KEY(approver_id) REFERENCES Users(user_id)
    )
    """)
    
    # 預填角色
    cursor.execute("SELECT COUNT(*) FROM Roles")
    if cursor.fetchone()[0] == 0:
        cursor.executemany("INSERT INTO Roles (role_name, default_quota) VALUES (?, ?)", [
            ('applicant', 1000.0),
            ('manager', 50000.0),
            ('admin', 100000.0)
        ])
        
    # 預填部門
    cursor.execute("SELECT COUNT(*) FROM Departments")
    if cursor.fetchone()[0] == 0:
        cursor.executemany("INSERT INTO Departments (dept_name) VALUES (?)", [
            ('市場部',), ('銷售部',), ('研發部',), ('業務部',), ('人事處',)
        ])
        
    # 預填費用類別
    cursor.execute("SELECT COUNT(*) FROM Categories")
    if cursor.fetchone()[0] == 0:
        cursor.executemany("INSERT INTO Categories (category_name) VALUES (?)", [
            ('辦公用品',), ('出差費用',), ('會議費用',), ('培訓費用',), ('其他',)
        ])
        
    # 預填測試使用者
    cursor.execute("SELECT COUNT(*) FROM Users")
    if cursor.fetchone()[0] == 0:
        users_data = [
            ('王', '明', 'M', 'applicant@company.com', generate_password_hash('123456'), 1, 1000.0, '啟用'),
            ('李', '經理', 'M', 'manager@company.com', generate_password_hash('123456'), 2, 50000.0, '啟用'),
            ('陳', '管理員', 'M', 'admin@company.com', generate_password_hash('123456'), 3, 100000.0, '啟用')
        ]
        cursor.executemany("""
        INSERT INTO Users (last_name, first_name, gender, email, password_hash, role_id, current_quota, status)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, users_data)
        
    # 預填測試申請單
    cursor.execute("SELECT COUNT(*) FROM Expense_Requests")
   # 預填測試申請單
    cursor.execute("SELECT COUNT(*) FROM Expense_Requests")
    if cursor.fetchone()[0] == 0:
        # 申請單 1：已自動通過，金額 450，申請人王明 (ID 1)
        # 💡 修正 1：在 SQL 尾巴加上 RETURNING request_id，並把問號 ? 改成 %s
        cursor.execute("""
        INSERT INTO Expense_Requests (applicant_id, amount, expense_date, category_id, dept_id, status, description, created_at)
        VALUES (1, 450.0, '2024-01-15', 1, 1, '已通過', '列印紙張', '2024-01-16 10:00:00') RETURNING request_id
        """)
        
        # 💡 修正 2：改用 fetchone() 拿回剛剛 Postgres 產生的流水號 ID
        req1_id = cursor.fetchone()['request_id']
        
        # 💡 修正 3：把這段 SQL 的預留記號 ? 改成 %s
        cursor.execute("""
        INSERT INTO Approval_Logs (request_id, approver_id, action, comments, quota_deducted, reviewed_at)
        VALUES (%s, NULL, '系統自動核准', '系統自動核准 (金額 <= 1000)', 0.0, '2024-01-16 10:00:00')
        """, (req1_id,))
        
        # 扣除額度
        cursor.execute("UPDATE Users SET current_quota = current_quota - 450.0 WHERE user_id = 1")
        
        # 申請單 2：待審核，金額 5000，申請人李經理 (ID 2)
        cursor.execute("""
        INSERT INTO Expense_Requests (applicant_id, amount, expense_date, category_id, dept_id, status, description, created_at)
        VALUES (2, 5000.0, '2024-02-01', 2, 2, '審核中', '深圳出差機票酒店', '2024-02-02 11:30:00')
        """)
        
    conn.commit()
    conn.close()

init_db()

# ============ 輔助函數 ============
import psycopg2
from psycopg2.extras import RealDictCursor

def get_db():
    # 優先讀取 Render 雲端環境變數，若本機測試則使用後方的 Supabase 連線網址
    db_url = os.environ.get("DATABASE_URL", "postgresql://postgres:[guanlixue0609]@aws-0-ap-southeast-1.pooler.supabase.com:6543/postgres")

    # 修正部分平台不支援 postgres:// 開頭的問題
    if db_url.startswith("postgres://"):
        db_url = db_url.replace("postgres://", "postgresql://", 1)

    conn = psycopg2.connect(db_url)
    # 讓 Postgres 回傳的資料格式等同於原本 SQLite 的 Row dict 格式
    conn.cursor_factory = RealDictCursor
    return conn

def check_ip_locked(ip_address):
    """檢查 IP 是否連續失敗 5 次且最近失敗在 2 小時內"""
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT is_success, login_time FROM IP_Log 
        WHERE ip_address = ? 
        ORDER BY login_time DESC LIMIT 5
    """, (ip_address,))
    recent = cursor.fetchall()
    conn.close()
    
    if len(recent) == 5 and all(not r['is_success'] for r in recent):
        # 取得最後一次失敗的時間
        last_fail_str = recent[0]['login_time']
        last_fail_time = datetime.strptime(last_fail_str, '%Y-%m-%d %H:%M:%S' if ' ' in last_fail_str else '%Y-%m-%d')
        if datetime.now() - last_fail_time < timedelta(hours=2):
            return True
    return False

# ============ 權限與登入狀態驗證裝飾器 ============
def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        user_id = session.get('user_id')
        if not user_id:
            return jsonify({"error": "未登入"}), 401
            
        conn = get_db()
        cursor = conn.cursor()
        cursor.execute("SELECT status, login_count FROM Users WHERE user_id = ?", (user_id,))
        user = cursor.fetchone()
        conn.close()
        
        if not user:
            return jsonify({"error": "使用者不存在"}), 401
            
        if user['status'] == '已註銷':
            session.clear()
            return jsonify({"error": "此帳號已註銷"}), 403
            
        if user['status'] == '鎖定':
            return jsonify({"error": "此帳號已被鎖定"}), 403
            
        # 連接攔截器：若 login_count >= 5 且不是修改密碼請求，則強制返回密碼修改錯誤
        if user['login_count'] >= 5 and request.path != '/api/change-password':
            return jsonify({"error": "密碼過期，請先更新密碼", "needsPasswordChange": True}), 403
            
        return f(*args, **kwargs)
    return decorated_function

# ============ 靜態網頁託管 ============
@app.route('/')
def serve_index():
    return send_file('index.html')

@app.route('/uploads/<filename>')
def serve_uploaded_file(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename)

# ============ 認證 API ============
@app.route('/api/login', methods=['POST'])
def login():
    ip_address = request.remote_addr
    
    # 檢查 IP 是否鎖定
    if check_ip_locked(ip_address):
        return jsonify({"error": "您的IP已被鎖定，請稍後再試"}), 403
        
    data = request.json
    email = data.get('email')
    password = data.get('password')
    
    if not email or not password:
        return jsonify({"error": "電子信箱與密碼為必填項"}), 400
        
    conn = get_db()
    cursor = conn.cursor()
    
    cursor.execute("""
        SELECT u.*, r.role_name, r.default_quota 
        FROM Users u 
        JOIN Roles r ON u.role_id = r.role_id 
        WHERE u.email = ?
    """, (email,))
    user = cursor.fetchone()
    
    if not user:
        # 紀錄 IP 失敗日誌 (找不到對應 user_id 時填 Null)
        cursor.execute("""
            INSERT INTO IP_Log (ip_address, user_id, is_success, failed_reason, login_time)
            VALUES (?, NULL, 0, '信箱不存在', ?)
        """, (ip_address, datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
        conn.commit()
        conn.close()
        return jsonify({"error": "電子郵件或密碼錯誤"}), 400
        
    user_id = user['user_id']
    status = user['status']
    
    if status == '已註銷':
        conn.close()
        return jsonify({"error": "該帳號已註銷"}), 403
    if status == '鎖定':
        conn.close()
        return jsonify({"error": "該帳號已被鎖定"}), 403
        
    # 驗證密碼
    if not check_password_hash(user['password_hash'], password):
        # 登入失敗紀錄與次數增加
        cursor.execute("""
            INSERT INTO IP_Log (ip_address, user_id, is_success, failed_reason, login_time)
            VALUES (?, ?, 0, '密碼錯誤', ?)
        """, (ip_address, user_id, datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
        
        cursor.execute("UPDATE Users SET login_failures = login_failures + 1 WHERE user_id = ?", (user_id,))
        
        # 獲取最近 IP 連續失敗次數，用於提示
        cursor.execute("SELECT login_failures FROM Users WHERE user_id = ?", (user_id,))
        failures = cursor.fetchone()[0]
        
        conn.commit()
        conn.close()
        
        attempts_left = 5 - failures
        if failures >= 5:
            # 此處雖然是 IP 鎖定，但也在介面上返回鎖定訊息
            return jsonify({"error": "密碼錯誤5次，您的IP已被鎖定2小時"}), 403
        return jsonify({"error": f"密碼錯誤，還有{attempts_left}次機會"}), 400
        
    # 登入成功邏輯
    # 1. 獲取上次成功登入時間之後的所有失敗次數
    cursor.execute("""
        SELECT login_time FROM IP_Log 
        WHERE user_id = ? AND is_success = 1 
        ORDER BY login_time DESC LIMIT 1
    """, (user_id,))
    last_success_row = cursor.fetchone()
    
    fail_count_warning = 0
    if last_success_row:
        last_success_time = last_success_row['login_time']
        cursor.execute("""
            SELECT COUNT(*) FROM IP_Log 
            WHERE user_id = ? AND is_success = 0 AND login_time > ?
        """, (user_id, last_success_time))
        fail_count_warning = cursor.fetchone()[0]
    else:
        cursor.execute("""
            SELECT COUNT(*) FROM IP_Log 
            WHERE user_id = ? AND is_success = 0
        """, (user_id,))
        fail_count_warning = cursor.fetchone()[0]
        
    # 2. 寫入本次成功登入紀錄至 IP_Log
    cursor.execute("""
        INSERT INTO IP_Log (ip_address, user_id, is_success, failed_reason, login_time)
        VALUES (?, ?, 1, NULL, ?)
    """, (ip_address, user_id, datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
    
    # 3. 重設失敗次數
    cursor.execute("UPDATE Users SET login_failures = 0 WHERE user_id = ?", (user_id,))
    
    # 4. 判斷是否需要強制更換密碼 (當前 login_count >= 5)
    # 注意：我們是在這一次登入"判定"是否已經達到 5 次。
    # 判定完後，如果是正常登入，就更新 login_count + 1
    needs_password_change = False
    if user['login_count'] >= 5:
        needs_password_change = True
        session['user_id'] = user_id  # 允許他們登入以調用變更密碼 API
        conn.commit()
        conn.close()
        return jsonify({
            "needsPasswordChange": True,
            "error": "您的帳號已累計登入超過5次，請變更密碼以確保帳號安全"
        })
        
    # 正常登入，login_count 加一
    cursor.execute("UPDATE Users SET login_count = login_count + 1 WHERE user_id = ?", (user_id,))
    conn.commit()
    
    session['user_id'] = user_id
    
    user_info = {
        "id": user['user_id'],
        "email": user['email'],
        "name": f"{user['last_name']}{user['first_name']}",
        "last_name": user['last_name'],
        "first_name": user['first_name'],
        "gender": user['gender'],
        "role": user['role_name'],
        "budget": user['default_quota'],
        "spent": user['default_quota'] - user['current_quota'],
        "current_quota": user['current_quota'],
        "failCount": fail_count_warning
    }
    
    conn.close()
    return jsonify({"message": "登入成功", "user": user_info})

@app.route('/api/logout', methods=['POST'])
def logout():
    session.clear()
    return jsonify({"message": "已成功登出"})

@app.route('/api/change-password', methods=['POST'])
@login_required
def change_password():
    user_id = session['user_id']
    data = request.json
    new_password = data.get('new_password')
    
    if not new_password or len(new_password) < 6:
        return jsonify({"error": "密碼至少6個字元"}), 400
        
    conn = get_db()
    cursor = conn.cursor()
    password_hash = generate_password_hash(new_password)
    
    # 修改密碼並重設 login_count
    cursor.execute("""
        UPDATE Users 
        SET password_hash = ?, login_count = 0 
        WHERE user_id = ?
    """, (password_hash, user_id))
    
    conn.commit()
    conn.close()
    
    session.clear()  # 修改成功後要求重新登入
    return jsonify({"message": "密碼已更新，請使用新密碼重新登入"})

@app.route('/api/user-info', methods=['GET'])
@login_required
def user_info():
    user_id = session['user_id']
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT u.*, r.role_name, r.default_quota 
        FROM Users u 
        JOIN Roles r ON u.role_id = r.role_id 
        WHERE u.user_id = ?
    """, (user_id,))
    user = cursor.fetchone()
    conn.close()
    
    if not user:
        return jsonify({"error": "使用者不存在"}), 404
        
    return jsonify({
        "id": user['user_id'],
        "email": user['email'],
        "name": f"{user['last_name']}{user['first_name']}",
        "last_name": user['last_name'],
        "first_name": user['first_name'],
        "gender": user['gender'],
        "role": user['role_name'],
        "budget": user['default_quota'],
        "spent": user['default_quota'] - user['current_quota'],
        "current_quota": user['current_quota']
    })

# ============ 報帳管理 API ============
@app.route('/api/reports', methods=['GET'])
@login_required
def get_reports():
    user_id = session['user_id']
    conn = get_db()
    cursor = conn.cursor()
    
    # 取得當前使用者角色
    cursor.execute("SELECT role_name FROM Users u JOIN Roles r ON u.role_id = r.role_id WHERE u.user_id = ?", (user_id,))
    role_name = cursor.fetchone()['role_name']
    
    # 申請者只能看自己的申請單，主管和管理者可以看所有單子
    if role_name == 'applicant':
        cursor.execute("""
            SELECT r.*, c.category_name, d.dept_name, a.file_name, a.file_url
            FROM Expense_Requests r
            JOIN Categories c ON r.category_id = c.category_id
            JOIN Departments d ON r.dept_id = d.dept_id
            LEFT JOIN Attachments a ON r.request_id = a.request_id
            WHERE r.applicant_id = ?
            ORDER BY r.request_id DESC
        """, (user_id,))
    else:
        cursor.execute("""
            SELECT r.*, c.category_name, d.dept_name, a.file_name, a.file_url, u.last_name || u.first_name as applicant_name
            FROM Expense_Requests r
            JOIN Categories c ON r.category_id = c.category_id
            JOIN Departments d ON r.dept_id = d.dept_id
            JOIN Users u ON r.applicant_id = u.user_id
            LEFT JOIN Attachments a ON r.request_id = a.request_id
            ORDER BY r.request_id DESC
        """)
        
    rows = cursor.fetchall()
    conn.close()
    
    reports = []
    for r in rows:
        report = {
            "id": r['request_id'],
            "applicantId": r['applicant_id'],
            "amount": r['amount'],
            "date": r['expense_date'],
            "category": r['category_name'],
            "department": r['dept_name'],
            "status": r['status'],
            "description": r['description'],
            "createdAt": r['created_at'],
            "fileName": r['file_name'],
            "fileUrl": r['file_url']
        }
        if 'applicant_name' in r.keys():
            report['applicantName'] = r['applicant_name']
        reports.append(report)
        
    return jsonify(reports)

@app.route('/api/reports', methods=['POST'])
@login_required
def submit_report():
    user_id = session['user_id']
    amount = float(request.form.get('amount'))
    date = request.form.get('date')
    category_name = request.form.get('category')
    department_name = request.form.get('department')
    description = request.form.get('description')
    
    # 處理檔案上傳
    file = request.files.get('file')
    
    conn = get_db()
    cursor = conn.cursor()
    
    # 查詢類別與部門 ID
    cursor.execute("SELECT category_id FROM Categories WHERE category_name = ?", (category_name,))
    cat_row = cursor.fetchone()
    category_id = cat_row['category_id'] if cat_row else None
    if not category_id:
        # 若不存在則動態新增
        cursor.execute("INSERT INTO Categories (category_name) VALUES (?)", (category_name,))
        category_id = cursor.lastrowid
        
    cursor.execute("SELECT dept_id FROM Departments WHERE dept_name = ?", (department_name,))
    dept_row = cursor.fetchone()
    dept_id = dept_row['dept_id'] if dept_row else None
    if not dept_id:
        cursor.execute("INSERT INTO Departments (dept_name) VALUES (?)", (department_name,))
        dept_id = cursor.lastrowid
        
    # 獲取申請者目前的剩餘個人可用額度
    cursor.execute("SELECT current_quota FROM Users WHERE user_id = ?", (user_id,))
    current_quota = cursor.fetchone()['current_quota']
    
    # 商務邏輯：若金額 <= 1000 且 餘額足夠 -> 自動通過，否則為審核中
    is_auto_approved = (amount <= 1000) and (current_quota >= amount)
    status = '已通過' if is_auto_approved else '審核中'
    
    # 插入報帳申請
    cursor.execute("""
        INSERT INTO Expense_Requests (applicant_id, amount, expense_date, category_id, dept_id, status, description, created_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    """, (user_id, amount, date, category_id, dept_id, status, description, datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
    request_id = cursor.lastrowid
    
    # 存檔附件
    file_saved = False
    if file and file.filename != '':
        filename = secure_filename(file.filename)
        # 為防檔名重複加上時間戳
        timestamp = datetime.now().strftime('%Y%m%d%H%M%S')
        unique_filename = f"{timestamp}_{filename}"
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], unique_filename)
        file.save(file_path)
        
        file_url = f"/uploads/{unique_filename}"
        cursor.execute("""
            INSERT INTO Attachments (request_id, file_name, file_url)
            VALUES (?, ?, ?)
        """, (request_id, filename, file_url))
        file_saved = True
        
    # 自動審批紀錄與額度扣除
    if is_auto_approved:
        cursor.execute("""
            INSERT INTO Approval_Logs (request_id, approver_id, action, comments, quota_deducted, reviewed_at)
            VALUES (?, NULL, '系統自動核准', '系統自動核准 (金額 <= 1000)', 0.0, ?)
        """, (request_id, datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
        
        # 扣減申請者個人可用額度
        cursor.execute("UPDATE Users SET current_quota = current_quota - ? WHERE user_id = ?", (amount, user_id))
        
    conn.commit()
    conn.close()
    
    msg = '報告已自動通過' if is_auto_approved else '報告已提交等待審批'
    return jsonify({"message": msg, "reportId": request_id, "status": status})

@app.route('/api/reports/pending', methods=['GET'])
@login_required
def get_pending_reports():
    user_id = session['user_id']
    conn = get_db()
    cursor = conn.cursor()
    
    # 檢查是否為主管或管理員
    cursor.execute("SELECT role_name FROM Users u JOIN Roles r ON u.role_id = r.role_id WHERE u.user_id = ?", (user_id,))
    role_name = cursor.fetchone()['role_name']
    
    if role_name == 'applicant':
        conn.close()
        return jsonify({"error": "權限不足"}), 403
        
    cursor.execute("""
        SELECT r.*, c.category_name, d.dept_name, u.last_name || u.first_name as applicant_name, a.file_name, a.file_url
        FROM Expense_Requests r
        JOIN Categories c ON r.category_id = c.category_id
        JOIN Departments d ON r.dept_id = d.dept_id
        JOIN Users u ON r.applicant_id = u.user_id
        LEFT JOIN Attachments a ON r.request_id = a.request_id
        WHERE r.status = '審核中'
        ORDER BY r.request_id DESC
    """)
    rows = cursor.fetchall()
    conn.close()
    
    pending = []
    for r in rows:
        pending.append({
            "id": r['request_id'],
            "applicantId": r['applicant_id'],
            "applicantName": r['applicant_name'],
            "amount": r['amount'],
            "date": r['expense_date'],
            "category": r['category_name'],
            "department": r['dept_name'],
            "status": r['status'],
            "description": r['description'],
            "createdAt": r['created_at'],
            "fileName": r['file_name'],
            "fileUrl": r['file_url']
        })
        
    return jsonify(pending)

@app.route('/api/reports/<int:request_id>/approve', methods=['POST'])
@login_required
def approve_report(request_id):
    user_id = session['user_id']
    comments = request.json.get('comments', '同意通過')
    
    conn = get_db()
    cursor = conn.cursor()
    
    # 驗證審核者角色
    cursor.execute("SELECT role_name, current_quota FROM Users u JOIN Roles r ON u.role_id = r.role_id WHERE u.user_id = ?", (user_id,))
    approver = cursor.fetchone()
    
    if approver['role_name'] == 'applicant':
        conn.close()
        return jsonify({"error": "權限不足"}), 403
        
    approver_quota = approver['current_quota']
    
    # 撈取申請單詳情
    cursor.execute("SELECT amount, applicant_id FROM Expense_Requests WHERE request_id = ? AND status = '審核中'", (request_id,))
    req = cursor.fetchone()
    if not req:
        conn.close()
        return jsonify({"error": "找不到此待審核申請或已完成審核"}), 404
        
    amount = req['amount']
    applicant_id = req['applicant_id']
    
    # 檢查主管審核額度是否足夠
    if approver_quota < amount:
        conn.close()
        return jsonify({"error": "您的剩餘可用審核額度不足以核准此筆申請"}), 400
        
    # 1. 更新申請單狀態為「已通過」
    cursor.execute("UPDATE Expense_Requests SET status = '已通過' WHERE request_id = ?", (request_id,))
    
    # 2. 扣除主管 current_quota
    cursor.execute("UPDATE Users SET current_quota = current_quota - ? WHERE user_id = ?", (amount, user_id))
    
    # 3. 扣除申請者的 current_quota (即使可能為負數，因為是主管同意透支額度)
    cursor.execute("UPDATE Users SET current_quota = current_quota - ? WHERE user_id = ?", (amount, applicant_id))
    
    # 4. 寫入審核歷史日誌
    cursor.execute("""
        INSERT INTO Approval_Logs (request_id, approver_id, action, comments, quota_deducted, reviewed_at)
        VALUES (?, ?, '准許通過', ?, ?, ?)
    """, (request_id, user_id, comments, amount, datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
    
    conn.commit()
    conn.close()
    
    return jsonify({"message": "申請已核准通過"})

@app.route('/api/reports/<int:request_id>/reject', methods=['POST'])
@login_required
def reject_report(request_id):
    user_id = session['user_id']
    comments = request.json.get('comments', '退回申請')
    
    conn = get_db()
    cursor = conn.cursor()
    
    # 驗證審核者角色
    cursor.execute("SELECT role_name FROM Users u JOIN Roles r ON u.role_id = r.role_id WHERE u.user_id = ?", (user_id,))
    approver = cursor.fetchone()
    if approver['role_name'] == 'applicant':
        conn.close()
        return jsonify({"error": "權限不足"}), 403
        
    cursor.execute("SELECT status FROM Expense_Requests WHERE request_id = ? AND status = '審核中'", (request_id,))
    if not cursor.fetchone():
        conn.close()
        return jsonify({"error": "找不到此待審核申請或已完成審核"}), 404
        
    # 1. 更新申請狀態為「已退回」
    cursor.execute("UPDATE Expense_Requests SET status = '已退回' WHERE request_id = ?", (request_id,))
    
    # 2. 寫入審核日誌 (無扣除主管額度)
    cursor.execute("""
        INSERT INTO Approval_Logs (request_id, approver_id, action, comments, quota_deducted, reviewed_at)
        VALUES (?, ?, '退回', ?, 0.0, ?)
    """, (request_id, user_id, comments, datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
    
    conn.commit()
    conn.close()
    
    return jsonify({"message": "申請已成功退回"})

# ============ 使用者管理 API ============
@app.route('/api/users', methods=['GET'])
@login_required
def get_users():
    user_id = session['user_id']
    conn = get_db()
    cursor = conn.cursor()
    
    # 驗證管理者角色
    cursor.execute("SELECT role_name FROM Users u JOIN Roles r ON u.role_id = r.role_id WHERE u.user_id = ?", (user_id,))
    role_name = cursor.fetchone()['role_name']
    if role_name != 'admin':
        conn.close()
        return jsonify({"error": "權限不足"}), 403
        
    cursor.execute("""
        SELECT u.user_id, u.last_name, u.first_name, u.gender, u.email, u.current_quota, u.status, r.role_name, r.default_quota
        FROM Users u
        JOIN Roles r ON u.role_id = r.role_id
    """)
    rows = cursor.fetchall()
    conn.close()
    
    users = []
    for u in rows:
        users.append({
            "id": u['user_id'],
            "name": f"{u['last_name']}{u['first_name']}",
            "last_name": u['last_name'],
            "first_name": u['first_name'],
            "gender": u['gender'],
            "email": u['email'],
            "budget": u['default_quota'],
            "spent": u['default_quota'] - u['current_quota'],
            "current_quota": u['current_quota'],
            "role": u['role_name'],
            "status": u['status']
        })
        
    return jsonify(users)

@app.route('/api/users/<int:target_user_id>', methods=['DELETE'])
@login_required
def delete_user(target_user_id):
    user_id = session['user_id']
    conn = get_db()
    cursor = conn.cursor()
    
    # 驗證管理者角色
    cursor.execute("SELECT role_name FROM Users u JOIN Roles r ON u.role_id = r.role_id WHERE u.user_id = ?", (user_id,))
    role_name = cursor.fetchone()['role_name']
    if role_name != 'admin':
        conn.close()
        return jsonify({"error": "權限不足"}), 403
        
    # 軟刪除：將狀態改為 "已註銷"
    cursor.execute("UPDATE Users SET status = '已註銷' WHERE user_id = ?", (target_user_id,))
    conn.commit()
    conn.close()
    
    return jsonify({"message": "用戶帳號已註銷"})

# ============ 數據分析 API ============
@app.route('/api/analytics', methods=['GET'])
@login_required
def get_analytics():
    user_id = session['user_id']
    conn = get_db()
    cursor = conn.cursor()
    
    # 驗證管理者角色
    cursor.execute("SELECT role_name FROM Users u JOIN Roles r ON u.role_id = r.role_id WHERE u.user_id = ?", (user_id,))
    role_name = cursor.fetchone()['role_name']
    if role_name != 'admin':
        conn.close()
        return jsonify({"error": "權限不足"}), 403
        
    # 1. 總金額
    cursor.execute("SELECT SUM(amount) FROM Expense_Requests")
    total_amount = cursor.fetchone()[0] or 0.0
    
    # 2. 報告數量
    cursor.execute("SELECT COUNT(*) FROM Expense_Requests")
    report_count = cursor.fetchone()[0] or 0
    
    # 3. 按類別統計
    cursor.execute("""
        SELECT c.category_name, SUM(r.amount) as amt
        FROM Expense_Requests r
        JOIN Categories c ON r.category_id = c.category_id
        GROUP BY c.category_id
    """)
    cat_rows = cursor.fetchall()
    conn.close()
    
    category_total = {}
    for row in cat_rows:
        category_total[row['category_name']] = row['amt']
        
    return jsonify({
        "totalAmount": total_amount,
        "reportCount": report_count,
        "averageAmount": round(total_amount / report_count, 2) if report_count > 0 else 0.0,
        "categoryTotal": category_total
    })

# ============ Excel 導出 API ============
@app.route('/api/export/reports', methods=['GET'])
@login_required
def export_reports():
    user_id = session['user_id']
    conn = get_db()
    cursor = conn.cursor()
    
    # 取得當前使用者角色與名稱
    cursor.execute("SELECT role_name, last_name || first_name as name FROM Users u JOIN Roles r ON u.role_id = r.role_id WHERE u.user_id = ?", (user_id,))
    user_row = cursor.fetchone()
    role_name = user_row['role_name']
    
    # 篩選資料
    if role_name == 'applicant':
        cursor.execute("""
            SELECT r.request_id, u.last_name || u.first_name as applicant, r.amount, r.expense_date, c.category_name, d.dept_name, r.status, r.description, r.created_at
            FROM Expense_Requests r
            JOIN Users u ON r.applicant_id = u.user_id
            JOIN Categories c ON r.category_id = c.category_id
            JOIN Departments d ON r.dept_id = d.dept_id
            WHERE r.applicant_id = ?
            ORDER BY r.request_id DESC
        """, (user_id,))
    else:
        cursor.execute("""
            SELECT r.request_id, u.last_name || u.first_name as applicant, r.amount, r.expense_date, c.category_name, d.dept_name, r.status, r.description, r.created_at
            FROM Expense_Requests r
            JOIN Users u ON r.applicant_id = u.user_id
            JOIN Categories c ON r.category_id = c.category_id
            JOIN Departments d ON r.dept_id = d.dept_id
            ORDER BY r.request_id DESC
        """)
        
    rows = cursor.fetchall()
    conn.close()
    
    wb = Workbook()
    ws = wb.active
    ws.title = "報帳紀錄"
    
    headers = ["申請單號", "申請人", "金額 (¥)", "費用日期", "類別", "部門", "狀態", "描述", "建立時間"]
    ws.append(headers)
    for row in rows:
        ws.append(list(row))
        
    # 流式傳輸 (Streaming)
    out = BytesIO()
    wb.save(out)
    out.seek(0)
    
    return send_file(
        out,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name="reports.xlsx"
    )

@app.route('/api/export/analytics', methods=['GET'])
@login_required
def export_analytics():
    user_id = session['user_id']
    conn = get_db()
    cursor = conn.cursor()
    
    # 驗證管理者角色
    cursor.execute("SELECT role_name FROM Users u JOIN Roles r ON u.role_id = r.role_id WHERE u.user_id = ?", (user_id,))
    role_name = cursor.fetchone()['role_name']
    if role_name != 'admin':
        conn.close()
        return jsonify({"error": "權限不足"}), 403
        
    # 按類別統計
    cursor.execute("""
        SELECT c.category_name, SUM(r.amount) as total_amount, COUNT(r.request_id) as count, AVG(r.amount) as avg_amount
        FROM Expense_Requests r
        JOIN Categories c ON r.category_id = c.category_id
        GROUP BY c.category_id
    """)
    rows = cursor.fetchall()
    conn.close()
    
    wb = Workbook()
    ws = wb.active
    ws.title = "報帳數據分析"
    
    headers = ["類別", "總報帳額 (¥)", "報告數量", "平均金額 (¥)"]
    ws.append(headers)
    for row in rows:
        ws.append([row[0], row[1], row[2], round(row[3], 2)])
        
    out = BytesIO()
    wb.save(out)
    out.seek(0)
    
    return send_file(
        out,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name="analytics.xlsx"
    )

@app.route('/api/register', methods=['POST'])
def register():
    data = request.json
    name = data.get('name')
    email = data.get('email')
    role_name = data.get('role')
    password = data.get('password')
    gender = data.get('gender', 'M')
    
    if not name or not email or not role_name or not password:
        return jsonify({"error": "姓名、帳號、密碼與身分為必填項目"}), 400
        
    if len(password) < 6:
        return jsonify({"error": "密碼至少為 6 個字元"}), 400
        
    conn = get_db()
    cursor = conn.cursor()
    
    cursor.execute("SELECT COUNT(*) FROM Users WHERE email = ?", (email,))
    if cursor.fetchone()[0] > 0:
        conn.close()
        return jsonify({"error": "此帳號已存在"}), 400
        
    cursor.execute("SELECT role_id, default_quota FROM Roles WHERE role_name = ?", (role_name,))
    role = cursor.fetchone()
    if not role:
        conn.close()
        return jsonify({"error": "無效的註冊身分"}), 400
        
    role_id = role['role_id']
    default_quota = role['default_quota']
    
    if len(name) > 0:
        last_name = name[0]
        first_name = name[1:]
    else:
        last_name = ''
        first_name = ''
        
    password_hash = generate_password_hash(password)
    
    try:
        cursor.execute("""
            INSERT INTO Users (last_name, first_name, gender, email, password_hash, role_id, current_quota, status)
            VALUES (?, ?, ?, ?, ?, ?, ?, '啟用')
        """, (last_name, first_name, gender, email, password_hash, role_id, default_quota))
        conn.commit()
    except Exception as e:
        conn.close()
        return jsonify({"error": f"註冊失敗: {str(e)}"}), 500
        
    conn.close()
    return jsonify({"message": "註冊成功，請登入"})

@app.route('/api/users/deactivate-self', methods=['POST'])
@login_required
def deactivate_self():
    user_id = session['user_id']
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("UPDATE Users SET status = '已註銷' WHERE user_id = ?", (user_id,))
    conn.commit()
    conn.close()
    session.clear()
    return jsonify({"message": "您的帳號已成功註銷"})

if __name__ == '__main__':
    app.run(host='127.0.0.1', port=5000, debug=True)
