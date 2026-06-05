import os
from datetime import datetime, timedelta
from functools import wraps
from flask import Flask, request, jsonify, session, send_file, send_from_directory
from flask_cors import CORS
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
from openpyxl import Workbook
from io import BytesIO
import psycopg2
from psycopg2.extras import RealDictCursor

app = Flask(__name__)
app.secret_key = "super_secret_reimbursement_system_key_12345"
CORS(app, supports_credentials=True)

UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB limit

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

# ============ 數據庫連線輔助函數 ============
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
        WHERE ip_address = %s 
        ORDER BY login_time DESC LIMIT 5
    """, (ip_address,))
    recent = cursor.fetchall()
    conn.close()
    
    if len(recent) == 5 and all(not r['is_success'] for r in recent):
        # 取得最後一次失敗的時間
        last_fail_time = recent[0]['login_time']
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
        cursor.execute("SELECT status, login_count FROM Users WHERE user_id = %s", (user_id,))
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
    # 採用絕對路徑導引，避免找不到 index.html 出現錯誤 HTML
    current_dir = os.path.dirname(os.path.abspath(__file__))
    return send_file(os.path.join(current_dir, 'index.html'))

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
        WHERE u.email = %s
    """, (email,))
    user = cursor.fetchone()
    
    if not user:
        # 紀錄 IP 失敗日誌 (找不到對應 user_id 時填 Null)
        cursor.execute("""
            INSERT INTO IP_Log (ip_address, user_id, is_success, failed_reason, login_time)
            VALUES (%s, NULL, false, '信箱不存在', %s)
        """, (ip_address, datetime.now()))
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
            VALUES (%s, %s, false, '密碼錯誤', %s)
        """, (ip_address, user_id, datetime.now()))
        
        cursor.execute("UPDATE Users SET login_failures = login_failures + 1 WHERE user_id = %s", (user_id,))
        
        # 獲取最近 IP 連續失敗次數，用於提示
        cursor.execute("SELECT login_failures FROM Users WHERE user_id = %s", (user_id,))
        failures = cursor.fetchone()['login_failures']
        
        conn.commit()
        conn.close()
        
        attempts_left = 5 - failures
        if failures >= 5:
            return jsonify({"error": "密碼錯誤5次，您的IP已被鎖定2小時"}), 403
        return jsonify({"error": f"密碼錯誤，還有{attempts_left}次機會"}), 400
        
    # 登入成功邏輯
    # 1. 獲取上次成功登入時間之後的所有失敗次數
    cursor.execute("""
        SELECT login_time FROM IP_Log 
        WHERE user_id = %s AND is_success = true 
        ORDER BY login_time DESC LIMIT 1
    """, (user_id,))
    last_success_row = cursor.fetchone()
    
    fail_count_warning = 0
    if last_success_row:
        last_success_time = last_success_row['login_time']
        cursor.execute("""
            SELECT COUNT(*) FROM IP_Log 
            WHERE user_id = %s AND is_success = false AND login_time > %s
        """, (user_id, last_success_time))
        fail_count_warning = cursor.fetchone()['count']
    else:
        cursor.execute("""
            SELECT COUNT(*) FROM IP_Log 
            WHERE user_id = %s AND is_success = false
        """, (user_id,))
        fail_count_warning = cursor.fetchone()['count']
        
    # 2. 寫入本次成功登入紀錄至 IP_Log
    cursor.execute("""
        INSERT INTO IP_Log (ip_address, user_id, is_success, failed_reason, login_time)
        VALUES (%s, %s, true, NULL, %s)
    """, (ip_address, user_id, datetime.now()))
    
    # 3. 重設失敗次數
    cursor.execute("UPDATE Users SET login_failures = 0 WHERE user_id = %s", (user_id,))
    
    # 4. 判斷是否需要強制更換密碼 (當前 login_count >= 5)
    if user['login_count'] >= 5:
        session['user_id'] = user_id  # 允許他們登入以調用變更密碼 API
        conn.commit()
        conn.close()
        return jsonify({
            "needsPasswordChange": True,
            "error": "您的帳號已累計登入超過5次，請變更密碼以確保帳號安全"
        })
        
    # 正常登入，login_count 加一
    cursor.execute("UPDATE Users SET login_count = login_count + 1 WHERE user_id = %s", (user_id,))
    conn.commit()
    
    session['user_id'] = user_id
    
    user_info = {
        "id": user['user_id'],
        "email": user['email'],
    }
