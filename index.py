import os
import json
import importlib
import hmac
import hashlib
import base64
import secrets
import struct
import time
from io import BytesIO

from flask import Flask, render_template_string, request, redirect, url_for, session, send_file, jsonify
from werkzeug.middleware.dispatcher import DispatcherMiddleware
from werkzeug.utils import secure_filename
from datetime import datetime, timedelta
from pathlib import PurePosixPath
import shutil
import signal

try:
    from flask_sock import Sock
except ImportError:
    Sock = None

def detect_raspberry_pi4():
                                                                        
    for model_path in (
        "/proc/device-tree/model",
        "/sys/firmware/devicetree/base/model",
    ):
        try:
            with open(model_path, "rb") as f:
                model = f.read().replace(b"\x00", b"").decode("utf-8", "ignore").strip()
            if "raspberry pi 4" in model.lower():
                return True
        except OSError:
            pass
    try:
        with open("/proc/cpuinfo", "r", encoding="utf-8", errors="ignore") as f:
            cpuinfo = f.read().lower()
        return "raspberry pi 4" in cpuinfo or "compute module 4" in cpuinfo
    except OSError:
        return False

IS_RASPBERRY_PI4 = detect_raspberry_pi4()

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(BASE_DIR, "config.json")
LOCK_FILE = os.path.join(BASE_DIR, "weblogin_lock.json")
LOCK_DURATION_MINUTES = 15

def _read_config_file():
                                                                                    
    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        cfg = json.load(f)
    if not isinstance(cfg, dict):
        raise ValueError("config.json 必須是 JSON object")
    return cfg

def _atomic_write_config(cfg):
                                                                                
    if not isinstance(cfg, dict):
        raise ValueError("Invalid config object")
                                                                                      
    if not cfg.get("username") or not cfg.get("password"):
        raise ValueError("拒絕儲存：config.json 缺少 username/password，避免清除登入資料")
    tmp_file = CONFIG_FILE + ".tmp"
    with open(tmp_file, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2, ensure_ascii=False)
        f.flush()
        os.fsync(f.fileno())
    os.replace(tmp_file, CONFIG_FILE)

def load_portal_config():
                                                                                            
    try:
        cfg = _read_config_file()
    except (OSError, json.JSONDecodeError, ValueError):
        cfg = {}
    cfg.setdefault("layout", {"mode": "stacked", "percentage": 50, "swapped": False})
                                                                                                
    if not cfg.get("admin_password"):
        cfg["admin_password"] = cfg.get("password", "")
    cfg.setdefault("two_factor_enabled", False)
    cfg.setdefault("two_factor_secret", "")
    cfg.setdefault("trusted_devices", [])
    return cfg

_cfg = load_portal_config()
IFRAME_SECRET_KEY = os.environ.get("IFRAME_SECRET_KEY") or _cfg.get("flask_secret_key")
if not IFRAME_SECRET_KEY:
    IFRAME_SECRET_KEY = os.urandom(32).hex()
    if _cfg.get("username") and _cfg.get("password"):
        _cfg["flask_secret_key"] = IFRAME_SECRET_KEY
        try:
            _atomic_write_config(_cfg)
        except Exception:
            pass
os.environ["IFRAME_SECRET_KEY"] = IFRAME_SECRET_KEY

def _configured_module(filename, default_name, data_filename):
                                                                                    
    filename = os.path.basename(str(filename or default_name))
    if not filename.endswith(".py"):
        filename += ".py"
    module_name = os.path.splitext(filename)[0]
    if not module_name.isidentifier():
        raise ValueError("Python 檔名格式不正確")
    module = importlib.import_module(module_name)
    data_filename = os.path.basename(str(data_filename or ""))
    if data_filename:
        data_path = os.path.join(BASE_DIR, data_filename)
        if hasattr(module, "DATA_FILE"):
            module.DATA_FILE = data_path
        if hasattr(module, "DATA_FILE_PATH"):
            module.DATA_FILE_PATH = data_path
    return module

_initial_workshop = load_portal_config().get("workshop", {})

def _workshop_worker2_enabled(settings):
                                                                                          
    web_url = str(settings.get("job2_url", "") or "").strip()
    if web_url:
        return True
    py_file = str(settings.get("job2_py", "cmhk.py") or "").strip()
    json_file = str(settings.get("job2_json", "salary_cmhk.json") or "").strip()
    return bool(py_file and json_file)

cargo = _configured_module(_initial_workshop.get("job1_py", "cargo.py"), "cargo.py",
                           _initial_workshop.get("job1_json", "salary_cargo.json"))
if _workshop_worker2_enabled(_initial_workshop):
    CMHK = _configured_module(_initial_workshop.get("job2_py", "cmhk.py"), "cmhk.py",
                              _initial_workshop.get("job2_json", "salary_cmhk.json"))
else:
    CMHK = None
cargo_app = cargo.app
cmhk_app = CMHK.app if CMHK is not None else None

app = Flask(__name__)
sock = Sock(app) if Sock is not None else None

from games.dinogame import dino_bp
from games.big2 import big2_bp
from games import djdino as djdino_game
app.register_blueprint(dino_bp)
app.register_blueprint(big2_bp)

# DJ dino RUN 3D：直接由 Salary Tool 的 Flask 主伺服器提供，毋須另開 8510 Port。
_djdino_html = djdino_game.HTML.replace('src="audio.js"', 'src="/djdino/audio.js"').replace(
    'src="game.js"', 'src="/djdino/game.js"'
)
_djdino_js = djdino_game.GAME_JS.replace("/api/scores", "/djdino/api/scores")

@app.route("/djdino/", methods=["GET"])
@app.route("/djdino", methods=["GET"])
def djdino_home():
    return app.response_class(_djdino_html, mimetype="text/html")

@app.route("/djdino/game.js", methods=["GET"])
def djdino_game_js():
    return app.response_class(_djdino_js, mimetype="application/javascript")

@app.route("/djdino/audio.js", methods=["GET"])
def djdino_audio_js():
    return app.response_class(djdino_game.AUDIO_JS, mimetype="application/javascript")

@app.route("/djdino/api/scores", methods=["GET", "POST"])
def djdino_scores_api():
    if request.method == "GET":
        return jsonify(djdino_game.load_scores()[:5])
    try:
        data = request.get_json(silent=True) or {}
        name = "".join(c for c in str(data.get("name", "PLAYER")) if c.isprintable())[:10].strip() or "PLAYER"
        score = max(0, min(int(data.get("score", 0)), 999999))
    except (TypeError, ValueError):
        return jsonify({"error": "bad request"}), 400
    return jsonify(djdino_game.save_score(name, score)[:5])
app.secret_key = IFRAME_SECRET_KEY
app.config["SESSION_COOKIE_NAME"] = "iframe_session"
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
app.config["SESSION_COOKIE_SECURE"] = False
app.config["PERMANENT_SESSION_LIFETIME"] = 60 * 60 * 24 * 3650

def generate_totp_secret():
    return base64.b32encode(secrets.token_bytes(20)).decode("ascii").rstrip("=")

def totp_code(secret, timestamp=None):
    if not secret:
        return None
    try:
        padded = secret + "=" * ((8 - len(secret) % 8) % 8)
        key = base64.b32decode(padded, casefold=True)
        counter = int((time.time() if timestamp is None else timestamp) // 30)
        msg = struct.pack(">Q", counter)
        digest = hmac.new(key, msg, hashlib.sha1).digest()
        offset = digest[-1] & 0x0F
        value = struct.unpack(">I", digest[offset:offset + 4])[0] & 0x7FFFFFFF
        return f"{value % 1000000:06d}"
    except Exception:
        return None

def verify_totp(secret, code):
    code = str(code or "").strip()
    if len(code) != 6 or not code.isdigit():
        return False
    now = time.time()
    for offset in (-30, 0, 30):
        if secrets.compare_digest(totp_code(secret, now + offset) or "", code):
            return True
    return False

def make_otpauth_uri(secret, username):
    issuer = "cargo-CMHK Salary"
    label = f"{issuer}:{username}"
    from urllib.parse import quote
    return (
        f"otpauth://totp/{quote(label)}?secret={secret}"
        f"&issuer={quote(issuer)}&algorithm=SHA1&digits=6&period=30"
    )

def make_qr_data_uri(secret, username):
    import qrcode
    qr = qrcode.QRCode(box_size=5, border=2)
    qr.add_data(make_otpauth_uri(secret, username))
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white").convert("RGB")
    output = BytesIO()
    img.save(output, format="PNG")
    return "data:image/png;base64," + base64.b64encode(output.getvalue()).decode("ascii")

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="zh-Hant">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{{ workshop.get('portal_title') }}</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <style>
        html, body { width:100%; height:100%; margin:0; }
        body { font-family: system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; }
        iframe { width:100%; height:100%; border:none; display:block; }
        @keyframes swapEffect {
            0% { transform:scale(1); opacity:1; filter:brightness(1); }
            50% { transform:scale(.97); opacity:.4; filter:brightness(1.3); }
            100% { transform:scale(1); opacity:1; filter:brightness(1); }
        }
        .swap-anim { animation:swapEffect .4s ease-in-out; }

        .login-bg {
            background:
                radial-gradient(circle at 20% 20%, rgba(59,130,246,.22), transparent 35%),
                radial-gradient(circle at 80% 80%, rgba(34,197,94,.18), transparent 35%),
                #111827;
        }
        .glass {
            background:rgba(31,41,55,.82);
            border:1px solid rgba(255,255,255,.10);
            box-shadow:0 25px 70px rgba(0,0,0,.35);
            backdrop-filter:blur(18px);
        }
        .login-input {
            width:100%; box-sizing:border-box; background:#111827;
            color:#f9fafb; border:1px solid #374151; border-radius:10px;
            padding:13px 14px; outline:none; transition:.2s;
        }
        .login-input:focus { border-color:#60a5fa; box-shadow:0 0 0 3px rgba(96,165,250,.15); }

        /* Portal-wide 3D mode: login screen and top control bar share the same raised, deep style. */
        html.portal-3d .login-bg { perspective: 1400px; }
        html.portal-3d .login-bg > .glass {
            transform: translateY(-5px) rotateX(2deg);
            transform-style: preserve-3d;
            border: 2px solid rgba(148, 197, 255, .58);
            box-shadow: 0 12px 0 rgba(8, 15, 30, .9), 0 22px 0 rgba(38, 69, 105, .48),
                        0 34px 55px rgba(0,0,0,.62), inset 0 2px 0 rgba(255,255,255,.24);
        }
        html.portal-3d .login-bg .login-input,
        html.portal-3d .login-bg button,
        html.portal-3d .login-bg select,
        html.portal-3d .login-bg a {
            box-shadow: 0 4px 0 rgba(0,0,0,.42), 0 7px 12px rgba(0,0,0,.25), inset 0 1px 0 rgba(255,255,255,.15);
            transition: transform .18s ease, box-shadow .18s ease;
        }
        html.portal-3d .login-bg button:hover,
        html.portal-3d .login-bg a:hover { transform: translateY(-3px); box-shadow: 0 7px 0 rgba(0,0,0,.4), 0 12px 18px rgba(0,0,0,.3); }
        html.portal-3d header {
            position: relative;
            z-index: 60;
            /* Remove the cyan translucent horizontal strip from the portal's top bar. */
            border-bottom: 0 !important;
            box-shadow: none !important;
            transform: none !important;
        }
        html.portal-3d header button,
        html.portal-3d header a,
        html.portal-3d #securityPanelCard,
        html.portal-3d #securityPanelCard details,
        html.portal-3d #securityPanelCard > div > div {
            box-shadow: 0 4px 0 rgba(0,0,0,.42), 0 7px 12px rgba(0,0,0,.28), inset 0 1px 0 rgba(255,255,255,.15);
        }
        html.portal-3d #securityPanelCard {
            border-width: 2px;
            box-shadow: 0 8px 0 rgba(0,0,0,.5), 0 16px 28px rgba(0,0,0,.48), inset 0 2px 0 rgba(255,255,255,.18) !important;
        }
        html:not(.portal-3d) .login-bg > .glass { transform: none; }
    </style>
    <style>
      /* 設定面板字體跟隨主頁日間／夜間主題 */
      html[data-portal-theme="dark"] #securityPanelCard { background-color: rgba(17,24,39,.98) !important; color: #ffffff !important; }
      html[data-portal-theme="dark"] #securityPanelCard *:not(button):not(input):not(img) { color: #ffffff !important; }
      html[data-portal-theme="light"] #securityPanelCard { background-color: rgba(255,255,255,.98) !important; color: #111827 !important; border-color: #d1d5db !important; }
      html[data-portal-theme="light"] #securityPanelCard *:not(button):not(input):not(img) { color: #111827 !important; }
      html[data-portal-theme="light"] #securityPanelCard button { color: #ffffff; }
      html[data-portal-theme="light"] #securityPanelCard input { color: #111827 !important; background-color: #ffffff !important; }
      html[data-portal-theme="dark"] #securityPanelCard select,
      html[data-portal-theme="dark"] #securityPanelCard select option { color: #ffffff !important; background-color: #111827 !important; }
      html[data-portal-theme="light"] #securityPanelCard select,
      html[data-portal-theme="light"] #securityPanelCard select option { color: #111827 !important; background-color: #ffffff !important; }
      html[data-portal-theme="dark"] #securityPanelCard .login-input { color: #ffffff !important; background-color: #111827 !important; }
      html[data-portal-theme="light"] #securityPanelCard .login-input { color: #111827 !important; background-color: #ffffff !important; }
    
/* cargo pane: remove the large enclosing frame around its top selector/viewport.
   Keep the CMHK pane and the actual cargo content untouched. */
#pane-1 {
    border: 0 !important;
    outline: 0 !important;
    box-shadow: none !important;
    border-radius: 0 !important;
}
#pane-1::before,
#pane-1::after {
    border: 0 !important;
    outline: 0 !important;
    box-shadow: none !important;
}
#pane-1 iframe {
    border: 0 !important;
    outline: 0 !important;
    box-shadow: none !important;
    border-radius: 0 !important;
}
html.portal-3d #pane-1 {
    transform: none !important;
    filter: none !important;
}

</style>
</head>
<body class="bg-gray-900 h-screen overflow-hidden select-none">

{% if not logged_in %}
<div class="login-bg min-h-screen w-full flex items-center justify-center p-5 box-border">
    <div class="glass w-full max-w-md rounded-2xl p-7 text-white">
        <div class="text-center mb-6">
            <h1 class="text-xl font-bold"><a href="/big2/" data-random-game="1" target="_blank" rel="noopener noreferrer" title="點擊進入隨機遊戲">{{ workshop.get("login_title") }}</a></h1>
            
        </div>

        {% if error %}
        <div class="mb-4 rounded-lg border border-red-500/30 bg-red-500/10 text-red-300 px-4 py-3 text-sm">
            {{ error }}
        </div>
        {% endif %}

        {% if system_locked %}
        <div class="mb-4 rounded-xl border border-amber-500/30 bg-amber-500/5 p-4 space-y-3">
            <div class="text-sm font-semibold text-amber-300">🔒 系統已安全鎖定</div>
            <div class="text-xs text-gray-400">測試期間如要立即解鎖，請輸入正確的 admin 密碼。</div>
            <form method="POST" action="{{ url_for('unlock_login_lock') }}" class="space-y-2">
                <input class="login-input" type="password" name="admin_password" placeholder="輸入 admin 密碼解鎖" autocomplete="current-password" required>
                <button type="submit" class="w-full rounded-lg bg-amber-600 hover:bg-amber-500 py-2.5 font-semibold transition">
                    🔓 解鎖系統
                </button>
            </form>
        </div>
        {% endif %}

        <form method="POST" action="{{ url_for('login') }}" class="space-y-4">
            {% if pending_2fa %}
            <!-- 第二步：只有完成帳號／密碼後才顯示 2FA -->
            <input type="hidden" name="two_factor_step" value="1">
            <div class="rounded-xl border border-blue-500/30 bg-blue-500/5 p-4 space-y-3">
                <div class="text-sm font-semibold text-gray-200">🔐 雙重認證 (2FA)</div>
                <div>
                    <label class="block text-sm text-gray-300 mb-1">雙重認證驗證碼</label>
                    <input class="login-input text-center tracking-[0.35em]" type="text" name="totp_code" inputmode="numeric" maxlength="6" pattern="[0-9]{6}" autocomplete="one-time-code" placeholder="輸入 6 位數驗證碼" required autofocus>
                </div>
                <p class="text-xs text-gray-400">帳號及密碼已確認，請輸入目前的 6 位數雙重認證驗證碼。</p>
            </div>
            <label class="flex items-center gap-2 text-sm text-gray-300 select-none">
                <input type="checkbox" name="remember_me" value="1" class="w-4 h-4 accent-blue-600" {% if pending_remember_me %}checked{% endif %}>
                <span>記住我的登入狀態</span>
            </label>
            <button type="submit" class="w-full rounded-lg bg-blue-600 hover:bg-blue-500 py-3 font-semibold transition">
                🔐 驗證並登入
            </button>
            {% else %}
            <!-- 第一步：先輸入登入名稱與密碼 -->
            <div class="rounded-xl border border-gray-700 bg-gray-900/50 p-4 space-y-3">
                <div class="flex items-center justify-between text-sm font-semibold text-gray-200"><span>👤 登入名稱與密碼</span><button type="button" id="loginLangToggle" onclick="(function(){var x=document.getElementById('loginLangForm');if(x){x.classList.toggle('hidden');}})()" class="rounded-lg px-2 py-1 hover:bg-gray-700" title="更改語言" aria-label="更改語言">🌏</button></div>
                <div id="loginLangForm" class="hidden !mt-2"><select id="loginLanguageSelect" name="language" class="login-input" aria-label="Language" onchange="changePortalLanguage(this.value)"><option value="zh-Hant">繁體中文</option><option value="en">English</option><option value="th">ภาษาไทย</option></select></div>
                <div><input class="login-input" type="text" name="username" autocomplete="username" placeholder="帳號" aria-label="帳號" required autofocus></div>
                <div><input class="login-input" type="password" name="password" autocomplete="current-password" placeholder="密碼" aria-label="密碼" required></div>
            </div>
            <label class="flex items-center gap-2 text-sm text-gray-300 select-none">
                <input type="checkbox" name="remember_me" value="1" class="w-4 h-4 accent-blue-600">
                <span>記住我的登入狀態</span>
            </label>
            <button type="submit" class="w-full rounded-lg bg-blue-600 hover:bg-blue-500 py-3 font-semibold transition">
                🔐 下一步
            </button>
            {% endif %}
        </form>
    </div>
</div>

{% else %}

<header class="relative bg-gray-800 text-white shadow px-4 py-1 grid grid-cols-[1fr_auto_1fr] items-center z-30 h-[38px] box-border">
    <div class="flex items-center justify-self-start space-x-2">
        {% if not single_pane_mode %}
        <a href="/big2/" data-random-game="1" target="_blank" rel="noopener noreferrer" title="隨機進入 Big 2、Dino 或 DJ dino RUN 3D 遊戲" aria-label="隨機進入 Big 2、Dino 或 DJ dino RUN 3D 遊戲" class="inline-flex items-center justify-center rounded px-1 hover:bg-gray-700 transition"><i class="fa-solid fa-wallet text-blue-400 text-sm"></i></a>
        {% endif %}
        <button type="button" onclick="toggleSecurityPanel()" class="text-xs font-bold tracking-wide bg-transparent border-0 text-white p-0" title="⚙️設定" aria-label="⚙️設定">⚙️</button>

    </div>
    <div class="flex items-center justify-self-end space-x-2 {% if single_pane_mode %}absolute left-1/2 -translate-x-1/2{% endif %}">
        {% if not single_pane_mode %}
        <div id="step-adjust-group" class="flex items-center space-x-1 bg-gray-900 p-0.5 rounded border border-gray-700">
            <button onclick="adjustSplit(-5)" id="btn-dec" title="減少 5%" class="px-2 py-0.5 text-[11px] font-medium rounded bg-gray-800 hover:bg-gray-700 text-gray-200 transition">
                <i id="icon-dec" class="fa-solid fa-arrow-up"></i>
            </button>
            <span id="step-label" class="text-[10px] text-gray-400 px-1 font-mono">50%</span>
            <button onclick="adjustSplit(5)" id="btn-inc" title="增加 5%" class="px-2 py-0.5 text-[11px] font-medium rounded bg-gray-800 hover:bg-gray-700 text-gray-200 transition">
                <i id="icon-inc" class="fa-solid fa-arrow-down"></i>
            </button>
        </div>

        <div class="flex items-center space-x-1 bg-gray-900 p-0.5 rounded border border-gray-700">
            <button onclick="setLayout('stacked')" id="btn-stack" title="上下排列" aria-label="上下排列" class="px-2 py-1 text-sm rounded bg-blue-600 text-white transition shadow">
                <i class="fa-solid fa-rows"></i>
            </button>
            <button onclick="setLayout('side-by-side')" id="btn-side" title="左右排列" aria-label="左右排列" class="px-2 py-1 text-sm rounded text-gray-300 hover:text-white transition">
                <i class="fa-solid fa-columns"></i>
            </button>
            <button onclick="swapPanes()" title="互換兩個網頁位置" aria-label="互換兩個網頁位置" class="px-2 py-1 text-sm rounded text-gray-300 hover:text-white hover:bg-gray-800 transition">
                <i class="fa-solid fa-right-left text-amber-400"></i>
            </button>
        </div>
        {% endif %}

        {% if not single_pane_mode %}
        {% if not is_raspberry_pi4 %}
        <a href="{{ url_for('logout') }}" title="登出" aria-label="登出" class="inline-flex items-center justify-center w-8 h-8 rounded bg-red-600 hover:bg-red-500 text-white transition">
            <i class="fa-solid fa-right-from-bracket"></i>
        </a>
        {% endif %}
        {% endif %}
        {% if single_pane_mode and not is_raspberry_pi4 %}
        <a href="{{ url_for('logout') }}" title="登出" aria-label="登出" class="inline-flex items-center justify-center w-8 h-8 rounded bg-red-600 hover:bg-red-500 text-white transition">
            <i class="fa-solid fa-right-from-bracket"></i>
        </a>
        {% endif %}
    </div>
</header>

<!-- 登入與安全設定：保留原頂部控制列，只額外加入齒輪開啟的設定面板 -->
<div id="securityPanel" class="{% if security_message %}block{% else %}hidden{% endif %} absolute top-[38px] left-2 z-[100] w-[min(680px,calc(100vw-16px))]">
    <div id="securityPanelCard" class="rounded-b-2xl border border-t-0 border-gray-700 bg-gray-900/98 text-white shadow-2xl overflow-visible backdrop-blur-xl">
        <div class="relative flex items-center justify-between px-4 py-2.5 border-b border-gray-700">
            <div class="flex items-center gap-3">
                <button type="button" onclick="showSecuritySettings()" id="settingsTabButton" class="rounded px-3 py-1.5 text-sm font-semibold border border-blue-500/40 bg-gray-800 hover:bg-gray-700 cursor-pointer" aria-label="返回設定">⚙️設定</button>
                <button type="button" onclick="showWorkshopSettings()" id="workshopTabButton" class="rounded px-3 py-1.5 text-sm font-semibold border border-amber-500/30 hover:bg-gray-800/70 cursor-pointer">🛠️ 工作坊</button>
                <button type="button" onclick="openDeveloper()" id="developerTabButton" class="rounded px-3 py-1.5 text-sm font-semibold border border-emerald-500/40 bg-gray-800 hover:bg-emerald-700/60 cursor-pointer" title="Developer" aria-label="Developer"><i class="fa-solid fa-code mr-1 text-emerald-300"></i>Developer</button>

            </div>
        </div>
        <div class="p-3 space-y-2 max-h-[calc(100vh-54px)] overflow-y-auto">            <div id="settingsContent">

            {% if security_message %}<div class="rounded-lg px-3 py-2 text-sm {% if security_error %}bg-red-500/10 text-red-300 border border-red-500/30{% else %}bg-green-500/10 text-green-300 border border-green-500/30{% endif %}">{{ security_message }}</div>{% endif %}

{% if not is_raspberry_pi4 %}
            <details class="rounded-xl border border-gray-700">
                <summary class="cursor-pointer list-none px-4 py-3 text-sm font-semibold hover:bg-gray-800/50">🔐 雙重認證 (2FA)</summary>
                <div class="px-4 pb-4 pt-1">
                <div class="text-sm font-semibold mb-2">🔐 雙重認證 (2FA)</div>
                {% if not security_unlocked %}
                <div class="text-xs text-gray-400 mb-3">QR Code 同密鑰已隱藏。請先輸入 admin 密碼，驗證成功後先會顯示。</div>
                <form method="POST" action="{{ url_for('unlock_2fa_settings') }}" class="flex gap-2 flex-wrap">
                    <input class="login-input max-w-xs" type="password" name="admin_password" placeholder="輸入 admin 密碼" required>
                    <button class="rounded-lg bg-gray-700 hover:bg-gray-600 px-4 py-2 text-sm font-semibold">🔓 顯示 2FA 設定</button>
                </form>
                {% else %}
                <form method="POST" action="{{ url_for('save_2fa_settings') }}" class="space-y-3">
                    <label class="flex items-center gap-2 text-sm"><input type="checkbox" name="two_factor_enabled" value="1" {% if two_factor_enabled %}checked{% endif %} class="w-4 h-4 accent-blue-600"> 啟用雙重認證</label>
                    {% if two_factor_secret %}
                    <div class="grid grid-cols-1 md:grid-cols-2 gap-4 items-center">
                        <div class="flex flex-col items-center gap-2">
                            <img src="{{ two_factor_qr }}" alt="2FA QR Code" class="w-48 h-48 bg-white p-2 rounded-lg">
                            <button type="submit" formaction="{{ url_for('regenerate_2fa_secret') }}" formmethod="POST" class="rounded-lg bg-purple-600 hover:bg-purple-500 px-4 py-2 text-sm font-semibold" title="重新產生新的 2FA QR Code / 密鑰">♻️ 重新產生 QR Code</button>
                        </div>
                        <div><div class="text-xs text-gray-400 mb-1">2FA 密鑰</div><div class="font-mono text-sm break-all bg-black/30 rounded p-2 select-all">{{ two_factor_secret }}</div><div class="text-xs text-gray-500 mt-2">請用 Google Authenticator / Microsoft Authenticator 掃描 QR Code。</div></div>
                    </div>
                    {% else %}
                    <div class="text-xs text-gray-400">尚未產生 2FA 密鑰，按下儲存後會建立新的密鑰。</div>
                    {% endif %}
                    <div class="flex gap-2">
                        <button class="rounded-lg bg-blue-600 hover:bg-blue-500 px-4 py-2 text-sm font-semibold">💾 儲存 2FA 設定</button>
                        <a href="{{ url_for('lock_2fa_settings') }}" class="rounded-lg bg-gray-700 hover:bg-gray-600 px-4 py-2 text-sm">🔒 隱藏 QR / 密鑰</a>
                    </div>
                </form>
                {% endif %}
                </div>
            </details>

{% endif %}

            <details class="rounded-xl border border-blue-500/30" open>
                <summary class="cursor-pointer list-none px-4 py-3 text-sm font-semibold hover:bg-gray-800/50">🌏 更改語言</summary>
                <div class="px-4 pb-4 pt-1"><form method="POST" action="/set_language" class="flex gap-2 items-center"><select name="language" class="login-input" aria-label="更改語言" onchange="changePortalLanguage(this.value)"><option value="zh-Hant">繁體中文</option><option value="en">English</option><option value="th">ภาษาไทย</option></select></form><p class="mt-2 text-xs text-gray-400">選擇後會立即切換語言。</p></div>
            </details>
            <div class="grid grid-cols-3 gap-2">
                <button type="button" id="portalThemeToggle" aria-pressed="true" onclick="togglePortalTheme()"
                        class="min-w-0 rounded-xl border border-gray-700 bg-gray-800 hover:bg-gray-700 px-2 py-3 text-xs sm:text-sm font-semibold text-center">🌙 夜間模式：開</button>
                <button type="button" id="portal3DModeToggle" aria-pressed="false" onclick="togglePortal3DMode()"
                        class="min-w-0 rounded-xl border border-cyan-500/30 bg-cyan-900 hover:bg-cyan-800 text-white px-2 py-3 text-xs sm:text-sm font-semibold text-center">🧊 3D模式：關</button>
                <button type="button" id="portalSoundVolumeButton" aria-label="音效音量" onclick="cyclePortalSoundVolume()"
                        class="min-w-0 rounded-xl border border-violet-500/30 bg-violet-900 hover:bg-violet-800 text-white px-2 py-3 text-xs sm:text-sm font-semibold text-center">🔊 音效：70%</button>
            </div>

{% if not is_raspberry_pi4 %}
            <details class="rounded-xl border border-gray-700">
                <summary class="cursor-pointer list-none px-4 py-3 text-sm font-semibold hover:bg-gray-800/50">👤 更改登入名稱和密碼</summary>
                <div class="px-4 pb-4 pt-1">
                <form method="POST" action="{{ url_for('update_login_credentials') }}" class="grid grid-cols-1 md:grid-cols-3 gap-2">
                    <input class="login-input" type="text" name="new_username" placeholder="新登入名稱" required>
                    <input class="login-input" type="password" name="new_password" placeholder="新登入密碼" required>
                    <input class="login-input" type="password" name="confirm_password" placeholder="確認新登入密碼" required>
                    <button class="md:col-span-3 rounded-lg bg-blue-600 hover:bg-blue-500 py-2 text-sm font-semibold">💾 儲存登入帳密</button>
                </form>
                </div>
            </details>

            <details class="rounded-xl border border-amber-500/30">
                <summary class="cursor-pointer list-none px-4 py-3 text-sm font-semibold hover:bg-gray-800/50">🛡️ 更改 admin 密碼</summary>
                <div class="px-4 pb-4 pt-1">
                <div class="text-sm font-semibold mb-1">🛡️ 更改 admin 密碼</div>
                <div class="text-xs text-gray-400 mb-3">admin 密碼只用於保護雙重認證 QR Code 及密鑰，不會取代原本登入名稱／密碼。</div>
                <form method="POST" action="{{ url_for('update_admin_password') }}" class="grid grid-cols-1 md:grid-cols-2 gap-2">
                    <input class="login-input md:col-span-2" type="password" name="old_admin_password" placeholder="舊 admin 密碼" required>
                    <input class="login-input" type="password" name="new_admin_password" placeholder="新 admin 密碼" required>
                    <input class="login-input" type="password" name="confirm_admin_password" placeholder="確認新 admin 密碼" required>
                    <button class="md:col-span-2 rounded-lg bg-amber-600 hover:bg-amber-500 py-2 text-sm font-semibold">🔑 儲存 admin 密碼</button>
                </form>
                </div>
            </details>
            {% endif %}
            </div>
            <div id="workshopContent" class="hidden">
                <div class="flex items-center justify-between mb-3">
                    <h3 class="text-sm font-bold">🛠️ 工作坊</h3>

                </div>
                <form method="POST" action="{{ url_for('save_workshop_settings') }}" class="px-4 pb-4 pt-2 space-y-3">
                    <div>
                        <label class="block text-xs font-semibold mb-1">主頁標題（瀏覽器標題列）</label>
                        <input class="login-input w-full" name="portal_title" value="{{ workshop.get('portal_title', '📦 cargo + 📶 CMHK 工資對照控制台') }}" required>
                    </div>
                    <div>
                        <label class="block text-xs font-semibold mb-1">登入頁名稱</label>
                        <input class="login-input w-full" name="login_title" value="{{ workshop.get('login_title', '📦 📶cargo + CMHK 工資系統') }}" required>
                        <p class="text-xs text-gray-400 mt-1">點擊登入頁名稱可開啟隨機遊戲。</p>
                    </div>
                    <div class="rounded-lg border border-gray-700 p-3 space-y-2">
                        <div class="text-sm font-semibold">散工 1</div>
                        <label class="block text-xs">Python + JSON 檔案</label>
                        <input class="login-input w-full" name="cargo_files" value="{{ workshop.get('job1_url', '') or (workshop.get('job1_py', 'cargo.py') ~ '+' ~ workshop.get('job1_json', 'salary_cargo.json')) }}" placeholder="job1.py+salary_job1.json 或 https://example.com" required>
                        <p class="text-xs text-gray-400">格式：程式檔.py+資料檔.json</p>
                    </div>
                    <div class="rounded-lg border border-gray-700 p-3 space-y-2">
                        <div class="text-sm font-semibold">散工 2</div>
                        <label class="block text-xs">Python + JSON 檔案</label>
                        <input class="login-input w-full" name="cmhk_files" value="{{ workshop.get('job2_url', '') or (workshop.get('job2_py', 'cmhk.py') ~ ('+' ~ workshop.get('job2_json', 'salary_cmhk.json') if workshop.get('job2_py', 'cmhk.py') and workshop.get('job2_json', 'salary_cmhk.json') else '')) }}" placeholder="job2.py+salary_job2.json 或 https://example.com">
                        <p class="text-xs text-gray-400">格式：程式檔.py+資料檔.json；留空可停用散工 2</p>
                    </div>
                    <div class="flex justify-end gap-2 pt-1">
                        <button type="submit" name="action" value="save" class="rounded-lg bg-blue-600 hover:bg-blue-500 px-4 py-2 text-sm font-semibold">💾 儲存</button>
                    </div>
                </form>
            </div>
        </div>
    </div>
</div>

<main id="main-container" class="flex-1 relative flex flex-col overflow-hidden" style="height:calc(100vh - 38px);">

    <div id="pane-1" class="bg-black flex flex-col overflow-hidden relative group" style="flex:{% if single_pane_mode %}1 1 100%{% else %}1 1 50%{% endif %};">
        {% if not single_pane_mode %}
        <button onclick="toggleMaximize(1)" id="max-btn-1" class="absolute top-2 right-2 z-40 bg-gray-800/80 hover:bg-blue-600 text-white px-2 py-1 rounded text-xs opacity-0 group-hover:opacity-100 transition shadow" title="最大化 / 還原">
            <i id="max-icon-1" class="fa-solid fa-expand"></i>
        </button>
        {% endif %}
        <iframe src="{{ workshop.get('job1_url') or '/cargo/' }}" title="{{ workshop.get('job1_py', 'cargo.py') }}"></iframe>
    </div>

    {% if not single_pane_mode %}
    <div id="pane-2" class="bg-black flex flex-col overflow-hidden relative group" style="flex:1 1 50%;">
        <button onclick="toggleMaximize(2)" id="max-btn-2" class="absolute top-2 right-2 z-40 bg-gray-800/80 hover:bg-blue-600 text-white px-2 py-1 rounded text-xs opacity-0 group-hover:opacity-100 transition shadow" title="最大化 / 還原">
            <i id="max-icon-2" class="fa-solid fa-expand"></i>
        </button>
        <iframe src="{{ workshop.get('job2_url') or '/cmhk/' }}" title="{{ workshop.get('job2_py', 'cmhk.py') }}"></iframe>
    </div>
    {% endif %}
</main>

<script>function openDeveloper(){window.location.href="{{ url_for('developer') }}";}</script>
<script>
(function(){
  let portalTheme = 'dark';
  try { portalTheme = localStorage.getItem('salary_tool_theme') || 'dark'; } catch(e) {}
  function updateThemeControl(){
    document.documentElement.setAttribute('data-portal-theme', portalTheme === 'light' ? 'light' : 'dark');
    const btn=document.getElementById('portalThemeToggle');
    const lang=(document.documentElement.lang || 'zh-Hant').toLowerCase();
    const dark=portalTheme==='dark';
    const text=lang.startsWith('th') ? (dark?'🌙 โหมดกลางคืน: เปิด':'☀️ โหมดกลางคืน: ปิด')
      : lang.startsWith('en') ? (dark?'🌙 Night mode: On':'☀️ Night mode: Off')
      : (dark?'🌙 夜間模式：開':'☀️ 夜間模式：關');
    if(btn){btn.textContent=text;btn.setAttribute('aria-pressed',dark?'true':'false');}
  }
  window.refreshPortalThemeControl=updateThemeControl;
  window.togglePortalTheme=function(){
    portalTheme=portalTheme==='dark'?'light':'dark';
    try { localStorage.setItem('salary_tool_theme',portalTheme); } catch(e) {}
    applyThemeToFrames(); updateThemeControl();
  };
  function applyThemeToFrames(){
    document.querySelectorAll('iframe').forEach(frame=>{try{frame.contentWindow.postMessage({type:'salary-tool-theme',theme:portalTheme},'*');}catch(e){}});
  }
  // Desktop computers default to 3D off. Phones and iPads keep the existing shared setting.
  // A desktop's explicit toggle is remembered separately so another device's config does not
  // unexpectedly turn 3D on for the computer.
  const desktop3DChoiceKey = 'salary_tool_3d_mode_desktop_choice';
  function isDesktopComputer(){
    const ua = navigator.userAgent || '';
    const platform = navigator.platform || '';
    const isIPad = /iPad/i.test(ua) || (platform === 'MacIntel' && navigator.maxTouchPoints > 1);
    const isPhoneOrTablet = /Android|iPhone|iPod|Mobile|Tablet/i.test(ua) || isIPad;
    return !isPhoneOrTablet;
  }
  function getDesktop3DChoice(){
    try {
      const value = localStorage.getItem(desktop3DChoiceKey);
      return value === 'on' ? true : value === 'off' ? false : null;
    } catch(e) { return null; }
  }
  let portal3DMode = false;
  try { portal3DMode = localStorage.getItem('salary_tool_3d_mode') === 'on'; } catch(e) {}
  if(isDesktopComputer()) portal3DMode = getDesktop3DChoice() === true;
  // config.json remains the source of the shared preference for mobile and iPad.
  async function loadPortal3DModeFromConfig(){
    try {
      const response = await fetch('/get_3d_mode', {credentials:'same-origin'});
      if(!response.ok) return;
      const data = await response.json();
      if(typeof data.enabled === 'boolean'){
        const desktopChoice = isDesktopComputer() ? getDesktop3DChoice() : null;
        portal3DMode = isDesktopComputer() ? (desktopChoice === true) : data.enabled;
        try { localStorage.setItem('salary_tool_3d_mode', portal3DMode ? 'on' : 'off'); } catch(e) {}
        applyPortal3DEffect(); apply3DModeToFrames(); update3DModeControl();
      }
      if(typeof data.sound_volume === 'number') applyPortalSoundVolume(data.sound_volume, false);
    } catch(e) {}
  }
  async function savePortal3DModeToConfig(){
    try {
      await fetch('/save_3d_mode', {
        method:'POST', credentials:'same-origin',
        headers:{'Content-Type':'application/json'},
        body:JSON.stringify({enabled:portal3DMode})
      });
    } catch(e) {}
  }
  function applyPortalSoundVolume(value, persist=true){
    const volume=Math.max(0,Math.min(100,Number(value)||0));
    const btn=document.getElementById('portalSoundVolumeButton');
    const lang=(document.documentElement.lang || 'zh-Hant').toLowerCase();
    const text=lang.startsWith('th') ? `🔊 เสียง: ${volume}%`
      : lang.startsWith('en') ? `🔊 Sound: ${volume}%`
      : `🔊 音效：${volume}%`;
    if(btn){btn.textContent=text;btn.setAttribute('aria-label',text);}
    try { localStorage.setItem('salary_tool_sound_volume', String(volume)); } catch(e) {}
    document.querySelectorAll('iframe').forEach(frame=>{try{frame.contentWindow.postMessage({type:'salary-tool-sound-volume',volume:volume},'*');}catch(e){}});
    if(persist) savePortalSoundVolumeToConfig(volume);
  }
  window.cyclePortalSoundVolume=function(){
    const current=Math.max(0,Math.min(100,Number(localStorage.getItem('salary_tool_sound_volume') || 0)));
    const sequence=[0,25,50,75,100];
    const idx=sequence.indexOf(current);\n    const next=idx < 0 ? 25 : sequence[(idx+1)%sequence.length];
    applyPortalSoundVolume(next,true);
  };
  async function savePortalSoundVolumeToConfig(volume){
    try {
      await fetch('/save_sound_volume', {
        method:'POST', credentials:'same-origin',
        headers:{'Content-Type':'application/json'},
        body:JSON.stringify({volume:Math.max(0,Math.min(100,Number(volume)||0))})
      });
    } catch(e) {}
  }
  function applyPortal3DEffect(){
    document.documentElement.classList.toggle('portal-3d', portal3DMode);
    document.documentElement.setAttribute('data-portal-3d', portal3DMode ? 'on' : 'off');
  }
  function update3DModeControl(){
    const btn=document.getElementById('portal3DModeToggle');
    const lang=(document.documentElement.lang || 'zh-Hant').toLowerCase();
    const text=lang.startsWith('th')
      ? (portal3DMode?'🧊 โหมด 3 มิติ: เปิด':'🧊 โหมด 3 มิติ: ปิด')
      : lang.startsWith('en')
        ? (portal3DMode?'🧊 3D Mode: On':'🧊 3D Mode: Off')
        : (portal3DMode?'🧊 3D模式：開':'🧊 3D模式：關');
    if(btn){btn.textContent=text;btn.setAttribute('aria-pressed',portal3DMode?'true':'false');btn.classList.toggle('bg-cyan-500',portal3DMode);btn.classList.toggle('bg-cyan-900',!portal3DMode);}
  }
  window.refreshPortal3DModeControl=update3DModeControl;
  function apply3DModeToFrames(){
    document.querySelectorAll('iframe').forEach(frame=>{try{
      frame.contentWindow.postMessage({type:'salary-tool-3d-mode',enabled:portal3DMode},'*');
      const volume=Number(localStorage.getItem('salary_tool_sound_volume') || 70);
      frame.contentWindow.postMessage({type:'salary-tool-sound-volume',volume:volume},'*');
    }catch(e){}});
  }
  window.togglePortal3DMode=function(){
    portal3DMode=!portal3DMode;
    try {
      localStorage.setItem('salary_tool_3d_mode',portal3DMode?'on':'off');
      if(isDesktopComputer()) localStorage.setItem(desktop3DChoiceKey,portal3DMode?'on':'off');
    } catch(e) {}
    applyPortal3DEffect(); apply3DModeToFrames(); update3DModeControl();
    savePortal3DModeToConfig();
  };
  window.addEventListener('message',e=>{
    if(e.data && e.data.type==='salary-tool-theme-ready') applyThemeToFrames();
    if(e.data && e.data.type==='salary-tool-3d-mode-ready') apply3DModeToFrames();
  });
  window.addEventListener('DOMContentLoaded',()=>{applyPortal3DEffect();updateThemeControl();update3DModeControl();applyPortalSoundVolume(Number(localStorage.getItem('salary_tool_sound_volume') || 70),false);applyThemeToFrames();apply3DModeToFrames();document.querySelectorAll('iframe').forEach(frame=>frame.addEventListener('load',()=>{applyThemeToFrames();apply3DModeToFrames();}));loadPortal3DModeFromConfig();});
  applyPortal3DEffect(); updateThemeControl(); update3DModeControl(); applyPortalSoundVolume(Number(localStorage.getItem('salary_tool_sound_volume') || 70),false);
  setTimeout(()=>{applyThemeToFrames();apply3DModeToFrames();loadPortal3DModeFromConfig();},500);
})();
const mainContainer = document.getElementById('main-container');
const pane1 = document.getElementById('pane-1');
const pane2 = document.getElementById('pane-2');
const btnSide = document.getElementById('btn-side');
const btnStack = document.getElementById('btn-stack');
const iconDec = document.getElementById('icon-dec');
const iconInc = document.getElementById('icon-inc');
const stepLabel = document.getElementById('step-label');

let currentMode = {{ layout.mode|tojson }};
let currentPercentage = {{ layout.percentage|tojson }};
let swapped = {{ layout.swapped|tojson }};
let maximizedPane = null;

function updateActiveButton(activeBtn) {
    [btnSide, btnStack].forEach(btn => {
        btn.className = "px-2 py-0.5 text-[11px] font-medium rounded text-gray-300 hover:text-white transition";
    });
    activeBtn.className = "px-2 py-0.5 text-[11px] font-medium rounded bg-blue-600 text-white transition shadow";
}

function updateSplitRatio(newPercent) {
    if (newPercent < 10) newPercent = 10;
    if (newPercent > 90) newPercent = 90;
    currentPercentage = newPercent;
    pane1.style.flex = `1 1 ${currentPercentage}%`;
    pane2.style.flex = `1 1 ${100-currentPercentage}%`;
    stepLabel.textContent = `${Math.round(currentPercentage)}%`;
}

function saveLayout() {
    fetch('{{ url_for("save_layout") }}', {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({mode: currentMode, percentage: currentPercentage, swapped: swapped})
    }).catch(() => {});
}

function adjustSplit(delta) {
    if (maximizedPane !== null) return;
    updateSplitRatio(currentPercentage + delta);
    saveLayout();
}

function setLayout(mode, persist = true) {
    if (maximizedPane !== null) toggleMaximize(maximizedPane);
    currentMode = mode;
    updateSplitRatio(currentPercentage);

    if (mode === 'side-by-side') {
        mainContainer.className = "flex-1 relative flex flex-row overflow-hidden";
        iconDec.className = "fa-solid fa-arrow-left";
        iconInc.className = "fa-solid fa-arrow-right";
        updateActiveButton(btnSide);
    } else {
        mainContainer.className = "flex-1 relative flex flex-col overflow-hidden";
        iconDec.className = "fa-solid fa-arrow-up";
        iconInc.className = "fa-solid fa-arrow-down";
        updateActiveButton(btnStack);
    }
    if (persist) saveLayout();
}

function swapPanes() {
    if (pane1.nextElementSibling === pane2) {
        mainContainer.insertBefore(pane2, pane1);
        swapped = true;
    } else {
        mainContainer.insertBefore(pane1, pane2);
        swapped = false;
    }

    [pane1, pane2].forEach(pane => {
        pane.classList.remove('swap-anim');
        void pane.offsetWidth;
        pane.classList.add('swap-anim');
    });
    saveLayout();
}

function showWorkshopSettings() {
    const panel = document.getElementById('securityPanel');
    panel.classList.remove('hidden');
    document.getElementById('settingsContent')?.classList.add('hidden');
    document.getElementById('workshopContent')?.classList.remove('hidden');
}
function showSecuritySettings() {
    const panel = document.getElementById('securityPanel');
    if (panel.classList.contains('hidden')) {
        panel.classList.remove('hidden');
    }
    document.getElementById('workshopContent')?.classList.add('hidden');
    document.getElementById('settingsContent')?.classList.remove('hidden');
}
function toggleSecurityPanel() {
    // Top-left ⚙️設定 button toggles the whole panel open/closed.
    const panel = document.getElementById('securityPanel');
    if (!panel) return;
    if (panel.classList.contains('hidden')) {
        showSecuritySettings();
    } else {
        panel.classList.add('hidden');
    }
}

document.addEventListener('click', function(e) {
    const panel = document.getElementById('securityPanel');
    const gear = e.target.closest('[onclick="toggleSecurityPanel()"]');
    if (!panel || panel.classList.contains('hidden') || panel.contains(e.target) || gear) return;
    panel.classList.add('hidden');
});

function toggleMaximize(paneNum) {
    const targetPane = paneNum === 1 ? pane1 : pane2;
    const otherPane = paneNum === 1 ? pane2 : pane1;
    const icon = document.getElementById(`max-icon-${paneNum}`);
    const adjustGroup = document.getElementById('step-adjust-group');

    if (maximizedPane === paneNum) {
        targetPane.style.flex = `1 1 ${currentPercentage}%`;
        otherPane.style.display = "flex";
        otherPane.style.flex = `1 1 ${100-currentPercentage}%`;
        adjustGroup.style.opacity = "1";
        adjustGroup.style.pointerEvents = "auto";
        icon.className = "fa-solid fa-expand";
        maximizedPane = null;
    } else {
        if (maximizedPane !== null) {
            document.getElementById(`max-icon-${maximizedPane}`).className = "fa-solid fa-expand";
        }
        targetPane.style.flex = "1 1 100%";
        otherPane.style.display = "none";
        adjustGroup.style.opacity = "0.4";
        adjustGroup.style.pointerEvents = "none";
        icon.className = "fa-solid fa-compress";
        maximizedPane = paneNum;
    }
}
// 自動偵測裝置：手機上下排列；iPad 及電腦左右排列。
// 以裝置類型判斷，避免 iPadOS 將自己偽裝成桌面瀏覽器時被誤判為手機。
function isPhoneBrowser() {
    const ua = navigator.userAgent || '';
    const isIpad = /iPad/i.test(ua) || (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);
    if (isIpad) return false;
    return /iPhone|iPod|Android.*Mobile|Windows Phone|Mobile/i.test(ua);
}

// 散工 2 的 Python 與 JSON 都空白時，散工 1 已全屏，不執行 split UI 初始化。
if (!{{ single_pane_mode|tojson }}) {
    // 每次開啟主頁都按裝置自動設定；不覆寫使用者儲存的版面設定。
    const autoMode = isPhoneBrowser() ? 'stacked' : 'side-by-side';
    setLayout(autoMode, false);
    updateSplitRatio(currentPercentage);
    if (swapped && pane1.nextElementSibling === pane2) {
        mainContainer.insertBefore(pane2, pane1);
    }
}
</script>
{% endif %}
<script>
// 點擊登入頁 📦 📶 或主頁 wallet 圖示時，隨機開啟其中一款遊戲。
(function () {
    const gamePaths = ['/dinogame/', '/big2/', '/djdino/'];
    document.querySelectorAll('[data-random-game="1"]').forEach(function (link) {
        link.addEventListener('click', function (event) {
            event.preventDefault();
            const selectedPath = gamePaths[Math.floor(Math.random() * gamePaths.length)];
            // 保留 target="_blank" 行為，並在同一點擊事件內開啟，避免彈窗攔截。
            window.open(selectedPath, '_blank', 'noopener,noreferrer');
        });
    });
})();
</script>
<script>
// Keep language controls independent from the login form and other page scripts.
window.changePortalLanguage = async function(language) {
  if (!['zh-Hant','en','th'].includes(language)) return;
  try { localStorage.setItem('portal_language', language); } catch(e) {}
  const body = new URLSearchParams(); body.set('language', language);
  try {
    const response = await fetch('/set_language', {method:'POST', credentials:'same-origin', headers:{'Content-Type':'application/x-www-form-urlencoded;charset=UTF-8','X-Requested-With':'fetch'}, body:body.toString(), cache:'no-store'});
    if (!response.ok) throw new Error('Language save failed: ' + response.status);
    window.location.reload();
  } catch (e) {
    // Native form submission is the fallback if fetch is unavailable.
    const f=document.createElement('form'); f.method='POST'; f.action='/set_language';
    const i=document.createElement('input'); i.type='hidden'; i.name='language'; i.value=language;
    f.appendChild(i); document.body.appendChild(f); f.submit();
  }
};
// Define the gear action in this independent script so earlier UI scripts cannot disable it.
window.toggleSecurityPanel = window.toggleSecurityPanel || function() {
  const panel=document.getElementById('securityPanel');
  if (panel) panel.classList.toggle('hidden');
};
fetch('/get_language', {credentials:'same-origin', cache:'no-store'}).then(r=>r.json()).then(({language:serverLanguage})=>{
 const language = (serverLanguage && ['zh-Hant','en','th'].includes(serverLanguage)) ? serverLanguage : (localStorage.getItem('portal_language') || 'zh-Hant');
 try { localStorage.setItem('portal_language', language); } catch(e) {}
 const dict={
  'zh-Hant':{},
  'en':{'cargo + CMHK 工資系統':'cargo + CMHK Salary System','帳號':'Username','密碼':'Password','記住我的登入狀態':'Remember me','👤 登入名稱與密碼':'👤 Username & Password','🔐 下一步':'🔐 Continue','🔐 驗證並登入':'🔐 Verify & Sign in','⚙️設定':'⚙️ Settings','👤 更改登入名稱和密碼':'👤 Change username and password','🌏 更改語言':'🌏 Change language','🌙 夜間模式':'🌙 Dark mode','🔊 音效音量':'🔊 Sound volume','音效音量獨立運作；調至 0% 即靜音。':'Sound volume works independently of 3D mode; set to 0% to mute.','☀️ 日間模式':'☀️ Light mode','選擇後會立即切換語言。':'Language changes immediately after selection.','上下':'Stacked','左右':'Side by side','互換':'Swap','🔒 登出':'🔒 Log out','套用':'Apply','登入與安全設定':'Settings','⚙️設定':'⚙️ Settings','更改登入名稱和密碼':'Change username and password','目前登入名稱':'Current username','新登入名稱':'New username','目前密碼':'Current password','新密碼':'New password','確認新密碼':'Confirm new password','儲存設定':'Save settings','取消':'Cancel','關閉':'Close','安全設定':'Security settings','雙重認證設定':'Two-factor authentication settings','啟用雙重認證':'Enable two-factor authentication','停用雙重認證':'Disable two-factor authentication','驗證碼':'Verification code','掃描 QR Code':'Scan QR code','語言':'Language','繁體中文':'Traditional Chinese','套用語言':'Apply language','登入':'Sign in','登出':'Log out','記住我':'Remember me','返回':'Back','確認':'Confirm','更新':'Update','更改語言':'Change language','🔐 雙重認證 (2FA)':'🔐 Two-factor authentication (2FA)','雙重認證驗證碼':'Two-factor authentication code','輸入 6 位數驗證碼':'Enter 6-digit verification code','帳號及密碼已確認，請輸入目前的 6 位數雙重認證驗證碼。':'Username and password verified. Enter your current 6-digit authentication code.','⚠️ 雙重認證目前未啟用，請重新登入。':'⚠️ Two-factor authentication is not enabled. Please sign in again.','❌ 雙重認證驗證碼錯誤或已過期。':'❌ Invalid or expired authentication code.','收起':'Collapse','👤 更改登入名稱和密碼':'👤 Change username and password','新登入名稱':'New username','新登入密碼':'New login password','確認新登入密碼':'Confirm new login password','💾 儲存登入帳密':'💾 Save login credentials','🛡️ 更改 admin 密碼':'🛡️ Change admin password','admin 密碼只用於保護雙重認證 QR Code 及密鑰，不會取代原本登入名稱／密碼。':'The admin password protects the 2FA QR code and secret only; it does not replace your regular username and password.','舊 admin 密碼':'Current admin password','新 admin 密碼':'New admin password','確認新 admin 密碼':'Confirm new admin password','🔑 儲存 admin 密碼':'🔑 Save admin password','QR Code 同密鑰已隱藏。請先輸入 admin 密碼，驗證成功後先會顯示。':'The QR code and secret are hidden. Enter the admin password to reveal them.','輸入 admin 密碼':'Enter admin password','🔓 顯示 2FA 設定':'🔓 Show 2FA settings','啟用雙重認證':'Enable two-factor authentication','2FA 密鑰':'2FA secret','請用 Google Authenticator / Microsoft Authenticator 掃描 QR Code。':'Scan the QR code with Google Authenticator or Microsoft Authenticator.','尚未產生 2FA 密鑰，按下儲存後會建立新的密鑰。':'No 2FA secret has been generated yet. Save to create one.','♻️ 重新產生 QR Code':'♻️ Regenerate QR code','💾 儲存 2FA 設定':'💾 Save 2FA settings','🔒 隱藏 QR / 密鑰':'🔒 Hide QR / secret','選擇後會立即切換語言。':'Language changes immediately after selection.','⚠️ 登入名稱／密碼不能為空，而且兩次新密碼必須一致。':'⚠️ Username and password cannot be empty, and the new passwords must match.','❌ 舊 admin 密碼錯誤，不能修改新密碼。':'❌ The current admin password is incorrect. The password cannot be changed.','⚠️ 新 admin 密碼不能為空，而且兩次新密碼必須一致。':'⚠️ The new admin password cannot be empty, and both entries must match.','⚠️ 新 admin 密碼不能與舊密碼相同。':'⚠️ The new admin password must be different from the current password.','✅ admin 密碼已成功更新，2FA QR Code / 密鑰已重新隱藏。':'✅ Admin password updated. The 2FA QR code and secret are hidden again.','❌ admin 密碼錯誤，不能顯示 QR Code / 密鑰。':'❌ Incorrect admin password. The QR code and secret cannot be shown.','🔓 2FA 敏感設定已解鎖。':'🔓 Sensitive 2FA settings unlocked.','♻️ 新的 2FA QR Code / 密鑰已重新產生，舊密鑰已失效。':'♻️ A new 2FA QR code and secret have been generated. The old secret is no longer valid.','✅ 雙重認證設定已儲存。':'✅ Two-factor authentication settings saved.','🔒 QR Code / 密鑰已隱藏。':'🔒 QR code and secret hidden.','❌ 系統已安全鎖定，請稍後再試。':'❌ The system is security-locked. Please try again later.','⚠️ 雙重認證目前未啟用，請重新登入。':'⚠️ Two-factor authentication is not enabled. Please sign in again.','❌ 雙重認證驗證碼錯誤或已過期。':'❌ Invalid or expired two-factor authentication code.','帳號及密碼已確認，請輸入目前的 6 位數雙重認證驗證碼。':'Username and password verified. Enter the current 6-digit two-factor authentication code.','🛠️ 工作坊':'🛠️ Workshop','主頁標題（瀏覽器標題列）':'Main page title (browser title bar)','登入頁名稱':'Login page name','點擊登入頁名稱可開啟隨機遊戲。':'Click the login page name to launch a random game.','散工 1':'Worker 1','散工 2':'Worker 2','Python + JSON 檔案':'Python + JSON files','格式：程式檔.py+資料檔.json':'Format: program.py+data.json','格式：程式檔.py+資料檔.json；留空可停用散工 2':'Format: program.py+data.json; leave blank to disable Worker 2','💾 儲存':'💾 Save','⚠️ 設定已儲存，但指定程式載入失敗：':'⚠️ Settings saved, but the selected module failed to load:','✅ 工作坊設定已儲存。':'✅ Workshop settings saved.','✅ 工作坊設定已還原。':'✅ Workshop settings restored.','工作坊設定已儲存，但指定程式載入失敗':'Workshop settings were saved, but the selected program failed to load'},
  'th':{'cargo + CMHK 工資系統':'ระบบเงินเดือน cargo + CMHK','帳號':'ชื่อผู้ใช้','密碼':'รหัสผ่าน','記住我的登入狀態':'จดจำการเข้าสู่ระบบ','👤 登入名稱與密碼':'👤 ชื่อผู้ใช้และรหัสผ่าน','🔐 下一步':'🔐 ดำเนินการต่อ','🔐 驗證並登入':'🔐 ตรวจสอบและเข้าสู่ระบบ','⚙️設定':'⚙️ ตั้งค่า','👤 更改登入名稱和密碼':'👤 เปลี่ยนชื่อผู้ใช้และรหัสผ่าน','🌏 更改語言':'🌏 เปลี่ยนภาษา','🌙 夜間模式':'🌙 โหมดกลางคืน','🔊 音效音量':'🔊 ระดับเสียง','音效音量獨立運作；調至 0% 即靜音。':'ระดับเสียงทำงานแยกจากโหมด 3 มิติ; ตั้งเป็น 0% เพื่อปิดเสียง','☀️ 日間模式':'☀️ โหมดกลางวัน','選擇後會立即切換語言。':'ภาษาจะเปลี่ยนทันทีหลังเลือก','上下':'แนวตั้ง','左右':'แนวนอน','互換':'สลับ','🔒 登出':'🔒 ออกจากระบบ','套用':'ใช้','登入與安全設定':'การตั้งค่า','⚙️設定':'⚙️ ตั้งค่า','更改登入名稱和密碼':'เปลี่ยนชื่อผู้ใช้และรหัสผ่าน','目前登入名稱':'ชื่อผู้ใช้ปัจจุบัน','新登入名稱':'ชื่อผู้ใช้ใหม่','目前密碼':'รหัสผ่านปัจจุบัน','新密碼':'รหัสผ่านใหม่','確認新密碼':'ยืนยันรหัสผ่านใหม่','儲存設定':'บันทึกการตั้งค่า','取消':'ยกเลิก','關閉':'ปิด','安全設定':'การตั้งค่าความปลอดภัย','雙重認證設定':'การตั้งค่าการยืนยันตัวตนสองขั้นตอน','啟用雙重認證':'เปิดใช้การยืนยันตัวตนสองขั้นตอน','停用雙重認證':'ปิดใช้การยืนยันตัวตนสองขั้นตอน','驗證碼':'รหัสยืนยัน','掃描 QR Code':'สแกน QR Code','語言':'ภาษา','繁體中文':'ภาษาจีนตัวเต็ม','套用語言':'ใช้ภาษา','登入':'เข้าสู่ระบบ','登出':'ออกจากระบบ','記住我':'จดจำฉัน','返回':'กลับ','確認':'ยืนยัน','更新':'อัปเดต','更改語言':'เปลี่ยนภาษา','🔐 雙重認證 (2FA)':'🔐 การยืนยันตัวตนสองขั้นตอน (2FA)','雙重認證驗證碼':'รหัสยืนยันตัวตนสองขั้นตอน','輸入 6 位數驗證碼':'กรอกรหัสยืนยัน 6 หลัก','帳號及密碼已確認，請輸入目前的 6 位數雙重認證驗證碼。':'ยืนยันชื่อผู้ใช้และรหัสผ่านแล้ว กรุณากรอกรหัสยืนยัน 6 หลักปัจจุบัน','⚠️ 雙重認證目前未啟用，請重新登入。':'⚠️ ยังไม่ได้เปิดใช้การยืนยันตัวตนสองขั้นตอน กรุณาเข้าสู่ระบบอีกครั้ง','❌ 雙重認證驗證碼錯誤或已過期。':'❌ รหัสยืนยันไม่ถูกต้องหรือหมดอายุแล้ว','收起':'ย่อ','👤 更改登入名稱和密碼':'👤 เปลี่ยนชื่อผู้ใช้และรหัสผ่าน','新登入名稱':'ชื่อผู้ใช้ใหม่','新登入密碼':'รหัสผ่านเข้าสู่ระบบใหม่','確認新登入密碼':'ยืนยันรหัสผ่านเข้าสู่ระบบใหม่','💾 儲存登入帳密':'💾 บันทึกชื่อผู้ใช้และรหัสผ่าน','🛡️ 更改 admin 密碼':'🛡️ เปลี่ยนรหัสผ่าน admin','admin 密碼只用於保護雙重認證 QR Code 及密鑰，不會取代原本登入名稱／密碼。':'รหัสผ่าน admin ใช้ป้องกัน QR Code และ secret ของ 2FA เท่านั้น ไม่ได้ใช้แทนชื่อผู้ใช้และรหัสผ่านปกติ','舊 admin 密碼':'รหัสผ่าน admin ปัจจุบัน','新 admin 密碼':'รหัสผ่าน admin ใหม่','確認新 admin 密碼':'ยืนยันรหัสผ่าน admin ใหม่','🔑 儲存 admin 密碼':'🔑 บันทึกรหัสผ่าน admin','QR Code 同密鑰已隱藏。請先輸入 admin 密碼，驗證成功後先會顯示。':'QR Code และ secret ถูกซ่อนอยู่ กรุณาป้อนรหัสผ่าน admin เพื่อแสดง','輸入 admin 密碼':'ป้อนรหัสผ่าน admin','🔓 顯示 2FA 設定':'🔓 แสดงการตั้งค่า 2FA','啟用雙重認證':'เปิดใช้การยืนยันตัวตนสองขั้นตอน','2FA 密鑰':'Secret ของ 2FA','請用 Google Authenticator / Microsoft Authenticator 掃描 QR Code。':'สแกน QR Code ด้วย Google Authenticator หรือ Microsoft Authenticator','尚未產生 2FA 密鑰，按下儲存後會建立新的密鑰。':'ยังไม่มี secret ของ 2FA กดบันทึกเพื่อสร้าง secret ใหม่','♻️ 重新產生 QR Code':'♻️ สร้าง QR Code ใหม่','💾 儲存 2FA 設定':'💾 บันทึกการตั้งค่า 2FA','🔒 隱藏 QR / 密鑰':'🔒 ซ่อน QR / secret','選擇後會立即切換語言。':'ภาษาจะเปลี่ยนทันทีหลังเลือก','⚠️ 登入名稱／密碼不能為空，而且兩次新密碼必須一致。':'⚠️ ชื่อผู้ใช้และรหัสผ่านต้องไม่เว้นว่าง และรหัสผ่านใหม่ทั้งสองช่องต้องตรงกัน','❌ 舊 admin 密碼錯誤，不能修改新密碼。':'❌ รหัสผ่าน admin ปัจจุบันไม่ถูกต้อง จึงไม่สามารถเปลี่ยนรหัสผ่านได้','⚠️ 新 admin 密碼不能為空，而且兩次新密碼必須一致。':'⚠️ รหัสผ่าน admin ใหม่ต้องไม่เว้นว่าง และรหัสผ่านทั้งสองช่องต้องตรงกัน','⚠️ 新 admin 密碼不能與舊密碼相同。':'⚠️ รหัสผ่าน admin ใหม่ต้องแตกต่างจากรหัสผ่านปัจจุบัน','✅ admin 密碼已成功更新，2FA QR Code / 密鑰已重新隱藏。':'✅ เปลี่ยนรหัสผ่าน admin สำเร็จแล้ว และซ่อน QR Code กับ secret ของ 2FA อีกครั้ง','❌ admin 密碼錯誤，不能顯示 QR Code / 密鑰。':'❌ รหัสผ่าน admin ไม่ถูกต้อง จึงไม่สามารถแสดง QR Code หรือ secret ได้','🔓 2FA 敏感設定已解鎖。':'🔓 ปลดล็อกการตั้งค่า 2FA ที่ละเอียดอ่อนแล้ว','♻️ 新的 2FA QR Code / 密鑰已重新產生，舊密鑰已失效。':'♻️ สร้าง QR Code และ secret ของ 2FA ใหม่แล้ว secret เดิมใช้ไม่ได้อีกต่อไป','✅ 雙重認證設定已儲存。':'✅ บันทึกการตั้งค่าการยืนยันตัวตนสองขั้นตอนแล้ว','🔒 QR Code / 密鑰已隱藏。':'🔒 ซ่อน QR Code และ secret แล้ว','❌ 系統已安全鎖定，請稍後再試。':'❌ ระบบถูกล็อกเพื่อความปลอดภัย กรุณาลองอีกครั้งภายหลัง','⚠️ 雙重認證目前未啟用，請重新登入。':'⚠️ ยังไม่ได้เปิดใช้การยืนยันตัวตนสองขั้นตอน กรุณาเข้าสู่ระบบอีกครั้ง','❌ 雙重認證驗證碼錯誤或已過期。':'❌ รหัสยืนยันตัวตนสองขั้นตอนไม่ถูกต้องหรือหมดอายุแล้ว','帳號及密碼已確認，請輸入目前的 6 位數雙重認證驗證碼。':'ยืนยันชื่อผู้ใช้และรหัสผ่านแล้ว กรุณากรอกรหัสยืนยันตัวตนสองขั้นตอน 6 หลักปัจจุบัน','🛠️ 工作坊':'🛠️ เวิร์กช็อป','主頁標題（瀏覽器標題列）':'ชื่อหน้าหลัก (แถบชื่อเบราว์เซอร์)','登入頁名稱':'ชื่อหน้าเข้าสู่ระบบ','點擊登入頁名稱可開啟隨機遊戲。':'คลิกชื่อหน้าเข้าสู่ระบบเพื่อเปิดเกมแบบสุ่ม','散工 1':'พนักงาน 1','散工 2':'พนักงาน 2','Python + JSON 檔案':'ไฟล์ Python + JSON','格式：程式檔.py+資料檔.json':'รูปแบบ: program.py+data.json','格式：程式檔.py+資料檔.json；留空可停用散工 2':'รูปแบบ: program.py+data.json; เว้นว่างเพื่อปิดใช้งานพนักงาน 2','💾 儲存':'💾 บันทึก','⚠️ 設定已儲存，但指定程式載入失敗：':'⚠️ บันทึกการตั้งค่าแล้ว แต่โหลดโปรแกรมที่เลือกไม่สำเร็จ:','✅ 工作坊設定已儲存。':'✅ บันทึกการตั้งค่าเวิร์กช็อปแล้ว','✅ 工作坊設定已還原。':'✅ คืนค่าการตั้งค่าเวิร์กช็อปแล้ว','工作坊設定已儲存，但指定程式載入失敗':'บันทึกการตั้งค่าเวิร์กช็อปแล้ว แต่โหลดโปรแกรมที่เลือกไม่สำเร็จ'}
 }[language]||{};
 document.documentElement.lang=language==='en'?'en':(language==='th'?'th':'zh-Hant');
 if (window.refreshPortalThemeControl) window.refreshPortalThemeControl();
 if (window.refreshPortal3DModeControl) window.refreshPortal3DModeControl();
 document.querySelectorAll('input[placeholder],input[aria-label]').forEach(el=>{if(dict[el.placeholder])el.placeholder=dict[el.placeholder];if(dict[el.getAttribute('aria-label')])el.setAttribute('aria-label',dict[el.getAttribute('aria-label')]);});
 const walker=document.createTreeWalker(document.body,NodeFilter.SHOW_TEXT);let n;while(n=walker.nextNode()){let t=n.nodeValue.trim();if(dict[t])n.nodeValue=n.nodeValue.replace(t,dict[t]);}
 document.querySelectorAll('option').forEach(el=>{if(dict[el.textContent.trim()])el.textContent=dict[el.textContent.trim()];});
 document.querySelectorAll('select[name="language"]').forEach(el=>el.value=language);
}).catch(()=>{
 try { const language=localStorage.getItem('portal_language'); if(language && ['zh-Hant','en','th'].includes(language)) document.documentElement.lang=language==='en'?'en':language==='th'?'th':'zh-Hant'; } catch(e) {}
});
</script>

<script>
(function(){
  // Built-in synthesized click sound: no external MP3/WAV file required.
  let audioContext=null;
  let modeEnabled=null;
  function enabled(){
    if(modeEnabled!==null) return modeEnabled;
    try { return localStorage.getItem('salary_tool_3d_mode')==='on'; } catch(e) { return false; }
  }
  function play(){
    // Sound volume works independently of the 3D visual mode.
    try {
      const AudioCtx=window.AudioContext||window.webkitAudioContext;
      if(!AudioCtx) return;
      if(!audioContext) audioContext=new AudioCtx();
      if(audioContext.state==='suspended') audioContext.resume();
      const now=audioContext.currentTime;
      let volume=70;
      try { volume=Math.max(0,Math.min(100,Number(localStorage.getItem('salary_tool_sound_volume')||70))); } catch(e) {}
      if(volume<=0) return;
      const volumeScale=volume/100;
      const osc=audioContext.createOscillator();
      const gain=audioContext.createGain();
      osc.type='sine';
      osc.frequency.setValueAtTime(1050,now);
      osc.frequency.exponentialRampToValueAtTime(720,now+0.045);
      gain.gain.setValueAtTime(0.0001,now);
      gain.gain.exponentialRampToValueAtTime(0.16*volumeScale,now+0.006);
      gain.gain.exponentialRampToValueAtTime(0.0001,now+0.075);
      osc.connect(gain); gain.connect(audioContext.destination);
      osc.start(now); osc.stop(now+0.08);
    } catch(_) {}
  }
  // Delegated capture listener covers regular and dynamically-created controls.
  document.addEventListener('click',function(e){
    const target=e.target;
    if(!target || !target.closest) return;
    const el=target.closest('button,input,select,textarea,a,summary,label,[role],[onclick],[tabindex],.btn,[data-bs-toggle],td,th,.calendar-card,.calendar-day,.day-cell,.date-cell,.clickable');
    if(el) play();
  },true);
  window.addEventListener('message',function(e){
    if(e.data&&e.data.type==='salary-tool-3d-mode'){
      modeEnabled=!!e.data.enabled;
      try { localStorage.setItem('salary_tool_3d_mode',modeEnabled?'on':'off'); } catch(_) {}
    }
    if(e.data&&e.data.type==='salary-tool-sound-volume'){
      const volume=Math.max(0,Math.min(100,Number(e.data.volume)||0));
      try { localStorage.setItem('salary_tool_sound_volume',String(volume)); } catch(_) {}
    }
  });
  try { modeEnabled=localStorage.getItem('salary_tool_3d_mode')==='on'; } catch(_) {}
})();
</script>
</body>
</html>
"""

def load_credentials():
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            cfg = json.load(f)
        return cfg.get("username"), cfg.get("password")
    except Exception:
        return None, None

def check_system_lock():
    try:
        if os.path.exists(LOCK_FILE):
            with open(LOCK_FILE, "r", encoding="utf-8") as f:
                info = json.load(f)
            until = info.get("lock_until", 0)
            if datetime.now().timestamp() < until:
                return True, until
    except Exception:
        pass
    return False, 0

def save_system_lock():
    try:
        until = (datetime.now() + timedelta(minutes=LOCK_DURATION_MINUTES)).timestamp()
        with open(LOCK_FILE, "w", encoding="utf-8") as f:
            json.dump({"lock_until": until}, f)
    except Exception:
        pass

def clear_system_lock():
    try:
        if os.path.exists(LOCK_FILE):
            os.remove(LOCK_FILE)
    except Exception:
        pass

def portal_token():
    return hmac.new(IFRAME_SECRET_KEY.encode(), b"salary-portal", hashlib.sha256).hexdigest()

def set_portal_cookie(response, remember):
    max_age = 60 * 60 * 24 * 3650 if remember else None
    response.set_cookie(
        "portal_access", portal_token(),
        max_age=max_age, httponly=True, secure=False, samesite="Lax", path="/"
    )

def save_config(cfg):
                                                                                    
    current = _read_config_file()
    if not current.get("username") or not current.get("password"):
        raise ValueError("現有 config.json 缺少登入資料，已拒絕覆寫")
                                                                                    
    merged = dict(current)
    merged.update(cfg)
    if not merged.get("username") or not merged.get("password"):
        raise ValueError("拒絕儲存：登入名稱或密碼為空")
    _atomic_write_config(merged)

@app.route("/save_layout", methods=["POST"])
def save_layout():
                            
    if not session.get("portal_logged_in"):
        return {"ok": False, "error": "未登入"}, 401

    try:
        data = request.get_json(silent=True) or {}
        mode = data.get("mode", "stacked")
        if mode not in ("stacked", "side-by-side"):
            mode = "stacked"

        percentage = float(data.get("percentage", 50))
        percentage = max(10, min(90, percentage))

        swapped = bool(data.get("swapped", False))

        cfg = load_portal_config()
        cfg["layout"] = {
            "mode": mode,
            "percentage": percentage,
            "swapped": swapped,
        }
        save_config(cfg)
        return {"ok": True}
    except Exception as e:
        return {"ok": False, "error": str(e)}, 400

@app.route("/get_3d_mode", methods=["GET"])
def get_3d_mode():
                                                               
    cfg = load_portal_config()
    return {"enabled": bool(cfg.get("3d_mode", False)), "sound_volume": max(0, min(100, int(cfg.get("sound_volume", 70))))}

@app.route("/save_sound_volume", methods=["POST"])
def save_sound_volume():
                                                                        
    if not session.get("portal_logged_in"):
        return {"ok": False, "error": "未登入"}, 401
    try:
        data = request.get_json(silent=True) or {}
        volume = data.get("volume")
        if isinstance(volume, bool) or not isinstance(volume, (int, float)):
            return {"ok": False, "error": "volume 必須是數值"}, 400
        volume = max(0, min(100, int(round(volume))))
        cfg = _read_config_file()
        if not cfg.get("username") or not cfg.get("password"):
            raise ValueError("config.json 缺少有效登入資料，音效音量未儲存。")
        cfg["sound_volume"] = volume
        _atomic_write_config(cfg)
        return {"ok": True, "volume": volume}
    except Exception as exc:
        app.logger.error("Sound volume setting was not saved: %s", exc)
        return {"ok": False, "error": "音效音量未能儲存至 config.json"}, 500

@app.route("/save_3d_mode", methods=["POST"])
def save_3d_mode():
                                                                                    
    if not session.get("portal_logged_in"):
        return {"ok": False, "error": "未登入"}, 401
    try:
        data = request.get_json(silent=True) or {}
        enabled = data.get("enabled")
        if not isinstance(enabled, bool):
            return {"ok": False, "error": "enabled 必須是布林值"}, 400
        cfg = _read_config_file()
        if not cfg.get("username") or not cfg.get("password"):
            raise ValueError("config.json 缺少有效登入資料，3D模式未儲存。")
        cfg["3d_mode"] = enabled
        _atomic_write_config(cfg)
        return {"ok": True, "enabled": enabled}
    except Exception as exc:
        app.logger.error("3D mode setting was not saved: %s", exc)
        return {"ok": False, "error": "3D模式未能儲存至 config.json"}, 500

@app.route("/get_language", methods=["GET"])
def get_language():
    from flask import jsonify
    return jsonify({"language": load_portal_config().get("language", "zh-Hant")})

@app.route("/set_language", methods=["POST"])
def set_language():
    lang = request.form.get("language", "zh-Hant")
    if lang not in ("zh-Hant", "en", "th"):
        lang = "zh-Hant"
    try:
                                                                           
        cfg = _read_config_file()
        if not cfg.get("username") or not cfg.get("password"):
            raise ValueError("config.json 缺少有效帳號或密碼；為保護登入資料，語言設定未儲存。")
        cfg["language"] = lang
        _atomic_write_config(cfg)
    except Exception as exc:
        app.logger.error("Language setting was not saved: %s", exc)
        if request.headers.get("X-Requested-With") == "fetch":
            from flask import jsonify
            return jsonify({"ok": False, "error": "語言設定未儲存，config.json 登入資料未被修改。"}), 500
        return "語言設定未儲存；為保護登入資料，config.json 未被覆寫。請檢查設定檔。", 500
    if request.headers.get("X-Requested-With") == "fetch":
        from flask import jsonify
        return jsonify({"ok": True, "language": lang})
    return redirect(request.referrer or url_for("login"))

WORKSHOP_DEFAULTS = {
    "portal_title": "📦 cargo + 📶 CMHK 工資對照控制台",
    "login_title": "📦 📶cargo + CMHK 工資系統",
    "job1_py": "cargo.py",
    "job1_json": "salary_cargo.json",
    "job2_py": "cmhk.py",
    "job2_json": "salary_cmhk.json",
    "job1_url": "",
    "job2_url": "",
}

def load_workshop_settings():
    cfg = load_portal_config()
    saved = cfg.get("workshop", {})
    if not isinstance(saved, dict):
        saved = {}
    return {**WORKSHOP_DEFAULTS, **saved}

@app.route("/save_workshop_settings", methods=["POST"])
def save_workshop_settings():
    if not session.get("portal_logged_in") and not IS_RASPBERRY_PI4:
        return redirect(url_for("login"))
    cfg = _read_config_file()
    if not cfg.get("username") or not cfg.get("password"):
        return "config.json 缺少登入資料，設定未儲存。", 400

    if request.form.get("action") == "reset":
        workshop = dict(WORKSHOP_DEFAULTS)
    else:
        workshop = {}
                                                                            
        def parse_worker_input(field_name, default_py, default_json, allow_blank=False):
            raw = str(request.form.get(field_name, "")).strip()
            if allow_blank and not raw:
                return "", "", ""
            if raw.lower().startswith(("http://", "https://")):
                from urllib.parse import urlparse
                parsed = urlparse(raw)
                if parsed.scheme not in ("http", "https") or not parsed.netloc:
                    raise ValueError(f"{field_name} 網址格式不正確")
                return "", "", raw
            if "+" not in raw:
                raise ValueError(f"{field_name} 必須使用「程式.py+資料.json」格式或 http(s) 網址")
            py_file, json_file = (part.strip() for part in raw.split("+", 1))
            if not py_file or not json_file:
                if allow_blank and not py_file and not json_file:
                    return "", "", ""
                raise ValueError(f"{field_name} 的 Python 檔及 JSON 檔不能只填其中一個")
            return os.path.basename(py_file), os.path.basename(json_file), ""

        job1_py_value, job1_json_value, job1_url_value = parse_worker_input("cargo_files", "cargo.py", "salary_cargo.json")
        job2_py_value, job2_json_value, job2_url_value = parse_worker_input("cmhk_files", "cmhk.py", "salary_cmhk.json", allow_blank=True)

        for key in WORKSHOP_DEFAULTS:
            if key == "job1_py":
                workshop[key] = job1_py_value
                continue
            if key == "job1_json":
                workshop[key] = job1_json_value
                continue
            if key == "job1_url":
                workshop[key] = job1_url_value
                continue
            if key == "job2_py":
                workshop[key] = job2_py_value
                continue
            if key == "job2_json":
                workshop[key] = job2_json_value
                continue
            if key == "job2_url":
                workshop[key] = job2_url_value
                continue
            value = str(request.form.get(key, WORKSHOP_DEFAULTS[key])).strip()
                                                                           
            if key in ("job2_py", "job2_json", "job2_url") and not value:
                workshop[key] = ""
                continue
            if not value:
                value = WORKSHOP_DEFAULTS[key]
            if key.endswith("_py") or key.endswith("_json"):
                                                                       
                value = os.path.basename(value)
                if not value:
                    value = WORKSHOP_DEFAULTS[key]
            workshop[key] = value

    cfg["workshop"] = workshop
    _atomic_write_config(cfg)
                                                                                
    if "mount_workshop_apps" in globals():
        try:
            mount_workshop_apps(workshop)
        except Exception as exc:
            app.logger.exception("工作坊設定已寫入，但載入指定程式失敗")
            session["security_message"] = f"⚠️ 設定已儲存，但指定程式載入失敗：{exc}"
            session["security_error"] = True
            return redirect(url_for("login"))
    session["security_message"] = "✅ 工作坊設定已還原。" if request.form.get("action") == "reset" else "✅ 工作坊設定已儲存。"
    session["security_error"] = False
    return redirect(url_for("login"))

@app.route("/", methods=["GET", "POST"])
def login():
    cfg = load_portal_config()
    layout = cfg.get("layout", {"mode":"stacked","percentage":50,"swapped":False})
    two_factor_enabled = bool(cfg.get("two_factor_enabled") and cfg.get("two_factor_secret"))

    def is_worker2_files_blank(settings=None):
                                                                                
        import tokenize
        import io

        settings = settings or load_workshop_settings()
                                                                  
        if str(settings.get("job2_url", "") or "").strip():
            return False
        py_setting = str(settings.get("job2_py", "cmhk.py") or "").strip()
        json_setting = str(settings.get("job2_json", "salary_cmhk.json") or "").strip()
        if not py_setting and not json_setting:
            return True
        py_name = os.path.basename(py_setting or "cmhk.py")
        json_name = os.path.basename(json_setting or "salary_cmhk.json")

        try:
            with open(os.path.join(BASE_DIR, py_name), "r", encoding="utf-8-sig") as f:
                py_text = f.read()
                                                                           
            py_tokens = [t for t in tokenize.generate_tokens(io.StringIO(py_text).readline)
                         if t.type not in (tokenize.COMMENT, tokenize.NL, tokenize.NEWLINE,
                                           tokenize.ENDMARKER, tokenize.INDENT, tokenize.DEDENT)]
            py_blank = len(py_tokens) == 0

            with open(os.path.join(BASE_DIR, json_name), "r", encoding="utf-8-sig") as f:
                json_text = f.read().strip()
            if not json_text:
                json_blank = True
            else:
                try:
                    json_value = json.loads(json_text)
                    def has_meaningful_json(value):
                        if value is None or value == "":
                            return False
                        if isinstance(value, dict):
                            return any(has_meaningful_json(v) for v in value.values())
                        if isinstance(value, list):
                            return any(has_meaningful_json(v) for v in value)
                        return True
                    json_blank = not has_meaningful_json(json_value)
                except json.JSONDecodeError:
                    json_blank = False
            return py_blank and json_blank
        except (OSError, UnicodeError, tokenize.TokenError, IndentationError):
                                                                
            return False

    def render_login(logged_in=False, error=None, security_message=None, security_error=False, pending_2fa=False, system_locked=False):
        if security_message is None:
            security_message = session.pop("security_message", None)
            security_error = session.pop("security_error", False)
        security_unlocked = bool(session.get("security_unlocked"))
        secret = cfg.get("two_factor_secret", "") if security_unlocked else ""
        username = cfg.get("username", "admin")
        qr = make_qr_data_uri(secret, username) if secret else ""
        return render_template_string(
            HTML_TEMPLATE,
            logged_in=logged_in, error=error, layout=layout,
            two_factor_enabled=two_factor_enabled,
            pending_2fa=pending_2fa,
            pending_remember_me=bool(session.get("pending_remember_me")),
            system_locked=system_locked,
            security_unlocked=security_unlocked,
            two_factor_secret=secret,
            two_factor_qr=qr,
            security_message=security_message,
            security_error=security_error,
            is_raspberry_pi4=IS_RASPBERRY_PI4,
            workshop=load_workshop_settings(),
            single_pane_mode=is_worker2_files_blank(load_workshop_settings()),
        )

    if request.method == "GET" and (session.get("portal_logged_in") or IS_RASPBERRY_PI4):
        if IS_RASPBERRY_PI4:
            session["portal_logged_in"] = True
            session["login_attempts"] = 0
            session.pop("pending_2fa", None)
            session.pop("pending_remember_me", None)
            session.permanent = True
        response = app.make_response(render_login(logged_in=True))
        if IS_RASPBERRY_PI4:
                                                                                 
            set_portal_cookie(response, True)
        return response

    if request.method == "POST":
        is_locked, lock_until = check_system_lock()
        attempts = session.get("login_attempts", 0)
        if attempts >= 5 or is_locked:
            lock_time = datetime.fromtimestamp(lock_until).strftime("%H:%M:%S") if lock_until else ""
            msg = "❌ 系統已安全鎖定，請稍後再試。"
            if lock_time:
                msg += f"（解鎖時間 {lock_time}）"
            return render_login(False, msg, system_locked=True), 429

        if request.form.get("two_factor_step") == "1" and session.get("pending_2fa"):
            if not two_factor_enabled:
                session.pop("pending_2fa", None)
                session.pop("pending_remember_me", None)
                return render_login(False, "⚠️ 雙重認證目前未啟用，請重新登入。"), 401

            if not verify_totp(cfg.get("two_factor_secret", ""), request.form.get("totp_code", "")):
                attempts += 1
                session["login_attempts"] = attempts
                return render_login(False, "❌ 雙重認證驗證碼錯誤或已過期。", pending_2fa=True), 401

            remember_me = bool(request.form.get("remember_me") or session.get("pending_remember_me"))
            session.pop("pending_2fa", None)
            session.pop("pending_remember_me", None)
            session["portal_logged_in"] = True
            session["login_attempts"] = 0
            clear_system_lock()
            session.permanent = remember_me
            response = redirect(url_for("login"))
            set_portal_cookie(response, remember_me)
            return response

        username = request.form.get("username", "")
        password = request.form.get("password", "")
        correct_user, correct_pass = load_credentials()
        if correct_user is None or correct_pass is None:
            return render_login(False, "⚠️ 找不到有效的 config.json 帳號密碼。"), 500

        if username != correct_user or password != correct_pass:
            attempts += 1
            session["login_attempts"] = attempts
            if attempts >= 5:
                save_system_lock()
                return render_login(False, "❌ 失敗次數超過 5 次！系統已鎖定 15 分鐘。"), 429
            return render_login(False, f"❌ 帳號或密碼錯誤（已失敗 {attempts} 次）。"), 401

        if two_factor_enabled:
            session["pending_2fa"] = True
            session["pending_remember_me"] = bool(request.form.get("remember_me"))
            return render_login(False, pending_2fa=True)

        session["portal_logged_in"] = True
        session["login_attempts"] = 0
        clear_system_lock()
        session.permanent = bool(request.form.get("remember_me"))
        response = redirect(url_for("login"))
        set_portal_cookie(response, bool(request.form.get("remember_me")))
        return response

    return render_login(False)

@app.route("/unlock_login_lock", methods=["POST"])
def unlock_login_lock():
                                             
    cfg = load_portal_config()
    admin_pass = request.form.get("admin_password", "")
    if not secrets.compare_digest(admin_pass, str(cfg.get("admin_password", ""))):
        return render_template_string(
            HTML_TEMPLATE,
            logged_in=False,
            error="❌ admin 密碼錯誤，系統仍然保持安全鎖定。",
            layout=cfg.get("layout", {"mode":"stacked","percentage":50,"swapped":False}),
            two_factor_enabled=bool(cfg.get("two_factor_enabled") and cfg.get("two_factor_secret")),
            pending_2fa=False,
            pending_remember_me=False,
            system_locked=True,
            security_unlocked=False,
            two_factor_secret="",
            two_factor_qr="",
            security_message=None,
            security_error=False,
        ), 401

    clear_system_lock()
    session["login_attempts"] = 0
    session.pop("pending_2fa", None)
    session.pop("pending_remember_me", None)
    return redirect(url_for("login"))

@app.route("/update_login_credentials", methods=["POST"])
def update_login_credentials():
    if not session.get("portal_logged_in"):
        return redirect(url_for("login"))
    new_user = request.form.get("new_username", "").strip()
    new_pass = request.form.get("new_password", "")
    confirm = request.form.get("confirm_password", "")
    cfg = load_portal_config()
    if not new_user or not new_pass or new_pass != confirm:
        msg = "⚠️ 登入名稱／密碼不能為空，而且兩次新密碼必須一致。"
        return render_template_string(HTML_TEMPLATE, logged_in=True, error=None, layout=cfg.get("layout", {"mode":"stacked","percentage":50,"swapped":False}), two_factor_enabled=bool(cfg.get("two_factor_enabled") and cfg.get("two_factor_secret")), security_unlocked=bool(session.get("security_unlocked")), two_factor_secret=cfg.get("two_factor_secret", "") if session.get("security_unlocked") else "", two_factor_qr=make_qr_data_uri(cfg.get("two_factor_secret",""), cfg.get("username","admin")) if session.get("security_unlocked") and cfg.get("two_factor_secret") else "", security_message=msg, security_error=True)
    cfg["username"] = new_user
    cfg["password"] = new_pass
    save_config(cfg)
    return redirect(url_for("login"))

@app.route("/update_admin_password", methods=["POST"])
def update_admin_password():
    if not session.get("portal_logged_in"):
        return redirect(url_for("login"))
    old_pass = request.form.get("old_admin_password", "")
    new_pass = request.form.get("new_admin_password", "")
    confirm = request.form.get("confirm_admin_password", "")
    cfg = load_portal_config()
    if not secrets.compare_digest(old_pass, str(cfg.get("admin_password", ""))):
        session["security_message"] = "❌ 舊 admin 密碼錯誤，不能修改新密碼。"
        session["security_error"] = True
        return redirect(url_for("login"))
    if not new_pass or new_pass != confirm:
        session["security_message"] = "⚠️ 新 admin 密碼不能為空，而且兩次新密碼必須一致。"
        session["security_error"] = True
        return redirect(url_for("login"))
    if secrets.compare_digest(old_pass, new_pass):
        session["security_message"] = "⚠️ 新 admin 密碼不能與舊密碼相同。"
        session["security_error"] = True
        return redirect(url_for("login"))
    cfg["admin_password"] = new_pass
    save_config(cfg)
    session["security_unlocked"] = False
    session["security_message"] = "✅ admin 密碼已成功更新，2FA QR Code / 密鑰已重新隱藏。"
    session["security_error"] = False
    return redirect(url_for("login"))

@app.route("/unlock_2fa_settings", methods=["POST"])
def unlock_2fa_settings():
    if not session.get("portal_logged_in"):
        return redirect(url_for("login"))
    cfg = load_portal_config()
    admin_pass = request.form.get("admin_password", "")
    if not secrets.compare_digest(admin_pass, str(cfg.get("admin_password", ""))):
        session["security_message"] = "❌ admin 密碼錯誤，不能顯示 QR Code / 密鑰。"
        session["security_error"] = True
        return redirect(url_for("login"))
    session["security_unlocked"] = True
    session["security_message"] = "🔓 2FA 敏感設定已解鎖。"
    session["security_error"] = False
    return redirect(url_for("login"))

@app.route("/regenerate_2fa_secret", methods=["POST"])
def regenerate_2fa_secret():
                                         
    if not session.get("portal_logged_in"):
        return redirect(url_for("login"))
    cfg = load_portal_config()
    cfg["two_factor_secret"] = generate_totp_secret()
    save_config(cfg)
    session["security_unlocked"] = True
    session["security_message"] = "♻️ 新的 2FA QR Code / 密鑰已重新產生，舊密鑰已失效。"
    session["security_error"] = False
    return redirect(url_for("login"))

@app.route("/save_2fa_settings", methods=["POST"])
def save_2fa_settings():
    if not session.get("portal_logged_in") or not session.get("security_unlocked"):
        return redirect(url_for("login"))
    cfg = load_portal_config()
    if not cfg.get("two_factor_secret"):
        cfg["two_factor_secret"] = generate_totp_secret()
    cfg["two_factor_enabled"] = bool(request.form.get("two_factor_enabled"))
    save_config(cfg)
    session["security_message"] = "✅ 雙重認證設定已儲存。"
    session["security_error"] = False
    return redirect(url_for("login"))

@app.route("/lock_2fa_settings")
def lock_2fa_settings():
    session["security_unlocked"] = False
    session["security_message"] = "🔒 QR Code / 密鑰已隱藏。"
    session["security_error"] = False
    return redirect(url_for("login"))

DEVELOPER_TEMPLATE = r"""
<!doctype html>
<html lang="zh-Hant">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Developer</title>
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.16/codemirror.min.css">
<style>
*{box-sizing:border-box}html,body{margin:0;width:100%;height:100%;overflow:hidden;background:#10151f;color:#e5e7eb;font-family:system-ui,-apple-system,"Segoe UI",sans-serif}
#app{height:100dvh;display:flex;flex-direction:column;background:#10151f}
.top{height:48px;flex:none;display:flex;align-items:center;gap:8px;padding:0 12px;background:#192231;border-bottom:1px solid #334155}
.top strong{font-size:14px;margin-right:auto}.iconbtn{border:1px solid #40516a;background:#263449;color:#e5e7eb;border-radius:7px;width:38px;height:34px;cursor:pointer;font-size:15px}
.iconbtn:hover{background:#365071}.danger{border-color:#8b4141;background:#512b32}
.body{min-height:0;flex:1;display:flex}
.sidebar{width:270px;min-width:190px;background:#151e2b;border-right:1px solid #334155;display:flex;flex-direction:column}
/* Explorer-first layout: hide the editor until a file is opened. */
#app.explorer-mode .sidebar{width:100%;min-width:0;border-right:0}
#app.explorer-mode .editorarea{display:none}
#app.explorer-mode .status{display:none}
#app.edit-mode .sidebar{display:none}
#app.edit-mode .editorarea{display:flex;flex:1}
#app.edit-mode .status{display:flex}
.sidehead{padding:11px 12px;font-size:12px;color:#94a3b8;border-bottom:1px solid #293548;display:flex;justify-content:space-between}
#files{overflow:auto;flex:1;padding:6px}
.file{display:flex;align-items:center;gap:8px;padding:8px;border-radius:6px;font-size:12px;cursor:pointer;color:#cbd5e1}
.file:hover,.file.active{background:#26374e;color:#fff}.file span{overflow:hidden;text-overflow:ellipsis;white-space:nowrap;flex:1}
.file button{background:transparent;border:0;color:#93c5fd;cursor:pointer;padding:3px}
.editorarea{flex:1;min-width:0;display:flex;flex-direction:column}
.filebar{height:37px;flex:none;background:#1c2838;border-bottom:1px solid #334155;display:flex;align-items:center;padding:0 12px;font-size:12px;color:#a5b4c7}
#editor{flex:1;min-height:0;overflow:hidden}
.CodeMirror{height:100%;font:14px/1.55 "JetBrains Mono",Consolas,monospace;background:#111827;color:#e5e7eb}
.CodeMirror-gutters{background:#182232;border-right:1px solid #334155}.CodeMirror-linenumber{color:#64748b}
.CodeMirror-cursor{border-left:2px solid #f8fafc}
.CodeMirror-selected{background:#334155!important}
.CodeMirror-focused .CodeMirror-selected{background:#334155!important}
.cm-s-default .cm-keyword{color:#c4b5fd}.cm-s-default .cm-string{color:#86efac}.cm-s-default .cm-number{color:#fbbf24}.cm-s-default .cm-comment{color:#94a3b8}
.status{height:27px;flex:none;background:#087e75;color:white;display:flex;align-items:center;justify-content:space-between;padding:0 10px;font:11px monospace}
#modeBadge{background:#2563eb;color:#fff;padding:3px 9px;border-radius:4px;font-weight:bold;min-width:95px;text-align:center}
#loginGate{position:fixed;inset:0;background:rgba(3,7,18,.94);display:flex;align-items:center;justify-content:center;padding:20px;z-index:20}
.gatebox{width:min(390px,100%);background:#1b2636;border:1px solid #475569;border-radius:14px;padding:24px;box-shadow:0 20px 70px #0009}
.gatebox h2{margin:0 0 8px;font-size:19px}.gatebox p{font-size:13px;color:#94a3b8;margin:0 0 18px}
.gatebox input{width:100%;padding:12px;border-radius:8px;border:1px solid #475569;background:#0f172a;color:white;margin-bottom:12px}
.gatebox button{width:100%;padding:11px;border:0;border-radius:8px;background:#0d9488;color:white;font-weight:700;cursor:pointer}
#message{font-size:12px;color:#fca5a5;min-height:18px;margin-bottom:7px}
@media(max-width:650px){.sidebar{width:135px;min-width:110px}.top{gap:5px;padding:0 7px}.iconbtn{width:34px}.file{font-size:10px;padding:8px 4px}.CodeMirror{font-size:12px}}
</style>
</head>
<body>
<div id="app" class="explorer-mode">
 <div class="top">
  <strong><i class="fa-solid fa-code" style="color:#6ee7b7"></i> Developer</strong>
  <button class="iconbtn" title="開啟本機檔案" aria-label="開啟本機檔案" onclick="document.getElementById('localPicker').click()"><i class="fa-regular fa-folder-open"></i></button>
  <button class="iconbtn" title="儲存" aria-label="儲存" onclick="saveFile()"><i class="fa-solid fa-floppy-disk"></i></button>
  <button class="iconbtn" title="上載檔案／資料夾" aria-label="上載檔案／資料夾" onclick="document.getElementById('uploadPicker').click()"><i class="fa-solid fa-upload"></i></button>
  <button class="iconbtn" title="下載目前檔案" aria-label="下載目前檔案" onclick="downloadCurrent()"><i class="fa-solid fa-download"></i></button>
  <button class="iconbtn" title="重新整理檔案清單" aria-label="重新整理檔案清單" onclick="refreshFiles()"><i class="fa-solid fa-rotate"></i></button>
  <button class="iconbtn danger" title="離開 Developer" aria-label="離開 Developer" onclick="leaveDeveloper()"><i class="fa-solid fa-xmark"></i></button>
  <input id="localPicker" type="file" accept=".py,.json" hidden onchange="openLocalFile(event)">
  <input id="uploadPicker" type="file" accept=".py,.json" multiple webkitdirectory directory hidden onchange="uploadFiles(event)">
 </div>
 <div class="body">
  <aside class="sidebar">
   <div class="sidehead"><span>EXPLORER · RECURSIVE</span><span id="fileCount">0</span></div>
   <div id="files"></div>
  </aside>
  <section class="editorarea">
   <div class="filebar"><span id="currentName">選擇 .py / .json 檔案以編輯</span><span id="saveState" style="margin-left:auto">Ready</span></div>
   <div id="editor"></div>
  </section>
 </div>
 <div class="status"><span>VIM · EX COMMANDS (:w :q :q! :wq :x)</span><span id="modeBadge">NORMAL</span><span id="position">Ln 1, Col 1</span><span id="saveInfo">Admin protected</span></div>
</div>
<div id="loginGate" {% if unlocked %}style="display:none"{% endif %}>
 <form class="gatebox" method="post" action="{{ url_for('developer_unlock') }}">
  <h2><i class="fa-solid fa-shield-halved" style="color:#6ee7b7"></i> Developer Access</h2>
  <p>請輸入 admin 密碼以進入 Developer。</p>
  <div id="message">{{ error or '' }}</div>
  <input type="password" name="admin_password" placeholder="Admin password" autocomplete="current-password" required autofocus>
  <button type="submit">🔓 解鎖進入</button>
 </form>
</div>
<script src="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.16/codemirror.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.16/mode/python/python.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.16/mode/javascript/javascript.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.16/keymap/vim.min.js"></script>
<script>
let currentFile=null,dirty=false,files=[];
const editor=CodeMirror(document.getElementById('editor'),{
 value:'',lineNumbers:true,keyMap:'vim',mode:'python',indentUnit:4,tabSize:4,
 indentWithTabs:false,autocorrect:false,spellcheck:false,viewportMargin:20,
 extraKeys:{'Ctrl-S':saveFile,'Cmd-S':saveFile}
});
function setViewMode(mode){const app=document.getElementById('app');app.classList.toggle('explorer-mode',mode==='explorer');app.classList.toggle('edit-mode',mode==='edit')}
function setDirty(v){dirty=v;document.getElementById('saveState').textContent=v?'● Modified':'Saved';document.getElementById('saveInfo').textContent=v?'Unsaved changes':'Admin protected'}
function setMode(mode){const el=document.getElementById('modeBadge');el.textContent=(mode||'normal').toUpperCase();el.style.background=(mode==='insert')?'#0f766e':(mode==='visual'?'#7c3aed':'#2563eb')}
editor.on('change',()=>setDirty(true));
editor.on('cursorActivity',()=>{const c=editor.getCursor();document.getElementById('position').textContent=`Ln ${c.line+1}, Col ${c.ch+1}`});
editor.on('vim-mode-change',(cm,mode)=>setMode(mode&&mode.mode));
async function api(url,options={}){const r=await fetch(url,{credentials:'same-origin',...options});if(!r.ok){let d={};try{d=await r.json()}catch(e){}throw new Error(d.error||`HTTP ${r.status}`)}return r}
async function refreshFiles(){
 try{
  const r=await api('/developer/files');files=await r.json();
  document.getElementById('fileCount').textContent=files.length;
  const root=document.getElementById('files');root.innerHTML='';
  files.forEach(name=>{
   const row=document.createElement('div');row.className='file'+(name===currentFile?' active':'');
   const icon=document.createElement('i');icon.className=name.toLowerCase().endsWith('.py')?'fa-brands fa-python':'fa-regular fa-file-code';
   const label=document.createElement('span');label.textContent=name;label.title=name;
   const dl=document.createElement('button');dl.title='下載';dl.innerHTML='<i class="fa-solid fa-download"></i>';dl.onclick=e=>{e.stopPropagation();downloadFile(name)};
   row.append(icon,label,dl);row.onclick=()=>loadFile(name);root.append(row)
  });
 }catch(e){document.getElementById('saveInfo').textContent='File list unavailable';}
}
async function loadFile(name){
 if(dirty&&!confirm('目前檔案尚未儲存，確定放棄修改並開啟另一個檔案？'))return;
 try{
  const r=await api('/developer/file/'+name.split('/').map(encodeURIComponent).join('/'));const d=await r.json();
  currentFile=name;setViewMode('edit');editor.setOption('mode',name.toLowerCase().endsWith('.py')?'python':{name:'javascript',json:true});
  editor.setValue(d.content);editor.clearHistory();setDirty(false);
  document.getElementById('currentName').textContent=name;await refreshFiles();editor.focus();setMode('normal');
 }catch(e){alert('開啟失敗：'+e.message)}
}
async function saveFile(){
 if(!currentFile){alert('請先從左側選擇伺服器檔案');return false}
 try{
  await api('/developer/file/'+currentFile.split('/').map(encodeURIComponent).join('/'),{method:'PUT',headers:{'Content-Type':'application/json'},body:JSON.stringify({content:editor.getValue()})});
  setDirty(false);document.getElementById('saveState').textContent='Saved';
  return true;
 }catch(e){alert('儲存失敗：'+e.message);return false}
}
function closeCurrentFile(force=false){
 if(dirty&&!force&&!confirm('目前檔案尚未儲存，確定放棄修改並關閉？'))return false;
 currentFile=null;setViewMode('explorer');editor.setValue('');editor.clearHistory();setDirty(false);
 document.getElementById('currentName').textContent='選擇 .py / .json 檔案以編輯';
 refreshFiles();editor.focus();setMode('normal');return true;
}
// Register Vim Ex commands. CodeMirror Vim mode is used as the editor engine,
// with explicit Ex commands for the common Vim write/quit workflow.
if(CodeMirror.Vim&&CodeMirror.Vim.defineEx){
 CodeMirror.Vim.defineEx('write','w',function(cm,params){saveFile();});
 CodeMirror.Vim.defineEx('quit','q',function(cm,params){closeCurrentFile(!!(params&&params.bang));});
 CodeMirror.Vim.defineEx('wq','wq',async function(cm,params){
  const saved=await saveFile();if(saved)closeCurrentFile(true);
 });
 CodeMirror.Vim.defineEx('xit','x',async function(cm,params){
  if(dirty){const saved=await saveFile();if(!saved)return;}
  closeCurrentFile(true);
 });
}
function downloadFile(name){window.location.href='/developer/download/'+name.split('/').map(encodeURIComponent).join('/')}
function downloadCurrent(){if(currentFile)downloadFile(currentFile);else alert('請先選擇伺服器檔案')}
function openLocalFile(e){
 const f=e.target.files[0];if(!f)return;
 const reader=new FileReader();reader.onload=()=>{if(dirty&&!confirm('目前檔案尚未儲存，確定放棄修改？'))return;currentFile=null;setViewMode('edit');document.getElementById('currentName').textContent=f.name+' (local copy)';editor.setValue(String(reader.result||''));setDirty(true)};
 reader.readAsText(f);e.target.value='';
}
async function uploadFiles(e){
 const selected=Array.from(e.target.files||[]);e.target.value='';
 for(const f of selected){
  const relative=f.webkitRelativePath||f.name;if(!/\\.(py|json)$/i.test(relative))continue;
  try{
   const form=new FormData();form.append('file',f);form.append('relative_path',relative);
   let r=await api('/developer/upload?conflict=check',{method:'POST',body:form});let d=await r.json();
   if(d.exists){
    const choice=prompt(`檔案「${relative}」已存在。輸入 replace 覆蓋、skip 跳過或 cancel 取消上載：`,'skip');
    if(choice===null||choice.toLowerCase()==='cancel')break;
    if(choice.toLowerCase()==='skip')continue;
    if(choice.toLowerCase()!=='replace'){alert('選項無效，已跳過：'+relative);continue}
   }
   const f2=new FormData();f2.append('file',f);f2.append('relative_path',relative);
   await api('/developer/upload?conflict=replace',{method:'POST',body:f2});
  }catch(err){alert('上載失敗：'+relative+'\\n'+err.message)}
 }
 await refreshFiles();
}
async function leaveDeveloper(){
 if(dirty&&!confirm('目前有未儲存的修改，確定離開？'))return;
 try{await api('/developer/lock',{method:'POST'})}catch(e){}
 window.location.href='/';
}
window.addEventListener('beforeunload',e=>{if(dirty){e.preventDefault();e.returnValue=''}});
setMode('normal');refreshFiles();
</script>
</body></html>
"""

DEVELOPER_UNLOCK_SECONDS = 15 * 60
DEVELOPER_ALLOWED_EXTENSIONS = {".py", ".json"}

def _developer_unlocked():
    if not session.get("portal_logged_in"):
        return False
    unlocked_at = session.get("developer_unlocked_at", 0)
    try:
        return bool(unlocked_at and time.time() - float(unlocked_at) <= DEVELOPER_UNLOCK_SECONDS)
    except (TypeError, ValueError):
        return False

def _developer_file_path(filename, create_parent=False):
                                                                                          
    raw = str(filename or "").replace("\\", "/")
    rel = PurePosixPath(raw)
    if not raw or rel.is_absolute() or any(part in ("", ".", "..") for part in rel.parts):
        raise ValueError("檔案路徑不正確")
    parts = list(rel.parts)
    if any(secure_filename(part) != part for part in parts):
        raise ValueError("檔案或資料夾名稱不正確")
    if os.path.splitext(parts[-1])[1].lower() not in DEVELOPER_ALLOWED_EXTENSIONS:
        raise ValueError("只允許 .py 及 .json 檔案")
    current = BASE_DIR
    for part in parts[:-1]:
        current = os.path.join(current, part)
        if os.path.lexists(current) and os.path.islink(current):
            raise ValueError("不允許存取 symbolic link")
        if create_parent and not os.path.exists(current):
            os.mkdir(current)
        if os.path.exists(current) and not os.path.isdir(current):
            raise ValueError("路徑中的資料夾不存在")
    path = os.path.join(BASE_DIR, *parts)
    if os.path.commonpath([BASE_DIR, os.path.realpath(path)]) != BASE_DIR:
        raise ValueError("不允許存取資料夾以外的檔案")
    if os.path.islink(path):
        raise ValueError("不允許存取 symbolic link")
    return path

@app.route("/developer", methods=["GET"])
def developer():
    if not session.get("portal_logged_in"):
        return redirect(url_for("login"))
    return render_template_string(
        DEVELOPER_TEMPLATE,
        unlocked=_developer_unlocked(),
        error=None,
        websocket_available=bool(sock),
        nvim_available=bool(shutil.which("nvim")),
    )

@app.route("/developer/unlock", methods=["POST"])
def developer_unlock():
    if not session.get("portal_logged_in"):
        return redirect(url_for("login"))
    cfg = load_portal_config()
    supplied = request.form.get("admin_password", "")
    expected = str(cfg.get("admin_password", ""))
    if not expected or not secrets.compare_digest(supplied, expected):
        return render_template_string(
            DEVELOPER_TEMPLATE, unlocked=False, error="❌ admin 密碼錯誤。",
            websocket_available=bool(sock), nvim_available=bool(shutil.which("nvim"))
        ), 401
    session["developer_unlocked_at"] = time.time()
    return redirect(url_for("developer"))

@app.route("/developer/lock", methods=["POST"])
def developer_lock():
    session.pop("developer_unlocked_at", None)
    return jsonify({"ok": True})

@app.route("/developer/files", methods=["GET"])
def developer_files():
    if not _developer_unlocked():
        return jsonify({"error": "Developer 尚未解鎖或已逾時"}), 403
    names = []
    for root, dirs, files in os.walk(BASE_DIR, followlinks=False):
        dirs[:] = [d for d in dirs if not os.path.islink(os.path.join(root, d))]
        for name in files:
            if os.path.splitext(name)[1].lower() not in DEVELOPER_ALLOWED_EXTENSIONS:
                continue
            full = os.path.join(root, name)
            if os.path.islink(full) or not os.path.isfile(full):
                continue
            rel = os.path.relpath(full, BASE_DIR).replace(os.sep, "/")
            names.append(rel)
    return jsonify(sorted(names, key=str.lower))

@app.route("/developer/file/<path:filename>", methods=["GET", "PUT"])
def developer_file(filename):
    if not _developer_unlocked():
        return jsonify({"error": "Developer 尚未解鎖或已逾時"}), 403
    try:
        path = _developer_file_path(filename)
        if request.method == "GET":
            if not os.path.isfile(path):
                return jsonify({"error": "檔案不存在"}), 404
            with open(path, "r", encoding="utf-8-sig") as f:
                return jsonify({"name": os.path.relpath(path, BASE_DIR), "content": f.read()})
        data = request.get_json(silent=True) or {}
        content = data.get("content")
        if not isinstance(content, str):
            return jsonify({"error": "檔案內容格式不正確"}), 400
        if len(content.encode("utf-8")) > 5 * 1024 * 1024:
            return jsonify({"error": "檔案不可大於 5 MB"}), 413
        temp_path = path + ".developer.tmp"
        with open(temp_path, "w", encoding="utf-8", newline="") as f:
            f.write(content)
            f.flush()
            os.fsync(f.fileno())
        os.replace(temp_path, path)
        return jsonify({"ok": True})
    except FileNotFoundError:
        return jsonify({"error": "檔案不存在"}), 404
    except (ValueError, OSError, UnicodeError) as exc:
        return jsonify({"error": str(exc)}), 400

@app.route("/developer/download/<path:filename>", methods=["GET"])
def developer_download(filename):
    if not _developer_unlocked():
        return jsonify({"error": "Developer 尚未解鎖或已逾時"}), 403
    try:
        path = _developer_file_path(filename)
        if not os.path.isfile(path):
            return jsonify({"error": "檔案不存在"}), 404
        return send_file(path, as_attachment=True, download_name=os.path.basename(path))
    except (ValueError, OSError) as exc:
        return jsonify({"error": str(exc)}), 400

@app.route("/developer/upload", methods=["POST"])
def developer_upload():
    if not _developer_unlocked():
        return jsonify({"error": "Developer 尚未解鎖或已逾時"}), 403
    uploaded = request.files.get("file")
    if not uploaded or not uploaded.filename:
        return jsonify({"error": "沒有選擇檔案"}), 400
    try:
        relative_path = request.form.get("relative_path") or uploaded.filename
        path = _developer_file_path(relative_path, create_parent=True)
        if uploaded.content_length and uploaded.content_length > 5 * 1024 * 1024:
            return jsonify({"error": "檔案不可大於 5 MB"}), 413
        conflict = request.args.get("conflict", "check")
        if os.path.exists(path) and conflict != "replace":
            return jsonify({"ok": True, "exists": True, "name": relative_path})
        payload = uploaded.read(5 * 1024 * 1024 + 1)
        if len(payload) > 5 * 1024 * 1024:
            return jsonify({"error": "檔案不可大於 5 MB"}), 413
        payload.decode("utf-8-sig")
        temp_path = path + ".upload.tmp"
        with open(temp_path, "wb") as f:
            f.write(payload)
            f.flush()
            os.fsync(f.fileno())
        os.replace(temp_path, path)
        return jsonify({"ok": True, "exists": False, "name": relative_path})
    except (ValueError, OSError, UnicodeError) as exc:
        return jsonify({"error": str(exc)}), 400

if sock is not None:
    @sock.route("/developer/terminal")
    def developer_terminal(ws):
                                                                              
        if not _developer_unlocked():
            ws.send("\\r\\nDeveloper access expired.\\r\\n")
            return
        nvim_bin = shutil.which("nvim")
        if not nvim_bin:
            ws.send("\\r\\nnvim is not installed on this server.\\r\\n")
            return
        filename = request.args.get("file", "")
        try:
            target = _developer_file_path(filename)
            if not os.path.isfile(target):
                ws.send("\\r\\nFile not found.\\r\\n")
                return
        except (ValueError, OSError):
            ws.send("\\r\\nInvalid file path.\\r\\n")
            return

        import pty
        import select
        import fcntl
        import termios
        import struct as _struct

        pid, master_fd = pty.fork()
        if pid == 0:
            os.chdir(BASE_DIR)
            env = os.environ.copy()
            env["TERM"] = "xterm-256color"
            env["COLORTERM"] = "truecolor"
                                                                           
            nvim_args = [
                nvim_bin,
                "--cmd", "highlight ModeMsg cterm=bold ctermfg=White ctermbg=Blue gui=bold guifg=#ffffff guibg=#2563eb",
                "--cmd", "highlight StatusLine cterm=bold ctermfg=White ctermbg=DarkGreen gui=bold guifg=#ffffff guibg=#0f766e",
                "--cmd", "highlight StatusLineNC ctermfg=White ctermbg=DarkGrey guifg=#e5e7eb guibg=#334155",
                "--cmd", "highlight MsgArea ctermfg=White guifg=#f8fafc",
                "--", target,
            ]
            os.execve(nvim_bin, nvim_args, env)

        try:
                                                                                
            fcntl.ioctl(master_fd, termios.TIOCSWINSZ, _struct.pack("HHHH", 30, 100, 0, 0))
            while True:
                ready, _, _ = select.select([master_fd], [], [], 0.04)
                if ready:
                    try:
                        output = os.read(master_fd, 65536)
                        if not output:
                            break
                        ws.send(output.decode("utf-8", errors="replace"))
                    except OSError:
                        break
                try:
                    incoming = ws.receive(timeout=0.01)
                except Exception:
                    incoming = None
                if incoming is not None:
                    if isinstance(incoming, bytes):
                        incoming = incoming.decode("utf-8", errors="ignore")
                    if incoming.startswith("{"):
                        try:
                            msg = json.loads(incoming)
                            if msg.get("type") == "resize":
                                cols = max(20, min(300, int(msg.get("cols", 100))))
                                rows = max(8, min(100, int(msg.get("rows", 30))))
                                fcntl.ioctl(master_fd, termios.TIOCSWINSZ, _struct.pack("HHHH", rows, cols, 0, 0))
                                os.kill(pid, signal.SIGWINCH)
                                continue
                        except Exception:
                            pass
                    try:
                        os.write(master_fd, incoming.encode("utf-8"))
                    except OSError:
                        break
                done, status = os.waitpid(pid, os.WNOHANG)
                if done == pid:
                    break
        except Exception:
            app.logger.exception("Developer Neovim PTY terminated unexpectedly")
        finally:
            try:
                os.close(master_fd)
            except OSError:
                pass
            try:
                os.kill(pid, signal.SIGHUP)
            except OSError:
                pass
            try:
                os.waitpid(pid, os.WNOHANG)
            except ChildProcessError:
                pass

@app.route("/logout")
def logout():
    session.clear()
    response = redirect(url_for("login"))
    response.delete_cookie("portal_access", path="/")
    return response

class cargoPrefixMiddleware:
                                                       
    def __init__(self, application):
        self.application = application

    def __call__(self, environ, start_response):
        captured = {}
        body_parts = []

        def capture_start_response(status, headers, exc_info=None):
            captured["status"] = status
            captured["headers"] = headers
            captured["exc_info"] = exc_info

        body = self.application(environ, capture_start_response)
        body_parts.extend(body)
        if hasattr(body, "close"):
            body.close()

        headers = captured.get("headers", [])
        content_type = ""
        for k, v in headers:
            if k.lower() == "content-type":
                content_type = v.lower()
                break

        raw = b"".join(
            x if isinstance(x, bytes) else x.encode("utf-8")
            for x in body_parts
        )

        if "text/html" in content_type:
            text = raw.decode("utf-8", errors="replace")
            text = text.replace('href="/', 'href="/cargo/')
            text = text.replace('src="/', 'src="/cargo/')
            text = text.replace('action="/', 'action="/cargo/')
            raw = text.encode("utf-8")

            new_headers = []
            for k, v in headers:
                if k.lower() in ("content-length", "transfer-encoding"):
                    continue
                new_headers.append((k, v))
            new_headers.append(("Content-Length", str(len(raw))))
            headers = new_headers

        start_response(
            captured.get("status", "200 OK"),
            headers,
            captured.get("exc_info"),
        )
        return [raw]

BASE_PORTAL_WSGI_APP = app.wsgi_app

def mount_workshop_apps(settings=None):
                                                                                           
    global cargo, CMHK, cargo_app, cmhk_app
    settings = settings or load_workshop_settings()

    def load_selected(py_name, json_name, default_py, default_json):
        filename = os.path.basename(str(settings.get(py_name, default_py) or default_py))
        if not filename.endswith(".py"):
            filename += ".py"
        module_name = os.path.splitext(filename)[0]
        if not module_name.isidentifier():
            raise ValueError(f"無效的 Python 檔案名稱：{filename}")
        importlib.invalidate_caches()
        module = importlib.import_module(module_name)
        module = importlib.reload(module)
        json_filename = os.path.basename(str(settings.get(json_name, default_json) or default_json))
        data_path = os.path.join(BASE_DIR, json_filename)
        if hasattr(module, "DATA_FILE"):
            module.DATA_FILE = data_path
        if hasattr(module, "DATA_FILE_PATH"):
            module.DATA_FILE_PATH = data_path
        return module

    if settings.get("job1_url"):
        cargo = load_selected("job1_py", "job1_json", "cargo.py", "salary_cargo.json")
    else:
        cargo = load_selected("job1_py", "job1_json", "cargo.py", "salary_cargo.json")
    cargo_app = cargo.app

    if _workshop_worker2_enabled(settings):
        if settings.get("job2_url"):
            CMHK = load_selected("job2_py", "job2_json", "cmhk.py", "salary_cmhk.json")
        else:
            CMHK = load_selected("job2_py", "job2_json", "cmhk.py", "salary_cmhk.json")
        cmhk_app = CMHK.app
    else:
        CMHK = None
        cmhk_app = None

    mounts = {"/cargo": cargoPrefixMiddleware(cargo_app)}
    if cmhk_app is not None:
        mounts["/cmhk"] = cmhk_app
    app.wsgi_app = DispatcherMiddleware(BASE_PORTAL_WSGI_APP, mounts)

mount_workshop_apps()

if __name__ == "__main__":
               
    app.run(host="0.0.0.0", port=8509, debug=True)
