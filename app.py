import os
import re
import base64
import hashlib
import sqlite3
from pathlib import Path
import streamlit as st
import pandas as pd
import time
import matplotlib.pyplot as plt
from docx import Document
from supabase import create_client, Client
import plotly.express as px
import json

os.environ["OPENSSL_CONF"] = ""

# ── 1. CONFIG UTAMA (HARUS PERTAMA) ───────────────────────────────────────────
st.set_page_config(page_title="JivaCare Analytics", layout="wide")

# ── 2. INITIALIZE SESSION STATE ───────────────────────────────────────────────
for _k, _v in {
    "logged_in": False,
    "nama_karyawan": "",
    "jabatan_karyawan": "",
    "username": "",
    "menu": "Dashboard Statistik",
    "edit_mode": False,
    "show_logout_confirm": False,
    "analisis_selesai": False,
    "section_results_all": {},
    "source_type": "lokal"
}.items():
    if _k not in st.session_state:
        st.session_state[_k] = _v

# ── 3. HELPERS & UTILITIES ─────────────────────────────────────────────────────
def hash_password(pw: str) -> str:
    return hashlib.sha256(pw.encode()).hexdigest()

def get_base64(file_path: str) -> str:
    try:
        with open(file_path, "rb") as f:
            return base64.b64encode(f.read()).decode()
    except Exception:
        return ""

def normalize_text(text: str) -> str:
    text = re.sub(r'[^\w\s\.\,\:\;\-\/\(\)\[\]]', ' ', text)
    text = re.sub(r'\s+', ' ', text)
    return text.strip()

def fallback_predict_fields(text: str) -> dict:
    txt_lower = text.lower()
    return {
        "nama_pasien": bool(re.search(r'(\bnama\b|\bpasien\b|\bpxt\b|\btn\.?\s|\bny\.?\s|\ban\.?\s|\bsdr\.?\s|\bidentitas\b|\bisteri\b|\bsuami\b)', txt_lower)),
        "no_rm": bool(re.search(r'(\bno\.?\s?rm\b|\bnorm\b|\brekam\s?medis\b|\bno\.?\s?cm\b|\bregister\b|\breg\b|\bno\s?medis\b)', txt_lower)),
        "diagnosis": bool(re.search(r'(\bdiagnos\w*|\bicd\b|\bdx\b|\bgejala\b|\bkeluhan\b|\bkeadaan\s?umum\b|\banamnes\w*|\bkejiwaan\b|\bstatus\s?mental\b)', txt_lower)),
        "dokter": bool(re.search(r'(\bdr\.?\s|\bdokter\b|\bdpjp\b|\bsp\.?\s?kj\b|\bttd\b|\btanda\s?tangan\b|\bpsikiater\b)', txt_lower)),
        "hasil_lab": bool(re.search(r'(\blab\b|\blaborat\w*|\bdarah\b|\burin\b|\btensi\b|\btd\b|\bnadi\b|\bsuhu\b|\bpsikotes\b|\bpenunjang\b|\bpemeriksaan\b)', txt_lower)),
        "resep_obat": bool(re.search(r'(\bresep\b|\bobat\b|\btherap\w*|\bterapi\b|\bmg\b|\btab\b|\btablet\b|\brx\b|\bkapsul\b|\bpemberian\s?obat\b|\b\d+\s?x\s?\d+\b)', txt_lower)),
        "rujukan": bool(re.search(r'(\brujuk\w*|\bfaskes\b|\bsurat\s?pengantar\b|\basal\b|\bkontrol\b|\bpkm\b|\bpuskesmas\b|\bdr\s?pengirim\b)', txt_lower)),
        "consent": bool(re.search(r'(\bconsent\b|\bsetuju\b|\bpersetujuan\b|\btindakan\b|\bpernyataan\b|\bedukasi\b|\bhak\s?dan\s?kewajiban\b)', txt_lower)),
        "resume": bool(re.search(r'(\bresume\b|\bringkasan\b|\bdischarge\b|\bsummary\b|\bkondisi\s?pulang\b|\bepicrisis\b|\bkeluar\s?rs\b)', txt_lower))
    }

def hitung_skor_kelengkapan(ml_result: dict) -> int:
    target_fields = ["nama_pasien","no_rm","diagnosis","dokter","hasil_lab","resep_obat","rujukan","consent","resume"]
    total_wajib = len(target_fields)
    total_terpenuhi = sum(1 for f in target_fields if ml_result.get(f) is True)
    return 0 if total_wajib == 0 else int((total_terpenuhi / total_wajib) * 100)

def extract_text(file_path: str, tipe_file: str) -> str:
    if tipe_file.lower() == "pdf":
        try:
            from pdfplumber import open as pdf_open
            text = ""
            with pdf_open(file_path) as pdf:
                for page in pdf.pages:
                    t = page.extract_text()
                    if t: text += t + "\n"
            return text
        except Exception:
            return ""
    elif tipe_file.lower() == "docx":
        try:
            doc = Document(file_path)
            return "\n".join([p.text for p in doc.paragraphs])
        except Exception:
            return ""
    return ""

FIELD_LABELS = {
    "nama_pasien": ("👤 Identitas: Nama Pasien", "Memastikan nama pasien tercantum jelas pada berkas klaim."),
    "no_rm": ("🔢 Identitas: Nomor Rekam Medis (No RM)", "Nomor identifikasi unik pasien untuk verifikasi internal."),
    "general_consent": ("✍️ General Consent", "Resume/Discharge Summary"),
    "informed_consent": ("✍️ Informed Consent", "Persetujuan tindakan medis"),
    "diagnosis": ("🩺 Klinis: Diagnosis Penyakit / Kode ICD-10", "Uraian hasil diagnosis gangguan jiwa primer maupun sekunder."),
    "dokter": ("🧑‍⚕️ Administrasi: Dokter Penanggung Jawab (DPJP)", "Nama lengkap beserta gelar dokter yang bertanggung jawab."),
    "hasil_lab": ("🧪 Penunjang: Hasil Laboratorium / Penunjang", "Berkas penunjang medis seperti laboratorium, radiologi, atau psikotes."),
    "resep_obat": ("💊 Terapi: Resep Obat / Lembar Farmasi", "Daftar obat psikiatri yang diberikan selama masa perawatan."),
    "rujukan": ("📄 Regulasi: Surat Rujukan / Pengantar", "Surat rujukan dari faskes pertama atau lembar kontrol berkala."),
    "radiologi": ("🩻 Radiologi", "Hasil imaging / rontgen / CT / MRI"),
    "consent": ("✍️ Legalitas: Informed Consent (Persetujuan)", "Lembar persetujuan tindakan medis yang telah ditandatangani."),
    "sep": ("📝 Administrasi: SEP", "Surat Eligibilitas Peserta BPJS"),
    "laporan_operasi": ("📝 Hasil Operasi", "Laporan hasil tindakan operasi"),
    "resume": ("📝 Ringkasan: Resume Medis (Discharge Summary)", "Lembar ringkasan pelayanan dari awal masuk hingga pasien pulang.")
}

# ── 4. DATABASE & SUPABASE ────────────────────────────────────────────────────
SUPABASE_URL = st.secrets.get("SUPABASE_URL") or os.environ.get("SUPABASE_URL")
SUPABASE_KEY = st.secrets.get("SUPABASE_KEY") or os.environ.get("SUPABASE_KEY")

# 1. Pastikan Anda menambahkan import ClientOptions di baris paling atas file atau di section 4
import httpx
from supabase import create_client, Client, ClientOptions 

@st.cache_resource
def init_supabase():
    if not SUPABASE_URL or not SUPABASE_KEY:
        return None
    
    # 1. Tetap bikin custom HTTP client tanpa HTTP/2
    custom_http_client = httpx.Client(
        http2=False,
        timeout=httpx.Timeout(20.0, read=30.0)
    )
    
    # 2. GANTI http_client MENJADI httpx_client Sesuai saran errornya 🌟
    supabase_options = ClientOptions(
        httpx_client=custom_http_client
    )
    
    # 3. Kirim ke create_client
    return create_client(SUPABASE_URL, SUPABASE_KEY, options=supabase_options)

# Inisialisasi ulang
supabase = init_supabase()

def get_supabase():
    return supabase

@st.cache_data(ttl=60)
def get_accounts() -> pd.DataFrame:
    try:
        conn = sqlite3.connect("eklaim.db")
        df = pd.read_sql("SELECT * FROM pusat_akun", conn)
        conn.close()
        return df.rename(columns={"nama_pengguna": "nama_karyawan"}).astype(str)
    except Exception:
        return pd.DataFrame(columns=["username", "password", "nama_karyawan", "jabatan"])

@st.cache_data(ttl=60)
def load_pasien_supabase(batch_size=50):
    if not supabase:
        return []
    try:
        response = supabase.table("training_dataset").select("*").limit(batch_size).execute()
        return response.data
    except Exception:
        return []

def init_local_db():
    conn = sqlite3.connect("eklaim.db", check_same_thread=False)
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS pusat_akun (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nama_pengguna TEXT, username TEXT UNIQUE, nip TEXT,
            jabatan TEXT, email TEXT, no_hp TEXT, status_akun TEXT, password TEXT
        )
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS berkas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nama_file TEXT UNIQUE,
            path_file TEXT,
            tipe_file TEXT DEFAULT 'docx'
        )
    """)
    conn.commit()
    conn.close()

init_local_db()

# ── 5. HALAMAN LOGIN ──────────────────────────────────────────────────────────
if not st.session_state.logged_in:
    bg = get_base64("meja.jpeg")
    ugmteam_logo = get_base64("logougmteam.png")
    jivacare_logo = get_base64("logojivacare.png")

    st.markdown(f"""
<style>
html, body {{ padding:0!important; margin:0!important; overflow:hidden!important; height:100vh!important; }}
[data-testid="stAppViewContainer"] > .main {{ padding-top:0!important; margin-top:0!important; }}
.block-container {{ padding-top:0!important; margin-top:0!important; }}
section.main {{ padding-top:0!important; overflow:hidden!important; }}
.main .block-container {{ padding:0!important;height:100vh!important;display: flex !important;justify-content: center !important;align-items: center !important; }}
[data-testid="stHeader"] {{ display:none; }}
.stApp {{
    background-image: url("data:image/jpeg;base64,{bg}");
    background-size:cover; background-position:center; background-attachment:fixed;
}}
.stApp::before {{ content:""; position:fixed; top:0; left:0; right:0; bottom:0; background:rgba(0,0,0,0.2); z-index:0; }}
[data-testid="stForm"] {{ background:rgba(255,255,255,0.85); padding:2rem; border-radius:15px; box-shadow:0 10px 25px rgba(0,0,0,0.15); margin: 0px; }}

.login-logo-container {{ display:flex; justify-content:center; align-items:center; margin-bottom:10px; width:100%; }}

.header-logos {{ display:flex; gap:15px; align-items:center; padding:5px 15px; }}
</style>
<div class="header-logos">
    <img src="data:image/png;base64,{ugmteam_logo}" width="250">
</div>
""", unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 1.5, 1])
    with col2:
        with st.form("login_form"):
            st.markdown(f"""
        <div class="login-logo-container">
            <img src="data:image/png;base64,{jivacare_logo}" width="280">
        </div>
        """, unsafe_allow_html=True)
            username = st.text_input("Username")
            password = st.text_input("Password", type="password")
            submitted = st.form_submit_button("Login", use_container_width=True)
        
        st.markdown("""
            <div style='
                position: absolute;
                left: 50%;
                transform: translateX(-50%);
                width: 100%;
                text-align: center;
                margin-top: 15px; 
                color: #ffffff; 
                font-family: "Segoe UI", sans-serif; 
                font-size: 14px; 
                text-shadow: 2px 2px 4px rgba(0,0,0,0.6);
                z-index: 9999;
            '>
                Platform analitik rekam medis yang membantu petugas memvalidasi kelengkapan dan ketepatan berkas administrasi sebelum pengajuan klaim pembiayaan.
            </div>
        """, unsafe_allow_html=True)
        
        if submitted:
            akun = get_accounts()
            hashed_input = hash_password(password)
            user = akun[
                (akun["username"] == username) &
                ((akun["password"] == hashed_input) | (akun["password"] == password))
            ]
            if not user.empty:
                st.session_state.logged_in = True
                st.session_state.username = username
                st.session_state.nama_karyawan = user.iloc[0]["nama_karyawan"]
                st.session_state.jabatan_karyawan = str(user.iloc[0].get("jabatan", "Petugas"))
                st.session_state.menu = "Dashboard Statistik"
                st.rerun()
            else:
                st.error("Username atau password salah")
        st.stop()
        

# ── 6. CORE CSS ───────────────────────────────────────────────────────────────
_m = st.session_state.menu
def _btn_style(page):
    if _m == page:
        return "background:#e0faf9;color:#00B7B3;font-weight:700;border-left:3px solid #00B7B3;"
    return "background:transparent;color:#4a6060;font-weight:500;"

st.markdown("""
<style>
header, [data-testid="stHeader"] { display: none !important; }
[data-testid="stSidebarCollapseButton"] { display: none !important; }
[data-testid="stSidebarNav"] { display: none !important; }

[data-testid="stSidebar"] {
    width: 240px !important; min-width: 240px !important;
    background: #ffffff !important; border-right: 1px solid #e0f5f4 !important;
    box-shadow: 2px 0 12px rgba(0,183,179,0.08) !important; top: 0 !important;
}
[data-testid="stSidebar"] > div:first-child { padding-top: 0 !important; background: #ffffff !important; }

/* ── Rapat antar item sidebar ── */
[data-testid="stSidebar"] [data-testid="stVerticalBlock"] { gap: 0px !important; }
[data-testid="stSidebar"] [data-testid="stVerticalBlock"] > div { margin: 0 !important; padding: 0 !important; }

.jivacare-sidebar-spacer { height: 8px !important; display: block; }

/* ── TOP NAVBAR hijau ── */
.jivacare-top-navbar {
    position: fixed; top: 0; left: 240px; right: 0; height: 60px;
    background: #00B7B3;
    display: flex; align-items: center; padding-left: 28px;
    z-index: 998;
    box-shadow: 0 2px 10px rgba(0,150,147,0.35);
}
.jivacare-page-title {
    font-size: 19px !important; font-weight: 800 !important;
    color: #ffffff !important; letter-spacing: 0.2px;
    text-shadow: 0 1px 2px rgba(0,0,0,0.1);
}

.main .block-container {
    padding-top: 86px !important; padding-left: 28px !important;
    padding-right: 28px !important; max-width: 100% !important;
}

[data-testid="stSidebar"] button {
    width: 100% !important; border: none !important; border-radius: 6px !important;
    padding: 7px 14px !important; text-align: left !important; font-size: 13px !important;
    cursor: pointer !important; box-shadow: none !important; margin: 1px 0 !important;
    transition: background 0.15s !important;
}
[data-testid="stSidebar"] button:hover { background: #f0fafa !important; color: #00B7B3 !important; }
[data-testid="stSidebar"] button p { font-size: 13px !important; margin: 0 !important; }
[data-testid="stSidebar"] hr { margin: 5px 14px 3px !important; border-color: #eef8f7 !important; }
[data-testid="stSidebar"] .stMarkdown p {
    font-size: 9px !important; font-weight: 700 !important; color: #bbb !important;
    letter-spacing: 1.1px !important; text-transform: uppercase !important;
    padding: 5px 14px 1px !important; margin: 0 !important;
}

/* ── LOGOUT MODAL OVERLAY ── */
.jc-logout-overlay {
    position: fixed; inset: 0;
    background: rgba(0,0,0,0.52);
    z-index: 99999;
    backdrop-filter: blur(5px);
    -webkit-backdrop-filter: blur(5px);
    display: flex; align-items: center; justify-content: center;
    animation: jcFadeIn 0.18s ease;
}
@keyframes jcFadeIn { from { opacity:0; } to { opacity:1; } }
.jc-modal {
    background: #fff;
    border-radius: 18px;
    padding: 40px 44px 36px;
    width: 400px; max-width: 90vw;
    box-shadow: 0 24px 60px rgba(0,0,0,0.30);
    text-align: center;
    animation: jcSlideUp 0.22s cubic-bezier(0.34,1.5,0.64,1);
}
@keyframes jcSlideUp {
    from { opacity:0; transform:translateY(26px) scale(0.96); }
    to   { opacity:1; transform:translateY(0)    scale(1);    }
}
.jc-modal-icon  { font-size:46px; margin-bottom:12px; display:block; }
.jc-modal-title { font-size:20px; font-weight:800; color:#1a3a3a; margin-bottom:8px; }
.jc-modal-desc  { font-size:14px; color:#666; margin-bottom:28px; line-height:1.6; }
.jc-modal-btns  { display:flex; gap:10px; }
.jc-btn-yes {
    flex:1; background:#e05555; color:#fff; border:none; border-radius:9px;
    padding:11px 0; font-size:14px; font-weight:700; cursor:pointer;
    transition:background 0.15s, transform 0.1s;
}
.jc-btn-yes:hover  { background:#c0392b; transform:translateY(-1px); }
.jc-btn-no {
    flex:1; background:#f0f0f0; color:#444; border:none; border-radius:9px;
    padding:11px 0; font-size:14px; font-weight:600; cursor:pointer;
    transition:background 0.15s, transform 0.1s;
}
.jc-btn-no:hover { background:#e2e2e2; transform:translateY(-1px); }
/* hide the Streamlit buttons used for logout action */
.logout-st-btns { display:none; }

/* ── Paksa sidebar spacing rapat — target semua level wrapper ── */
[data-testid="stSidebar"] section[data-testid="stSidebarContent"] { padding: 0 !important; }
[data-testid="stSidebar"] .block-container { padding: 0 !important; }
[data-testid="stSidebar"] .element-container { margin: 0 !important; padding: 0 !important; }
[data-testid="stSidebar"] .stButton { margin: 0 !important; padding: 0 2px !important; }
[data-testid="stSidebar"] .stButton > button { margin: 0 !important; }
[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] { margin: 0 !important; padding: 0 !important; }
[data-testid="stSidebar"] div[data-testid="stHorizontalBlock"] { gap: 0 !important; }
</style>
""", unsafe_allow_html=True)

# ── 7. RENDER SIDEBAR ─────────────────────────────────────────────────────────
# ── 7. RENDER SIDEBAR & POP-UP LOGOUT ─────────────────────────────────────────
jivacare_logo = get_base64("logojivacare.png")

# 🌟 1. POP-UP MODAL LOGOUT YANG CANTIK (TIDAK MERUSAK LAYOUT)
@st.dialog("Konfirmasi Keluar")
def konfirmasi_logout_modal():
    st.write("Apakah Anda yakin ingin keluar dari aplikasi JivaCare Analytics?")
    st.markdown("<br>", unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    with col1:
        if st.button("✅ Ya, Keluar", use_container_width=True, type="primary", key="modal_yes_logout"):
            st.session_state.clear()
            st.rerun()
    with col2:
        if st.button("❌ Batalkan", use_container_width=True, key="modal_no_logout"):
            st.rerun()

# 🌟 2. SIDEBAR UTAMA (TIDAK ADA TOMBOL LOGOUT GANDA LAGI)
with st.sidebar:
    st.markdown(f"""
<style>
.jc-user-box    {{ padding:10px 16px 8px; background:#f8fffe; border-bottom:1px solid #e8f8f7; }}
.jc-user-name   {{ font-size:16px; font-weight:700; color:#1a3a3a; margin-bottom:1px; }}
.jc-user-role   {{ font-size:13px; color:#00B7B3; font-weight:500; }}
</style>

<div style="display:flex; justify-content:center; align-items:center; height:80px; margin-bottom:10px; background:#ffffff;">
    <img src="data:image/png;base64,{jivacare_logo}" style="width:200px; height:60px; object-fit:cover;"/>
</div>
<div class="jc-user-box">
    <div class="jc-user-name">{st.session_state.get("nama_karyawan","Nama")}</div>
    <div class="jc-user-role">{st.session_state.get("jabatan_karyawan","Jabatan")}</div>
</div>
<div class="jivacare-sidebar-spacer"></div>
""", unsafe_allow_html=True)

    st.markdown("MENU UTAMA")
    def _nav(label, page):
        style = _btn_style(page)
        st.markdown(f"<style>button[data-testid='baseButton-secondary'][key='nav_{page}'] {{ {style} }}</style>", unsafe_allow_html=True)
        if st.button(label, key=f"nav_{page}", use_container_width=True):
            st.session_state.menu = page
            st.rerun()

    _nav("📊 Dashboard Statistik", "Dashboard Statistik")
    _nav("🗂️ Validasi Berkas", "Validasi Berkas")
    _nav("💰 Daftar Tarif INA-CBG", "Daftar Tarif INA-CBG")
    st.divider()
    st.markdown("SISTEM")
    _nav("⚙️ Pengaturan", "Pengaturan")
    _nav("❓ Bantuan", "Bantuan")
    st.divider()
    _nav("👤 Pusat Akun", "Pusat Akun")

    # Tombol logout tunggal yang aman dan stylish
    st.markdown("<style>button[key='nav_Logout'] p { color:#e05555 !important; }</style>", unsafe_allow_html=True)
    if st.button("🚪 Logout Aplikasi", key="nav_Logout", use_container_width=True):
        konfirmasi_logout_modal()

# ── TOP NAVBAR ────────────────────────────────────────────────────────────────
_menu_display = st.session_state.menu
st.markdown("""
<style>
.jivacare-page-title {font-size: 30px !important; font-weight: bold !important;}
</style>
""", unsafe_allow_html=True)

_menu_display = st.session_state.menu
st.markdown(f"""
<div class="jivacare-top-navbar">
    <div class="jivacare-page-title">{_menu_display}</div>
</div>
""", unsafe_allow_html=True)

@st.cache_data(ttl=10)
def ambil_data_validasi_aman():
    if not supabase:
        return []
    try:
        # Paksa SSL Context lokal ke tingkat SECLEVEL 1 sebelum eksekusi query
        import ssl
        try:
            ssl._create_default_https_context = ssl._create_unverified_context
        except Exception:
            pass
            
        res = supabase.table("hasil_validasi").select("*").execute()
        return res.data if res.data else []
    except Exception:
        # Jika terkena bad record mac akibat gangguan thread di Windows, coba bypass sekali lagi
        time.sleep(0.5)
        try:
            res = supabase.table("hasil_validasi").select("*").execute()
            return res.data if res.data else []
        except Exception:
            return []

# ── 9. ROUTING HALAMAN ────────────────────────────────────────────────────────
menu = st.session_state.menu

# ── 📊 DASHBOARD STATISTIK ────────────────────────────────────────────────────
if menu == "Dashboard Statistik":
    # Ganti pemanggilan data mentah kemarin dengan fungsi aman kita yang baru:
    data_raw = ambil_data_validasi_aman()
    
    if data_raw:
        df_dashboard = pd.DataFrame(data_raw)
# ── 9. ROUTING HALAMAN ────────────────────────────────────────────────────────
menu = st.session_state.menu

# ── 📊 DASHBOARD STATISTIK ────────────────────────────────────────────────────
if menu == "Dashboard Statistik":
    # 🌟 KITA HAPUS TRIK CSS MARGIN NEGATIF YANG BIKIN BRUTAL KEMARIN 🌟
    
    # Ambil data utama dari Supabase
    res = supabase.table("hasil_validasi").select("*").execute()
    if res.data:
        df_dashboard = pd.DataFrame(res.data)

        if 'hasil_json' in df_dashboard.columns:
            hasil_df = pd.json_normalize(df_dashboard['hasil_json'].apply(lambda x: x if isinstance(x, dict) else {}))
            hasil_df = hasil_df.add_prefix("json_")
            df_dashboard = pd.concat([df_dashboard.drop(columns=['hasil_json']), hasil_df], axis=1)

        if 'json_skor_kelengkapan' not in df_dashboard.columns:
            df_dashboard['json_skor_kelengkapan'] = 0

        hijau  = df_dashboard[df_dashboard['json_skor_kelengkapan'] >= 80].shape[0]
        kuning = df_dashboard[(df_dashboard['json_skor_kelengkapan'] >= 50) & (df_dashboard['json_skor_kelengkapan'] < 80)].shape[0]
        merah  = df_dashboard[df_dashboard['json_skor_kelengkapan'] < 50].shape[0]
        total  = df_dashboard.shape[0]
        
        card_style = """
        border-radius: 12px;
        padding: 18px;
        text-align: center;
        box-shadow: 0 6px 15px rgba(0,0,0,0.08);
        font-family: 'Segoe UI', sans-serif;
        margin-bottom: 12px;
        """
    else:
        hijau, kuning, merah, total = 0, 0, 0, 0
        st.write("Belum ada data untuk dashboard")

    # ── 1. BARIS PERTAMA: PIE CHART (KIRI) vs CARDS JUMLAH (KANAN) ──
    main_col1, main_col2 = st.columns([1.5, 2])

    # KIRI: Tempat Pie Chart berada
    with main_col1:
        if total > 0:
            with st.container(border=True):
                st.markdown("<p style='text-align:center;font-weight:bold;font-size:15px;margin-bottom:5px;color:#1a3a3a;'>Proporsi Status Kelayakan Berkas</p>", unsafe_allow_html=True)
                labels = ['Layak Klaim', 'Perlu Peninjauan', 'Tidak Layak Klaim']
                sizes = [hijau, kuning, merah]
                colors = ['#27ae60', '#f1c40f', '#e74c3c']

                if sum(sizes) > 0:
                    fig, ax = plt.subplots(figsize=(4, 3.2)) # Sedikit dikecilkan agar pas tingginya
                    wedges, texts, autotexts = ax.pie(
                        sizes, labels=labels, autopct='%1.1f%%', 
                        startangle=140, colors=colors, 
                        textprops=dict(color="black", size=9)
                    )
                    plt.setp(autotexts, size=8, weight="bold")
                    ax.axis('equal')  
                    fig.patch.set_alpha(0.0)
                    st.pyplot(fig)

    with main_col2:
        if total > 0:
            with st.container():
                box_status_style = """
                border-radius: 12px;
                padding: 15px 10px;
                height: 105px; /* Mengunci tinggi agar 3 kotak sama rata presisi */
                text-align: center;
                box-shadow: 0 6px 15px rgba(0,0,0,0.08);
                font-family: 'Segoe UI', sans-serif;
                display: flex;
                flex-direction: column;
                justify-content: center;
                align-items: center;
                """
                
                sub_col1, sub_col2, sub_col3 = st.columns(3)
                sub_col1.markdown(f"<div style='background-color:#b6fcb6; color:#0e4b0e; {box_status_style}'>Layak Klaim<br><b style='font-size:22px;'>{hijau}</b></div>", unsafe_allow_html=True)
                sub_col2.markdown(f"<div style='background-color:#fff3b6; color:#5c4d00; {box_status_style}'>Perlu Peninjauan<br><b style='font-size:22px;'>{kuning}</b></div>", unsafe_allow_html=True)
                sub_col3.markdown(f"<div style='background-color:#fcb6b6; color:#5c0e0e; {box_status_style}'>Tidak Layak Klaim<br><b style='font-size:22px;'>{merah}</b></div>", unsafe_allow_html=True)
                
                # 🌟 DIBAWAHIN DIKIT: Menambahkan margin-top sebesar 20px agar kotak abu-abu turun dan punya jarak
                total_card_style = """
                border-radius: 12px;
                padding: 30px 18px; 
                margin-top: 20px; 
                text-align: center;
                box-shadow: 0 6px 15px rgba(0,0,0,0.08);
                font-family: 'Segoe UI', sans-serif;
                background-color:#e05555; 
                color:#e05555;
                """
                st.markdown(f"<div style='{total_card_style}'>Total Berkas Keseluruhan<br><b style='font-size:32px; line-height: 1.4;'>{total}</b></div>", unsafe_allow_html=True)
    # ── 2. BARIS KEDUA: BAR CHART MEMANJANG FULL KIRI KANAN ──
    if total > 0:
        st.markdown("<br>", unsafe_allow_html=True)
        with st.container(border=True):
            st.markdown("<p style='text-align:center;font-weight:bold;font-size:15px;margin-bottom:10px;color:#1a3a3a;'>Grafik Tingkat Kepatuhan Per Jenis Dokumen</p>", unsafe_allow_html=True)
            
            # 🌟 FUNGSI HITUNG PINTAR: Mencari kolom asli json_ atau teks bersih secara akurat
            def hitung_persen(nama_kolom, df):
                # Cari kemungkinan nama kolom dengan awalan 'json_' atau nama asli langsung
                kolom_ketemu = None
                target_nama = str(nama_kolom).strip().lower()
                
                for col in df.columns:
                    col_clean = str(col).strip().lower()
                    if col_clean == target_nama or col_clean == f"json_{target_nama}":
                        kolom_ketemu = col
                        break
                
                # Jika kolom ditemukan di dalam DataFrame
                if kolom_ketemu is not None:
                    # Ubah paksa data kolom menjadi angka biner, jika gagal bypass jadi NaN
                    seri_angka = pd.to_numeric(df[kolom_ketemu], errors='coerce')
                    # Hitung baris yang isinya angka 1 (Patuh)
                    jumlah_patuh = (seri_angka == 1).sum()
                    return round((jumlah_patuh / total) * 100, 1)
                
                return 0.0

            # Struktur DataFrame tetap menggunakan cetakan asli milikmu 🌟
            df_kepatuhan = pd.DataFrame({
                "Dokumen": ["Identitas Pasien", "No Rekam Medis", "Nama Dokter", "Resume Medis", "General Consent", "Surat Rujukan", "Diagnosis", "Hasil Lab", "Resep Obat"],
                "Kepatuhan (%)": [
                    hitung_persen("nama_pasien", df_dashboard),
                    hitung_persen("no_rm", df_dashboard),
                    hitung_persen("dokter", df_dashboard),
                    hitung_persen("resume", df_dashboard),
                    hitung_persen("consent", df_dashboard),
                    hitung_persen("rujukan", df_dashboard),
                    hitung_persen("diagnosis", df_dashboard),
                    hitung_persen("hasil_lab", df_dashboard),
                    hitung_persen("resep_obat", df_dashboard)
                ]
            })
            
            import plotly.express as px
            fig_bar = px.bar(
                df_kepatuhan, x="Dokumen", y="Kepatuhan (%)", text="Kepatuhan (%)",
                color="Kepatuhan (%)", color_continuous_scale="Tealgrn"
            )
            fig_bar.update_layout(
                yaxis=dict(range=[0, 105]), 
                uniformtext_minsize=12, 
                uniformtext_mode='hide',
                margin=dict(l=20, r=20, t=10, b=20),
                height=320 
            )
            # Tambahin format % di ujung teks biar informatif
            fig_bar.update_traces(texttemplate='%{text}%', textposition='outside')
            st.plotly_chart(fig_bar, use_container_width=True)

    # ── 3. BARIS KETIGA: DATAFRAME UTAMA DI PALING BAWAH + TOMBOL REFRESH AMAN ──
    st.markdown("---")
    
    # Pindahkan tombol refresh ke bagian subheader tabel biar rapi dan presisi
    tabel_col1, tabel_col2 = st.columns([4, 1])
    with tabel_col1:
        st.subheader("📋 Riwayat Hasil Validasi")
    with tabel_col2:
        if st.button("🔄 Refresh Data", type="secondary", use_container_width=True, key="btn_refresh_tabel"):
            st.rerun()

    try:
        if res.data:
            df = pd.DataFrame(res.data)
            if 'hasil_json' in df.columns:
                hasil_df = pd.json_normalize(df['hasil_json'].apply(lambda x: x if isinstance(x, dict) else {}))
                hasil_df = hasil_df.add_prefix("json_")
                df = pd.concat([df.drop(columns=['hasil_json']), hasil_df], axis=1)
                
            kolom_hapus = ['id','no_rm','skor_kelengkapan','status_validasi','created_at']
            df = df.drop(columns=[c for c in kolom_hapus if c in df.columns])
            df = df.rename(columns={
                'nama_pasien':'Nama Berkas',
                'json_no_rm':'No Rekam Medis',
                'json_dokter':'Nama Dokter',
                'json_resume':'Resume Medis',
                'json_consent':'General Consent',
                'json_rujukan':'Surat Rujukan',
                'json_diagnosis':'Diagnosis',
                'json_hasil_lab':'Hasil Lab',
                'json_resep_obat':'Resep Obat',
                'json_nama_pasien':'Identitas Pasien',
                'json_skor_kelengkapan':'Skor Kelengkapan (%)'
            })
            st.dataframe(df, use_container_width=True)
        else:
            st.write("Belum ada data hasil validasi")
    except Exception as e:
        st.error(f"Terjadi error saat memproses tabel: {e}")

# ── 🗂️ VALIDASI BERKAS ────────────────────────────────────────────────────────
elif menu == "Validasi Berkas":
    try:
        from train_transformer_validator import predict_fields, normalize_text, extract_text, hitung_skor_kelengkapan, get_supabase
        HAS_TRANSFORMER_MODEL = True
    except Exception:
        HAS_TRANSFORMER_MODEL = False
        from train_transformer_validator import normalize_text, extract_text, hitung_skor_kelengkapan, get_supabase

    import uuid
    sb = get_supabase()

    st.markdown("## Upload Dokumen Klaim")
    st.caption("Unggah berkas rekam medis lokal Anda di sini untuk diprediksi kelengkapannya oleh sistem.")

    UPLOAD_FOLDER = "uploads"
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)

    uploaded_files = st.file_uploader(
        "Pilih Berkas Rekam Medis (PDF/DOCX)",
        type=["pdf","docx"],
        accept_multiple_files=True,
        key="local_upload"
    )

    if st.button("🚀 Jalankan Validasi Data Test", type="primary", key="btn_val_local"):
        if not uploaded_files:
            st.warning("Silakan unggah dokumen terlebih dahulu!")
        else:
            progress_bar = st.progress(0)
            temp_storage = {}
            total_len = len(uploaded_files)
            conn = sqlite3.connect("eklaim.db")
            c = conn.cursor()
            for idx, file in enumerate(uploaded_files):
                safe_name = Path(file.name).name
                file_path = os.path.join(UPLOAD_FOLDER, safe_name)
                with open(file_path, "wb") as f:
                    f.write(file.getbuffer())
                tipe_file = safe_name.split(".")[-1].lower()
                c.execute("INSERT OR REPLACE INTO berkas (nama_file, path_file, tipe_file) VALUES (?,?,?)", (safe_name, file_path, tipe_file))
                conn.commit()
                raw_text     = extract_text(file_path, tipe_file)
                cleaned_text = normalize_text(raw_text)
                ml_result = {}
                if HAS_TRANSFORMER_MODEL:
                    try:
                        ml_result = predict_fields(cleaned_text)
                    except Exception:
                        ml_result = fallback_predict_fields(cleaned_text)
                else:
                    ml_result = fallback_predict_fields(cleaned_text)
                ml_result["skor_kelengkapan"] = hitung_skor_kelengkapan(ml_result)
                temp_storage[safe_name] = {
                    "nama_pasien": safe_name.split(".")[0],
                    "fields": ml_result,
                    "raw_text": raw_text
                }
                progress_bar.progress((idx + 1) / total_len)
            conn.close()
            st.session_state.section_results_all = temp_storage
            st.session_state.source_type = "lokal"
            st.session_state.analisis_selesai = True
            progress_bar.empty()
            st.success("Analisis dokumen lokal selesai!")

    if st.session_state.analisis_selesai and st.session_state.section_results_all:
        st.markdown("---")
        st.subheader("Hasil Validasi")

        html_rows = ""
        for pid, data in st.session_state.section_results_all.items():
            f = data["fields"]
            skor_aktif = f.get("skor_kelengkapan", 0)
            if skor_aktif >= 80:
                status_label, warna = "Layak Klaim", "#b6fcb6"
            elif skor_aktif >= 50:
                status_label, warna = "Perlu Peninjauan", "#fff3b6"
            else:
                status_label, warna = "Tidak Layak Klaim", "#fcb6b6"
            html_rows += f"""
            <tr>
                <td>{pid}</td><td>{data['nama_pasien']}</td>
                <td>{skor_aktif}%</td>
                <td style="background-color:{warna};font-weight:bold;text-align:center">{status_label}</td>
            </tr>"""

        html_table = f"""
        <table style="width:100%;border-collapse:collapse;">
            <thead><tr><th>ID / Nama Berkas</th><th>Nama Pasien</th><th>Kelengkapan (%)</th><th>Status Klaim</th></tr></thead>
            <tbody>{html_rows}</tbody>
        </table>"""
        total = len(st.session_state.section_results_all)
        height = min(max(200, total * 35), 600)
        import streamlit.components.v1 as components
        components.html(f"<div style='overflow:auto;'>{html_table}</div>", height=height)

        st.markdown("---")
        st.subheader("Detail Kelengkapan")
        for pid, data in st.session_state.section_results_all.items():
            f = data["fields"]
            skor_pct = f.get("skor_kelengkapan", 0)
            with st.expander(f"👤 {data['nama_pasien']} (ID/File: {pid}) — Capaian Kelengkapan: {skor_pct}%"):
                st.metric("Persentase Kepatuhan Berkas", f"{skor_pct}%")
                st.markdown("Berikut adalah rincian hasil validasi otomatis sistem per dokumen:")
                labels_mapping = FIELD_LABELS
                for field_key, (label, deskripsi) in labels_mapping.items():
                    terdeteksi = f.get(field_key, False)
                    with st.container(border=True):
                        c_title, c_status = st.columns([4, 1])
                        with c_title:
                            st.markdown(f"**{label}**")
                            st.caption(deskripsi)
                        with c_status:
                            if terdeteksi:
                                st.markdown("<span style='color:#27ae60;font-weight:bold;background:#e8f8f5;padding:4px 10px;border-radius:5px;display:inline-block;'>✅ LENGKAP</span>", unsafe_allow_html=True)
                            else:
                                st.markdown("<span style='color:#c0392b;font-weight:bold;background:#fce4d6;padding:4px 10px;border-radius:5px;display:inline-block;'>❌ TIDAK ADA</span>", unsafe_allow_html=True)
                        if terdeteksi:
                            st.markdown(f"🤖 *Hasil Analisis:* Komponen kalimat penunjang terkait `{field_key}` sukses diverifikasi dalam teks rekam medis.")
                        else:
                            clean_label = label.split(': ')[1] if ': ' in label else label
                            st.markdown(f"⚠️ *Rekomendasi:* Bagian **{clean_label}** tidak ditemukan. Mohon lengkapi data fisik atau input teks sebelum diajukan.")

        if st.session_state.source_type == "lokal":
            st.markdown("---")
            if st.button("💾 Simpan Hasil Analisis", use_container_width=True, type="primary"):
                if sb is None:
                    st.error("❌ Supabase URL/KEY belum dikonfigurasi.")
                else:
                    with st.spinner("Mengunggah berkas baru ke dalam database..."):
                        fail_counter = 0
                        success_counter = 0
                        for filename, data in st.session_state.section_results_all.items():
                            try:
                                payload = {
                                    "nama_pasien": str(data["nama_pasien"]),
                                    "no_rm": str(filename.split(".")[0][:10]),
                                    "hasil_json": dict(data["fields"])
                                }
                                sb.table("hasil_validasi").insert(payload).execute()
                                success_counter += 1
                            except Exception as e:
                                fail_counter += 1
                                st.error(f"⚠️ Gagal upload {filename}")
                                with st.expander("Lihat Detail Error"):
                                    st.code(str(e))
                        if success_counter > 0:
                            st.success(f"🎉 Berhasil menyimpan {success_counter} data database!")
                            st.session_state.analisis_selesai = False
                            if "upload_file" in st.session_state:
                                st.session_state.upload_file = None
                        else:
                            st.warning(f"Proses selesai dengan {fail_counter} berkas gagal dianalsis")


# ── 💰 DAFTAR TARIF INA-CBG ───────────────────────────────────────────────────
elif menu == "Daftar Tarif INA-CBG":
    try:
        st.dataframe(pd.read_excel("database_ina_cbg.xlsx"), use_container_width=True, hide_index=True)
    except Exception as e:
        st.error(f"Berkas database_ina_cbg.xlsx tidak ditemukan: {e}")


# ── 👤 PUSAT AKUN ─────────────────────────────────────────────────────────────
elif menu == "Pusat Akun":
    conn_pa = sqlite3.connect("eklaim.db", check_same_thread=False)
    try:
        df_pa = pd.read_sql("SELECT * FROM pusat_akun", conn_pa)
        user_row = df_pa[df_pa["username"] == st.session_state.username]
        if user_row.empty:
            st.error("Data pengguna tidak ditemukan.")
        else:
            user = user_row.iloc[0]
            if st.session_state.edit_mode:
                st.subheader("✏️ Edit Profil")
                with st.form("edit_form"):
                    nama    = st.text_input("Nama Pengguna", value=user["nama_pengguna"])
                    nip     = st.text_input("NIP", value=user["nip"])
                    jab_opt = ["Verifikator RM", "Casemix", "IT"]
                    jabatan = st.selectbox("Jabatan", jab_opt, index=jab_opt.index(user["jabatan"]) if user["jabatan"] in jab_opt else 0)
                    email   = st.text_input("Email", value=user["email"])
                    no_hp   = st.text_input("Nomor HP", value=user["no_hp"])
                    c1, c2 = st.columns(2)
                    with c1: simpan = st.form_submit_button("💾 Simpan", use_container_width=True)
                    with c2: batal  = st.form_submit_button("❌ Batal",  use_container_width=True)
                if simpan:
                    try:
                        cursor_pa = conn_pa.cursor()
                        cursor_pa.execute("""
                            UPDATE pusat_akun SET nama_pengguna=?, nip=?, jabatan=?, email=?, no_hp=? WHERE username=?
                        """, (nama, nip, jabatan, email, no_hp, st.session_state.username))
                        conn_pa.commit()
                        st.session_state.nama_karyawan    = nama
                        st.session_state.jabatan_karyawan = jabatan
                        st.session_state.edit_mode        = False
                        st.success("Profil Berhasil Diperbarui!")
                        st.rerun()
                    except Exception as ex:
                        st.error(f"Gagal memperbarui database: {ex}")
                if batal:
                    st.session_state.edit_mode = False
                    st.rerun()
            else:
                st.subheader("👤 Profil Pengguna")
                with st.container(border=True):
                    st.write(f"**Nama Karyawan:** {user['nama_pengguna']}")
                    st.write(f"**NIP:** {user['nip']}")
                    st.write(f"**Jabatan:** {user['jabatan']}")
                    st.write(f"**Email:** {user['email']}")
                    st.write(f"**No HP:** {user['no_hp']}")
                if st.button("✏️ Ubah Profil", type="primary"):
                    st.session_state.edit_mode = True
                    st.rerun()
    finally:
        conn_pa.close()


# ── ⚙️ PENGATURAN ─────────────────────────────────────────────────────────────
elif menu == "Pengaturan":
    _kategori = {
        "🖥️ Tampilan": ["Light/Dark Mode","Ukuran Tampilan","Bahasa Sistem","Format Tanggal dan Waktu","Jumlah Data per Halaman"],
        "🔒 Privasi dan Keamanan": ["Riwayat Login","Aktivitas Akun","Hak Akses dan Izin Pengguna","Perizinan Aplikasi"],
        "📂 Master Data": ["Master ICD-10 Gangguan Jiwa","Master ICD-9 CM Tindakan","Master Obat Psikiatri","Master Dokter DPJP","Master Poli/Ruangan"],
        "💬 Feedback Sistem": ["Laporkan Bug","Saran Fitur","Kendala Coding"],
    }
    for judul, items in _kategori.items():
        with st.expander(judul, expanded=True):
            cols = st.columns(2)
            for i, item in enumerate(items):
                with cols[i % 2]:
                    st.markdown(
                        f'<div style="padding:10px 14px;background:#f8fffe;border:1px solid #d0f0ee;'
                        f'border-radius:8px;margin-bottom:8px;font-size:14px;color:#333;">{item}</div>',
                        unsafe_allow_html=True
                    )


# ── ❓ BANTUAN ─────────────────────────────────────────────────────────────────
elif menu == "Bantuan":
    try:
        df_ban = pd.read_excel("psaksetban.xlsx", sheet_name="Bantuan", header=None)
        df_ban.columns = ["Kategori","Pertanyaan","Jawaban"]
        df_ban["Kategori"] = df_ban["Kategori"].ffill()
        for kategori, grp in df_ban.groupby("Kategori", sort=False):
            with st.expander(f"📂 {kategori}", expanded=False):
                for _, row in grp.iterrows():
                    with st.expander(f"❓ {row['Pertanyaan']}", expanded=False):
                        st.markdown(
                            f"<div style='padding:12px 16px;background:#f8fffe;border-left:3px solid #00B7B3;"
                            f"border-radius:0 8px 8px 0;font-size:14px;color:#333;line-height:1.6;'>"
                            f"{row['Jawaban']}</div>",
                            unsafe_allow_html=True
                        )
    except Exception as e:
        st.error(f"Gagal memuat dokumen bantuan: {e}")