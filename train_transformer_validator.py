# train_transformer_validator.py
from supabase import create_client
from transformers import AutoTokenizer, AutoModelForTokenClassification, AutoModelForSequenceClassification, pipeline
import torch
import os
from dotenv import load_dotenv
import re
from pathlib import Path
from pypdf import PdfReader
import docx

load_dotenv()

# ===== Supabase config =====
SUPABASE_URL = os.environ.get("SUPABASE_URL")
SUPABASE_KEY = os.environ.get("SUPABASE_KEY")
supabase = None
import httpx # Pastikan ada import httpx di bagian paling atas file jika belum ada

import httpx
from supabase import create_client, ClientOptions  # <-- Pastikan ClientOptions di-import di sini

def get_supabase():
    global supabase
    if supabase is None:
        if not SUPABASE_URL or not SUPABASE_KEY:
            return None
        
        # Buat custom HTTP client untuk bypass error SSL/HTTP2 kemarin
        custom_client = httpx.Client(http2=False)
        
        opts = ClientOptions(httpx_client=custom_client)
        
        # Masukkan opts ke dalam create_client
        supabase = create_client(SUPABASE_URL, SUPABASE_KEY, options=opts)
        
    return supabase

# ===== Model Paths =====
NER_MODEL_PATH = "C:/Users/HP/dashboardpjd/model_ml"
CLS_MODEL_PATH = "C:/Users/HP/dashboardpjd/model_ml"

# ===== Load ML Models =====
def load_models(ner_model_path=NER_MODEL_PATH, cls_model_path=CLS_MODEL_PATH):
    tokenizer_ner = AutoTokenizer.from_pretrained(ner_model_path)
    model_ner = AutoModelForTokenClassification.from_pretrained(ner_model_path)
    ner_pipeline_obj = pipeline(
        "ner",
        model=model_ner,
        tokenizer=tokenizer_ner,
        aggregation_strategy="simple"
    )
    tokenizer_cls = AutoTokenizer.from_pretrained(cls_model_path)
    model_cls = AutoModelForSequenceClassification.from_pretrained(cls_model_path)
    return tokenizer_cls, model_cls, ner_pipeline_obj

try:
    tokenizer_cls, model_cls, ner_pipeline = load_models()
except Exception as e:
    print(f"⚠️ Model ML tidak terdeteksi, fallback regex digunakan: {e}")

# ===== EKSTRAKSI TEKS =====
import urllib.request

def extract_text(path_file, tipe_file):
    """
    Fungsi cerdas: Mendukung link URL internet (Supabase Storage) & file lokal,
    serta membaca PDF/DOCX secara akurat menggunakan library aslinya.
    """
    # 1. Penanganan jika file ditaruh di Cloud/Internet (untuk keperluan Deploy)
    if path_file.startswith("http://") or path_file.startswith("https://"):
        try:
            print(f"🌐 Mengunduh berkas dari cloud storage...")
            temp_file, _ = urllib.request.urlretrieve(path_file)
            path_to_read = temp_file
        except Exception as e:
            print(f"❌ Gagal mengunduh file dari URL: {e}")
            return ""
    else:
        path_to_read = path_file

    # 2. Proses Ekstraksi Teks Asli sesuai Tipe File
    text = ""
    if os.path.exists(path_to_read):
        try:
            # Jika tipenya PDF
            if tipe_file.lower() == 'pdf':
                reader = PdfReader(path_to_read)
                for page in reader.pages:
                    page_text = page.extract_text()
                    if page_text:
                        text += page_text + "\n"
                        
            # Jika tipenya Word (DOCX)
            elif tipe_file.lower() in ['docx', 'doc']:
                doc = docx.Document(path_to_read)
                text = "\n".join([para.text for para in doc.paragraphs])
                
            # Jika tipenya teks biasa
            else:
                with open(path_to_read, 'r', encoding='utf-8', errors='ignore') as f:
                    text = f.read()
                    
            return text if text.strip() else "Berkas kosong"
            
        except Exception as e:
            print(f"⚠️ Gagal membaca konten file {path_to_read}: {e}")
            return "Gagal ekstraksi teks berkas"
            
    print(f"⚠️ File tidak ditemukan di: {path_to_read}")
    return "File tidak ditemukan"

def normalize_text(text: str) -> str:
    text = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]', ' ', str(text))
    text = re.sub(r'\s+', ' ', text)
    return text.strip()

# ===== HITUNG SKOR KELENGKAPAN =====
FIELD_WAJIB = [
    "sep",
    "laporan_operasi",
    "rujukan",
    "hasil_penunjang",
    "general_consent",
    "informed_consent"
]

FIELD_OPSIONAL = ["resep_obat", "hasil_lab", "radiologi", "consent", "resume"]

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
    "hasil_penunjang": ("🧪 Penunjang: Hasil Penunjang", "Gabungan hasil laboratorium, radiologi, dan pemeriksaan pendukung lainnya."),
    "resume": ("📝 Ringkasan: Resume Medis (Discharge Summary)", "Lembar ringkasan pelayanan dari awal masuk hingga pasien pulang.")
}

def hitung_skor_kelengkapan(ml_result: dict) -> int:
    # Ambil semua field dari FIELD_LABELS
    field_keys = [k for k in FIELD_LABELS.keys()]
    total_field = len(field_keys)
    if total_field == 0:
        return 0
    total_terpenuhi = sum(1 for k in field_keys if ml_result.get(k) is True)
    return int((total_terpenuhi / total_field) * 100)

# ===== PREDICT FIELDS =====
def predict_fields(text: str):
    if not text or not text.strip():
        return {f: False for f in FIELD_WAJIB + FIELD_OPSIONAL} | {"skor_kelengkapan": 0}

    txt_lower = text.lower()

    # Field wajib
    sep = bool(re.search(r'\b(sep|surat\s+eligibilitas\s+peserta|no\.?\s*sep|nomor\s*sep)\b', txt_lower))
    hasil_operasi = bool(re.search(r'\b(hasil\s*operasi|laporan\s*operasi|post\s*op|tindakan\s*operasi)\b', txt_lower))
    rujukan = bool(re.search(r'\bsurat\s+rujukan\b|\bno\.?\s*rujukan\b', txt_lower))
    hasil_lab = bool(re.search(r'hematologi|hemoglobin|leukosit|trombosit|sgot|sgpt|ureum|kreatinin', txt_lower))
    radiologi = bool(re.search(r'radiologi|rontgen|ct\s*scan|mri|usg|x-ray|imaging', txt_lower))
    hasil_penunjang = hasil_lab or radiologi
    general_consent = bool(re.search(r'resume\s+medis|discharge\s+summary|keluhan\s+utama', txt_lower))
    informed_consent = bool(re.search(r'consent|persetujuan|tanda\s+tangan\s+(pasien|wali)', txt_lower))

    # Field opsional tambahan
    resep_obat = bool(re.search(r'\d+\s*mg|resep|haloperidol|risperidone|olanzapin|clozapin|quetiapine', txt_lower))
    consent = bool(re.search(r'consent|persetujuan|tanda\s+tangan', txt_lower))
    resume = bool(re.search(r'resume\s+medis|discharge\s+summary', txt_lower))
    nama_pasien = bool(re.search(r'nama\s*pasien|identitas\s*pasien|pasien\s*:|tn\.|ny\.',txt_lower))
    no_rm = bool(re.search(r'no\.?\s*rm|nomor\s*rm|rekam\s*medis|norm',txt_lower))
    diagnosis = bool(re.search(r'diagnosis|diagnosa|dx|icd-10|icd',txt_lower))
    dokter = bool(re.search(r'dokter|dpjp|dr\.|sp\.kj|spkj',txt_lower))

    final_result = {
        "nama_pasien": nama_pasien,
        "no_rm": no_rm,
        "diagnosis": diagnosis,
        "dokter": dokter,
        "sep": sep,
        "laporan_operasi": hasil_operasi,
        "rujukan": rujukan,
        "hasil_penunjang": hasil_penunjang,
        "general_consent": general_consent,
        "informed_consent": informed_consent,
        "resep_obat": resep_obat,
        "hasil_lab": hasil_lab,
        "radiologi": radiologi,
        "consent": consent,
        "resume": resume
    }

    final_result["skor_kelengkapan"] = hitung_skor_kelengkapan(final_result)
    return final_result

# ===== RUN PIPELINE =====
def run_pipeline(batch_size=50):
    print("--- MEMULAI UPDATE DATA KE SUPABASE ---")
    sb = get_supabase()
    if sb is None:
        print("❌ Supabase URL/KEY belum dikonfigurasi.")
        return
        
    # 1. Ambil metadata file dari tabel 'berkas'
    try:
        response_berkas = sb.table("berkas").select("id, path_file, tipe_file").limit(batch_size).execute()
        records_berkas = response_berkas.data
    except Exception as e:
        print(f"❌ Gagal mengambil data dari tabel 'berkas': {e}")
        return
    
    if not records_berkas:
        print("ℹ️ Tidak ada data ditemukan di tabel berkas.")
        return

    for r in records_berkas:
        id_berkas = r.get("id")
        path_file = r.get("path_file")
        tipe_file = r.get("tipe_file")
        
        print(f"\n📄 Memproses Berkas ID {id_berkas}...")
        
        # Ekstrak & analisis teks berkas
        raw_text = extract_text(path_file, tipe_file)
        text_norm = normalize_text(raw_text)
        
        prediction = predict_fields(text_norm)
        skor = prediction.get("skor_kelengkapan", 0)
        
        # 2. LOGIKA STATUS SESUAI KODEMU 📋
        if skor >= 80:
            status_label = "Layak Klaim"
        elif skor >= 50:
            status_label = "Perlu Peninjauan"
        else:
            status_label = "Tidak Layak Klaim"
        
        # 3. Update hasil ke tabel 'hasil_validasi'
        try:
            sb.table("hasil_validasi").update({
                "hasil_json": prediction,          # Sesuai kolom hasil_json (jsonb)
                "skor_kelengkapan": skor,          # Sesuai kolom skor_kelengkapan (int4)
                "status_validasi": status_label    # Menggunakan status label baru milikmu!
            }).eq("id", id_berkas).execute()
            
            print(f"✅ Pasien ID {id_berkas} berhasil diproses: Skor {skor}% -> Status: {status_label}")
        except Exception as e:
            print(f"❌ Gagal update ID {id_berkas} ke hasil_validasi: {e}")

if __name__ == "__main__":
    print("--- MEMULAI UPDATE DATA KE SUPABASE ---")
    run_pipeline(batch_size=50)
    print("--- SELESAI UPDATE ---")