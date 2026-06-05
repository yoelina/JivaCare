from supabase import create_client
import os

url = os.environ.get("SUPABASE_URL")
key = os.environ.get("SUPABASE_KEY")

supabase = create_client(url, key)
try:
    data = supabase.table("training_dataset").select("*").limit(1).execute()
    print("Koneksi berhasil, data contoh:", data.data)
except Exception as e:
    print("Error:", e)