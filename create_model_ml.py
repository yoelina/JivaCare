from transformers import AutoTokenizer, AutoModelForTokenClassification
import os

# Folder target
MODEL_DIR = "C:/Users/HP/dashboardpjd/model_ml"
os.makedirs(MODEL_DIR, exist_ok=True)

# Ambil model base IndoBERT
tokenizer = AutoTokenizer.from_pretrained("indobenchmark/indobert-base-p1")
tokenizer.save_pretrained(MODEL_DIR)

model = AutoModelForTokenClassification.from_pretrained(
    "indobenchmark/indobert-base-p1",
    num_labels=5  # nanti sesuaikan dengan label RM_50
)
model.save_pretrained(MODEL_DIR)

print("Folder model_ml sudah berisi config.json, pytorch_model.bin, dan tokenizer files.")