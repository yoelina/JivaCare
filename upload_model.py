from huggingface_hub import HfApi, upload_folder

api = HfApi()
api.upload_folder(
    folder_path=r"C:/Users/HP/dashboardpjd/model_ml",  # folder yang berisi model
    path_in_repo=".",                                   # simpan di root repo Hugging Face
    repo_id="yoelina/model",                            # nama repo Hugging Face
    repo_type="model",
    token="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"                            # token Hugging Face yang valid
)
print("Upload selesai ✅")