import os
import time
import shutil

UPLOAD_DIR = "uploads"
OUTPUT_DIR = "outputs"

MAX_AGE_SECONDS = 30 * 60  

def cleanup_old_files():
    """Hapus file/folder yang tersisa di uploads/outputs lebih dari MAX_AGE_SECONDS."""
    now = time.time()

    for base_dir in (UPLOAD_DIR, OUTPUT_DIR):
        if not os.path.isdir(base_dir):
            continue

        for entry in os.listdir(base_dir):
            path = os.path.join(base_dir, entry)
            try:
                age = now - os.path.getmtime(path)
                if age <= MAX_AGE_SECONDS:
                    continue

                if os.path.isdir(path):
                    shutil.rmtree(path, ignore_errors=True)
                else:
                    os.remove(path)

                print(f"[safety-net] Membersihkan file sisa: {path} (umur {int(age)} detik)")

            except FileNotFoundError:
                continue
            except Exception as e:
                print(f"[safety-net] Gagal membersihkan {path}: {e}")