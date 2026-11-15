import os
import time
import shutil

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(BASE_DIR)

# Folder uploads & outputs berada di dalam folder backend atau project root
# Jika ada di backend/, prioritaskan backend/; jika di root, gunakan root.
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
OUTPUT_DIR = os.path.join(BASE_DIR, "outputs")

MAX_AGE_SECONDS = 30 * 60  

def cleanup_old_files():
    """Hapus file/folder yang tersisa di uploads/outputs lebih dari MAX_AGE_SECONDS."""
    now = time.time()

    # Periksa baik di backend/ maupun di root (jika pernah terbuat di CWD berbeda)
    target_dirs = {
        UPLOAD_DIR,
        OUTPUT_DIR,
        os.path.join(PROJECT_ROOT, "uploads"),
        os.path.join(PROJECT_ROOT, "outputs")
    }

    for base_dir in target_dirs:
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
