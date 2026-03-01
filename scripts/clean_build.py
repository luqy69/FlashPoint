import os
import shutil
import subprocess
import time

def clean_and_build():
    print("Starting Deep Clean...")
    
    # Files/Dirs to remove
    targets = [
        "build",
        "dist",
        "FlashPoint.spec",
        "FlashPoint_v2.spec",
        "FlashPoint_v3.spec",
        "package/FlashPoint.exe"
    ]
    
    for t in targets:
        if os.path.exists(t):
            try:
                if os.path.isdir(t):
                    shutil.rmtree(t)
                else:
                    os.remove(t)
                print(f"[REMOVED] {t}")
            except Exception as e:
                print(f"[ERROR] Could not remove {t}: {e}")
    
    print("Clean complete. Waiting 2 seconds...")
    time.sleep(2)
    
    print("Running build.py...")
    subprocess.check_call(["python", "build.py"])

if __name__ == "__main__":
    clean_and_build()
