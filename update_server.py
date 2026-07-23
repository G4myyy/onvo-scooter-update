"""
ONVO Scooter - Google Drive Guncelleme Script'i
================================================
Kullanim:
  1. OnvoScooter.exe dosyasini Google Drive'a yukle
  2. Herkese acik paylasim linki al
  3. Bu scripti calistir: python update_server.py
  4. Linki yapistir, ENTER'a bas
"""

import os
import re
import json
import requests

FIREBASE_URL = "https://kaen-onvo-scooter-default-rtdb.firebaseio.com"

def extract_drive_id(url):
    patterns = [
        r'/file/d/([a-zA-Z0-9_-]+)',
        r'id=([a-zA-Z0-9_-]+)',
        r'/open\?id=([a-zA-Z0-9_-]+)',
    ]
    for p in patterns:
        m = re.search(p, url)
        if m:
            return m.group(1)
    return None

def main():
    ver_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "main.py")
    app_version = ""
    with open(ver_file, "r", encoding="utf-8") as f:
        for line in f:
            if "APP_VERSION" in line and "=" in line:
                app_version = line.split("=")[1].strip().strip('"').strip("'")
                break

    if not app_version:
        print("HATA: APP_VERSION bulunamadi!")
        return

    print(f"Guncel versiyon: v{app_version}")
    print()
    print("Google Drive linkini yapistir:")
    print("(Ornek: https://drive.google.com/file/d/XXXXX/view?usp=sharing)")
    print()

    link = input("Link: ").strip()
    if not link:
        print("Link bos!")
        return

    file_id = extract_drive_id(link)
    if not file_id:
        print("HATA: Google Drive linki taninamadi!")
        return

    download_url = f"https://drive.google.com/uc?id={file_id}&export=download"
    print(f"\nIndirme linki: {download_url}")

    print("\nFirebase guncelleniyor...")
    r = requests.put(
        f"{FIREBASE_URL}/app_version.json",
        json={"version": app_version, "url": download_url},
        timeout=10,
    )
    if r.status_code == 200:
        print(f"Basarili! v{app_version} Firebase'e kaydedildi.")
        print("Diger cihazlar uygulamayi actiginda guncellemeyi alacak.")
    else:
        print(f"HATA: Firebase guncellenemedi (HTTP {r.status_code})")

if __name__ == "__main__":
    main()
