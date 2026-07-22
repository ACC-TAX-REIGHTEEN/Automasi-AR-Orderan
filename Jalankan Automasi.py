import os
import glob
import shutil
import subprocess
import sys

def jalankan_otomatisasi():
    if not os.path.isfile("Piutang.xls"):
        print("--> File Piutang.xls tidak ditemukan. Proses digagalkan.")
        input("--> Tekan enter untuk keluar.")
        return

    folder_dapur = "Dapur"
    file_syarat = [
        "0_DownloaderData.py",
        "0_HDownloaderData.py",
        "1_CleanerAccAR.py",
        "1_CleanerAccARD.py",
        "2_CleanerAccGiro.py",
        "2_HCleanerAccGiroDue.py",
        "3_AddGiroToSheet.py",
        "4_PatchFallbackCash.py",
        "5_AdjDateFormat.py",
        "6_InjectDataToSS.py",
        "config.conf",
        "credentials.json",
        "__init__.py"
    ]

    if not os.path.exists(folder_dapur) or not os.path.isdir(folder_dapur):
        print("--> Folder Dapur tidak ditemukan.")
        input("--> Tekan enter untuk keluar.")
        return

    for file in file_syarat:
        jalur_file = os.path.join(folder_dapur, file)
        if not os.path.isfile(jalur_file):
            print(f"--> File {file} tidak ditemukan di dalam folder Dapur.")
            input("--> Tekan enter untuk keluar.")
            return

    file_xls = glob.glob(os.path.join(folder_dapur, "*.xls"))
    file_xlsx = glob.glob(os.path.join(folder_dapur, "*.xlsx"))
    semua_file_lama = file_xls + file_xlsx

    for file in semua_file_lama:
        try:
            os.remove(file)
        except Exception:
            pass

    try:
        shutil.copy("Piutang.xls", os.path.join(folder_dapur, "Piutang.xls"))
        if os.path.isfile("Giro.xls"):
            shutil.copy("Giro.xls", os.path.join(folder_dapur, "Giro.xls"))
    except Exception:
        print("--> Gagal memindahkan file ke folder Dapur.")
        input("--> Tekan enter untuk keluar.")
        return

    scripts = [
        "0_DownloaderData.py",
        "0_HDownloaderData.py",
        "1_CleanerAccAR.py",
        "1_CleanerAccARD.py",
        "2_CleanerAccGiro.py",
        "2_HCleanerAccGiroDue.py",
        "3_AddGiroToSheet.py",
        "4_PatchFallbackCash.py",
        "5_AdjDateFormat.py"
    ]

    try:
        for script in scripts:
            print(f"--> Memulai eksekusi {script}")
            subprocess.run([sys.executable, script], cwd=folder_dapur)

        print("--> Memulai eksekusi 6_InjectDataToSS.py")
        print("--> Tekan Ctrl+C untuk menghentikan loop sinkronisasi.")
        subprocess.run([sys.executable, "6_InjectDataToSS.py"], cwd=folder_dapur)
    except KeyboardInterrupt:
        print("\n--> Looping dihentikan oleh pengguna.")

    print("--> Semua proses telah selesai dijalankan.")
    input("--> Tekan enter untuk keluar.")

if __name__ == "__main__":
    jalankan_otomatisasi()
