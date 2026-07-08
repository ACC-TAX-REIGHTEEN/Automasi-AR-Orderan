import os
import glob
import shutil
import subprocess
import sys

def jalankan_otomatisasi():
    folder_dapur = "Dapur"
    file_syarat = ["__init__.py", "1_CleanerAcc.py", "2_FilterAR.py", "3_CalculateAR.py", "4_HelperCleaningData.py", "5_InjectDataToSS.py", "credentials.json", "piutang.conf"]
    
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

    file_temp = glob.glob(os.path.join(folder_dapur, "*temp.xlsx"))
    file_export = glob.glob(os.path.join(folder_dapur, "ExportFile.xls"))

    semua_file_lama = file_temp + file_export

    for file in semua_file_lama:
        try:
            os.remove(file)
        except Exception:
            pass

    file_sumber = ["ExportFile.xls"]
    ada_file_dipindah = False
    for file in file_sumber:
        if os.path.isfile(file):
            shutil.copy2(file, os.path.join(folder_dapur, file))
            ada_file_dipindah = True

    if not ada_file_dipindah:
        print("--> File ExportFile.xls tidak ditemukan untuk diproses.")
        input("--> Tekan enter untuk keluar.")
        return

    print("--> Memulai eksekusi pembersihan data")
    subprocess.run([sys.executable, "1_CleanerAcc.py"], cwd=folder_dapur)

    print("--> Memulai eksekusi filter data sementara")
    subprocess.run([sys.executable, "2_FilterAR.py"], cwd=folder_dapur)
    
    print("--> Memulai eksekusi menyalin dan menyusun data pada template")
    subprocess.run([sys.executable, "3_CalculateAR.py"], cwd=folder_dapur)

    file_laporan = glob.glob(os.path.join(folder_dapur, "*AR.xlsx"))
    for laporan in file_laporan:
        nama_file = os.path.basename(laporan)
        shutil.copy2(laporan, nama_file)
        
    print("--> Memulai persiapan data untuk di susun ke Spreadsheets")
    subprocess.run([sys.executable, "4_HelperCleaningData.py"], cwd=folder_dapur)
    
    print("--> Memulai unggah data ke Spreadsheets")
    subprocess.run([sys.executable, "5_InjectDataToSS.py"], cwd=folder_dapur)

    file_temp = glob.glob(os.path.join(folder_dapur, "*temp.xlsx"))
    file_export = glob.glob(os.path.join(folder_dapur, "ExportFile.xls"))
    file_print = glob.glob(os.path.join(folder_dapur, "Print_AR.xlsx"))

    semua_file_dihapus = file_temp + file_export + file_print

    for file in semua_file_dihapus:
        try:
            os.remove(file)
        except Exception:
            pass
            
    print("--> Semua proses telah selesai dijalankan.")
    input("--> Tekan enter untuk keluar.")

if __name__ == "__main__":
    jalankan_otomatisasi()
