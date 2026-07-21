import subprocess
import sys
import os
import time
import configparser
import shutil
import threading

def stream_logs(prefix_folder, pipe):
    try:
        for line in iter(pipe.readline, ''):
            if line:
                print(f"[{prefix_folder}] {line.rstrip()}")
    except Exception:
        pass
    finally:
        pipe.close()

def main():
    config_file = 'depo.conf'
    
    if not os.path.exists(config_file):
        print(f"--> [ERROR] File {config_file} tidak ditemukan di folder ini!")
        return

    config = configparser.ConfigParser()
    config.read(config_file)

    if not config.has_section('DEPO'):
        print("--> [ERROR] Section [DEPO] tidak ditemukan di dalam depo.conf!")
        return

    proses_berjalan = []
    threads = []
    files_to_transfer = ['Piutang.xls', 'Giro.xls']

    print("--> " + "=" * 50)
    print("--> Menjalankan Sistem Automasi AR DEPO")
    print("--> " + "=" * 50)

    try:
        for option, path_raw in config.items('DEPO'):
            path = path_raw.strip()
            
            if path.startswith('#'):
                print(f"--> [LEWAT] Jalur dinonaktifkan via tanda komen: {path}")
                continue
            
            if not path:
                continue

            if os.path.exists(path):
                nama_folder = os.path.dirname(path)
                nama_file = os.path.basename(path)
                
                print(f"\n--> [PROSES] Mengamankan direktori: '{nama_folder}'")
                
                for f_name in files_to_transfer:
                    target_file_path = os.path.join(nama_folder, f_name)
                    source_file_path = f_name
                    
                    if os.path.exists(target_file_path):
                        try:
                            os.remove(target_file_path)
                            print(f"    [-] Berhasil menghapus file lama: {target_file_path}")
                        except Exception as e:
                            print(f"    [!] Gagal menghapus {target_file_path}: {e}")
                    
                    if os.path.exists(source_file_path):
                        try:
                            shutil.copy2(source_file_path, target_file_path)
                            print(f"    [+] Berhasil menyalin file master baru ke: {target_file_path}")
                        except Exception as e:
                            print(f"    [!] Gagal menyalin ke {target_file_path}: {e}")
                    else:
                        print(f"    [!] Peringatan: File sumber master {source_file_path} tidak ditemukan!")

                print(f"--> [INFO] Menjalankan '{nama_file}' di dalam lingkungan '{nama_folder}'...")
                
                proses = subprocess.Popen(
                    [sys.executable, "-u", nama_file],
                    cwd=nama_folder,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    text=True,
                    bufsize=1
                )
                proses_berjalan.append(proses)

                t = threading.Thread(
                    target=stream_logs, 
                    args=(nama_folder, proses.stdout), 
                    daemon=True
                )
                t.start()
                threads.append(t)
            else:
                print(f"--> [PERINGATAN] File tidak ditemukan pada lokasi target: {path}")

        print("\n--> " + "=" * 50)
        print("--> [SUKSES] Semua automasi aktif berhasil diperbarui dan kini berjalan.")
        print("--> -> Tekan CTRL + C pada terminal ini untuk menghentikan semuanya sekaligus.\n")

        while True:
            time.sleep(1)

    except KeyboardInterrupt:
        print("\n--> [MENDETEKSI CTRL+C] Menghentikan seluruh sub-sistem automasi, mohon tunggu...")
        for proses in proses_berjalan:
            proses.terminate()
        print("--> [SELESAI] Semua proses automasi aktif telah berhasil dimatikan secara bersih.")

if __name__ == "__main__":
    main()
