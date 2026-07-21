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
                print(f"[{prefix_folder}] {line.rstrip()}", flush=True)
    except Exception:
        pass
    finally:
        pipe.close()

def main():
    config_file = 'depo.conf'
    
    if not os.path.exists(config_file):
        print(f"--> [ERROR] File {config_file} tidak ditemukan di folder ini!", flush=True)
        return

    config = configparser.ConfigParser()
    config.read(config_file)

    if not config.has_section('DEPO'):
        print("--> [ERROR] Section [DEPO] tidak ditemukan di dalam depo.conf!", flush=True)
        return

    proses_berjalan = []
    threads = []
    files_to_transfer = ['Piutang.xls', 'Giro.xls']

    print("--> " + "=" * 50, flush=True)
    print("--> Menjalankan Sistem Automasi AR DEPO", flush=True)
    print("--> " + "=" * 50, flush=True)

    env_unbuffered = os.environ.copy()
    env_unbuffered["PYTHONUNBUFFERED"] = "1"
    
    try:
        for option, path_raw in config.items('DEPO'):
            path = path_raw.strip()
            
            if path.startswith('#'):
                print(f"--> [LEWAT] Jalur dinonaktifkan via tanda komen: {path}", flush=True)
                continue
            
            if not path:
                continue

            if os.path.exists(path):
                nama_folder = os.path.dirname(path)
                nama_file = os.path.basename(path)
                
                print(f"\n--> [PROSES] Mengamankan direktori: '{nama_folder}'", flush=True)
                
                for f_name in files_to_transfer:
                    target_file_path = os.path.join(nama_folder, f_name)
                    source_file_path = f_name
                    
                    if os.path.exists(target_file_path):
                        try:
                            os.remove(target_file_path)
                            print(f"    [-] Berhasil menghapus file lama: {target_file_path}", flush=True)
                        except Exception as e:
                            print(f"    [!] Gagal menghapus {target_file_path}: {e}", flush=True)
                    
                    if os.path.exists(source_file_path):
                        try:
                            shutil.copy2(source_file_path, target_file_path)
                            print(f"    [+] Berhasil menyalin file master baru ke: {target_file_path}", flush=True)
                        except Exception as e:
                            print(f"    [!] Gagal menyalin ke {target_file_path}: {e}", flush=True)
                    else:
                        print(f"    [!] Peringatan: File sumber master {source_file_path} tidak ditemukan!", flush=True)

                print(f"--> [INFO] Menjalankan '{nama_file}' di dalam lingkungan '{nama_folder}'...", flush=True)
                
                proses = subprocess.Popen(
                    [sys.executable, "-u", nama_file],
                    cwd=nama_folder,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    text=True,
                    bufsize=1,
                    env=env_unbuffered
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
                print(f"--> [PERINGATAN] File tidak ditemukan pada lokasi target: {path}", flush=True)

        print("\n--> " + "=" * 50, flush=True)
        print("--> [SUKSES] Semua automasi aktif berhasil diperbarui dan kini berjalan.", flush=True)
        print("--> -> Tekan CTRL + C pada terminal ini untuk menghentikan semuanya sekaligus.\n", flush=True)

        while True:
            time.sleep(1)

    except KeyboardInterrupt:
        print("\n--> [MENDETEKSI CTRL+C] Menghentikan seluruh sub-sistem automasi, mohon tunggu...", flush=True)
        for proses in proses_berjalan:
            proses.terminate()
        print("--> [SELESAI] Semua proses automasi aktif telah berhasil dimatikan secara bersih.", flush=True)

if __name__ == "__main__":
    main()