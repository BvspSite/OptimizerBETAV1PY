import os
import psutil
import platform
import subprocess
import shutil
import sys
import ctypes
import winreg
from typing import List, Optional
from tkinter import Tk, filedialog, messagebox, Menu, Label, Frame, Button, Text, Scrollbar, END
from comtypes import client
from pdf2docx import Converter

class XboyOptimizer:
    def __init__(self, root):
        self.root = root
        self.os_type = platform.system()
        self.is_admin = self._check_admin()
        self.setup_ui()
        
        # Untuk Windows, coba jalankan sebagai admin
        if self.os_type == "Windows" and not self.is_admin:
            if messagebox.askyesno("Admin Required", "Beberapa fitur membutuhkan hak admin. Jalankan sebagai admin?"):
                self._run_as_admin()

    def _check_admin(self) -> bool:
        """Check if running as admin"""
        try:
            if self.os_type == "Windows":
                return ctypes.windll.shell32.IsUserAnAdmin() != 0
            return os.getuid() == 0
        except:
            return False

    def _run_as_admin(self):
        """Run as admin (Windows only)"""
        if self.os_type == "Windows":
            ctypes.windll.shell32.ShellExecuteW(
                None, "runas", sys.executable, " ".join(sys.argv), None, 1)
            sys.exit()

    def setup_ui(self):
        """Setup the main application UI"""
        self.root.title("Xboy Optimizer & Converter")
        self.root.geometry("900x700")
        
        # Main menu
        menubar = Menu(self.root)
        
        # Optimizer menu
        optimizer_menu = Menu(menubar, tearoff=0)
        optimizer_menu.add_command(label="Bersihkan File Temporary", command=self.clear_temp_files)
        optimizer_menu.add_command(label="Bersihkan Cache Browser", command=self.clean_browser_cache)
        optimizer_menu.add_command(label="Optimasi Startup", command=self.optimize_startup)
        optimizer_menu.add_command(label="Cek Penggunaan Disk", command=self.check_disk_usage)
        optimizer_menu.add_command(label="Cek Penggunaan Resource", command=self.check_resource_usage)
        optimizer_menu.add_separator()
        optimizer_menu.add_command(label="Defragmentasi Disk", command=self.defragment_disk)
        optimizer_menu.add_command(label="Cek Update Sistem", command=self.update_system)
        optimizer_menu.add_command(label="Boost FPS", command=self.boost_fps)
        menubar.add_cascade(label="Optimizer", menu=optimizer_menu)
        
        # Converter menu
        converter_menu = Menu(menubar, tearoff=0)
        converter_menu.add_command(label="Word ke PDF", command=self.word_to_pdf)
        converter_menu.add_command(label="PDF ke Word", command=self.pdf_to_word)
        menubar.add_cascade(label="Konversi Dokumen", menu=converter_menu)
        
        # Help menu
        help_menu = Menu(menubar, tearoff=0)
        help_menu.add_command(label="Tentang", command=self.show_about)
        menubar.add_cascade(label="Bantuan", menu=help_menu)
        
        self.root.config(menu=menubar)

        # Main content
        Label(self.root, text="Xboy Optimizer & Converter", 
              font=("Arial", 18, "bold")).pack(pady=10)
        
        Label(self.root, text="Alat Optimasi Sistem dan Konversi Dokumen", 
              font=("Arial", 12)).pack(pady=5)
        
        # Output console
        self.output_frame = Frame(self.root)
        self.output_frame.pack(pady=10, padx=10, fill="both", expand=True)
        
        self.output_text = Text(self.output_frame, wrap="word", state="disabled")
        scrollbar = Scrollbar(self.output_frame, command=self.output_text.yview)
        self.output_text.configure(yscrollcommand=scrollbar.set)
        
        scrollbar.pack(side="right", fill="y")
        self.output_text.pack(side="left", fill="both", expand=True)
        
        # Quick action buttons
        button_frame = Frame(self.root)
        button_frame.pack(pady=10)
        
        Button(button_frame, text="Bersihkan Cepat", command=self.quick_clean, 
              width=20, height=2).grid(row=0, column=0, padx=5, pady=5)
        Button(button_frame, text="Boost Performa", command=self.boost_fps, 
              width=20, height=2).grid(row=0, column=1, padx=5, pady=5)
        Button(button_frame, text="Word ke PDF", command=self.word_to_pdf, 
              width=20, height=2).grid(row=1, column=0, padx=5, pady=5)
        Button(button_frame, text="PDF ke Word", command=self.pdf_to_word, 
              width=20, height=2).grid(row=1, column=1, padx=5, pady=5)

    def log_message(self, message):
        """Menampilkan pesan di output console"""
        self.output_text.config(state="normal")
        self.output_text.insert(END, message + "\n")
        self.output_text.see(END)
        self.output_text.config(state="disabled")
        self.root.update()

    def clear_temp_files(self):
        """Membersihkan file temporary"""
        self.log_message("\n[1] Membersihkan file temporary...")
        try:
            if self.os_type == "Windows":
                temp_paths = [
                    os.path.join(os.environ['TEMP']),
                    os.path.join(os.environ['WINDIR'], 'Temp'),
                    os.path.expanduser('~\\AppData\\Local\\Temp')
                ]
                
                for temp_path in temp_paths:
                    if os.path.exists(temp_path):
                        for root, dirs, files in os.walk(temp_path):
                            for file in files:
                                try:
                                    file_path = os.path.join(root, file)
                                    os.unlink(file_path)
                                    self.log_message(f"Menghapus: {file_path}")
                                except Exception as e:
                                    continue
                
                # Bersihkan Prefetch
                prefetch_path = r'C:\Windows\Prefetch'
                if os.path.exists(prefetch_path):
                    for filename in os.listdir(prefetch_path):
                        try:
                            file_path = os.path.join(prefetch_path, filename)
                            if os.path.isfile(file_path):
                                os.unlink(file_path)
                                self.log_message(f"Menghapus prefetch: {filename}")
                        except:
                            continue
                
                self.log_message("✅ Berhasil membersihkan file temporary di Windows")
            
            elif self.os_type == "Linux":
                tmp_paths = ['/tmp', '/var/tmp', os.path.expanduser('~/.cache')]
                for tmp_path in tmp_paths:
                    if os.path.exists(tmp_path):
                        for filename in os.listdir(tmp_path):
                            try:
                                path = os.path.join(tmp_path, filename)
                                if os.path.isfile(path) or os.path.islink(path):
                                    os.unlink(path)
                                    self.log_message(f"Menghapus: {path}")
                                elif os.path.isdir(path):
                                    shutil.rmtree(path)
                                    self.log_message(f"Menghapus folder: {path}")
                            except:
                                continue
                
                self.log_message("✅ Berhasil membersihkan file temporary di Linux")
            
            elif self.os_type == "Darwin":
                cache_paths = [
                    os.path.expanduser('~/Library/Caches'),
                    os.path.expanduser('~/Library/Logs'),
                    '/Library/Caches',
                    '/Library/Logs'
                ]
                for cache_path in cache_paths:
                    if os.path.exists(cache_path):
                        for root, dirs, files in os.walk(cache_path):
                            for file in files:
                                try:
                                    os.unlink(os.path.join(root, file))
                                    self.log_message(f"Menghapus: {file}")
                                except:
                                    continue
                
                self.log_message("✅ Berhasil membersihkan file temporary di macOS")
            
        except Exception as e:
            self.log_message(f"❌ Gagal membersihkan file temporary: {e}")

    def clean_browser_cache(self):
        """Membersihkan cache browser populer"""
        self.log_message("\n[2] Membersihkan cache browser...")
        browsers = {
            "Chrome": {
                "Windows": [
                    os.path.expanduser('~\\AppData\\Local\\Google\\Chrome\\User Data\\Default\\Cache'),
                    os.path.expanduser('~\\AppData\\Local\\Google\\Chrome\\User Data\\Default\\Media Cache')
                ],
                "Linux": [
                    os.path.expanduser('~/.cache/google-chrome'),
                    os.path.expanduser('~/.config/google-chrome/Default/Cache')
                ],
                "Darwin": [
                    os.path.expanduser('~/Library/Caches/Google/Chrome'),
                    os.path.expanduser('~/Library/Application Support/Google/Chrome/Default/Cache')
                ]
            },
            "Firefox": {
                "Windows": [
                    os.path.expanduser('~\\AppData\\Local\\Mozilla\\Firefox\\Profiles')
                ],
                "Linux": [
                    os.path.expanduser('~/.mozilla/firefox'),
                    os.path.expanduser('~/.cache/mozilla/firefox')
                ],
                "Darwin": [
                    os.path.expanduser('~/Library/Caches/Firefox'),
                    os.path.expanduser('~/Library/Application Support/Firefox/Profiles')
                ]
            },
            "Edge": {
                "Windows": [
                    os.path.expanduser('~\\AppData\\Local\\Microsoft\\Edge\\User Data\\Default\\Cache'),
                    os.path.expanduser('~\\AppData\\Local\\Microsoft\\Edge\\User Data\\Default\\Media Cache')
                ],
                "Linux": [
                    os.path.expanduser('~/.cache/microsoft-edge'),
                    os.path.expanduser('~/.config/microsoft-edge/Default/Cache')
                ],
                "Darwin": [
                    os.path.expanduser('~/Library/Caches/Microsoft Edge'),
                    os.path.expanduser('~/Library/Application Support/Microsoft Edge/Default/Cache')
                ]
            }
        }
        
        cleaned = 0
        for browser, paths in browsers.items():
            path_list = paths.get(self.os_type, [])
            for path in path_list:
                if os.path.exists(path):
                    try:
                        if browser == "Firefox" and self.os_type == "Windows":
                            for profile in os.listdir(path):
                                if profile.endswith('.default-release'):
                                    cache_dirs = ['cache2', 'thumbnails', 'startupCache']
                                    for cache_dir in cache_dirs:
                                        cache_path = os.path.join(path, profile, cache_dir)
                                        if os.path.exists(cache_path):
                                            shutil.rmtree(cache_path)
                                            cleaned += 1
                                            self.log_message(f"Membersihkan cache {browser} di {cache_path}")
                        else:
                            shutil.rmtree(path)
                            os.makedirs(path, exist_ok=True)
                            cleaned += 1
                            self.log_message(f"Membersihkan cache {browser} di {path}")
                    except Exception as e:
                        self.log_message(f"❌ Gagal membersihkan cache {browser}: {e}")
        
        if cleaned > 0:
            self.log_message(f"✅ Berhasil membersihkan cache {cleaned} browser")
        else:
            self.log_message("⚠ Tidak ditemukan cache browser yang bisa dibersihkan")

    def optimize_startup(self):
        """Menonaktifkan program startup yang tidak perlu"""
        self.log_message("\n[3] Mengoptimalkan program startup...")
        try:
            if self.os_type == "Windows":
                self.log_message("\n📂 Program startup untuk user saat ini:")
                startup_path = os.path.join(os.environ['APPDATA'], 'Microsoft', 'Windows', 'Start Menu', 'Programs', 'Startup')
                for item in os.listdir(startup_path):
                    self.log_message(f"- {item}")
                
                self.log_message("\n🔧 Registry startup locations:")
                reg_locations = [
                    (winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run"),
                    (winreg.HKEY_LOCAL_MACHINE, r"Software\Microsoft\Windows\CurrentVersion\Run")
                ]
                
                for hive, subkey in reg_locations:
                    try:
                        with winreg.OpenKey(hive, subkey) as key:
                            i = 0
                            while True:
                                try:
                                    name, value, _ = winreg.EnumValue(key, i)
                                    self.log_message(f"- {name}: {value}")
                                    i += 1
                                except OSError:
                                    break
                    except WindowsError:
                        continue
                
                self.log_message("\nℹ Untuk menonaktifkan program startup:")
                self.log_message("1. Buka Task Manager (Ctrl+Shift+Esc)")
                self.log_message("2. Pergi ke tab 'Startup'")
                self.log_message("3. Klik kanan program yang ingin dinonaktifkan")
                self.log_message("4. Pilih 'Disable'")
            
            elif self.os_type == "Linux":
                if os.path.exists('/etc/systemd/system'):
                    self.log_message("\n🛠 Systemd services yang aktif:")
                    result = subprocess.run(['systemctl', 'list-unit-files', '--type=service', '--state=enabled'], 
                                          capture_output=True, text=True)
                    self.log_message(result.stdout)
                
                self.log_message("\n⏰ Cron jobs untuk user saat ini:")
                result = subprocess.run(['crontab', '-l'], capture_output=True, text=True)
                self.log_message(result.stdout if result.stdout else "Tidak ada cron job")
                
                self.log_message("\nℹ Untuk menonaktifkan service:")
                self.log_message("sudo systemctl disable nama-service.service")
                self.log_message("\nℹ Untuk menghapus cron job:")
                self.log_message("crontab -e")
            
            elif self.os_type == "Darwin":
                result = subprocess.run(['osascript', '-e', 'tell application "System Events" to get the name of every login item'], 
                                      capture_output=True, text=True)
                self.log_message("\n🍏 Login items untuk user saat ini:")
                self.log_message(result.stdout if result.stdout else "Tidak ada login item")
                
                self.log_message("\nℹ Untuk menonaktifkan login items:")
                self.log_message("1. Buka System Preferences > Users & Groups")
                self.log_message("2. Pilih tab Login Items")
                self.log_message("3. Pilih item dan klik tombol '-'")
            
            self.log_message("\n✅ Rekomendasi startup program telah ditampilkan")
        
        except Exception as e:
            self.log_message(f"❌ Gagal memeriksa program startup: {e}")

    def check_disk_usage(self):
        """Memeriksa dan membersihkan space disk"""
        self.log_message("\n[4] Memeriksa penggunaan disk...")
        try:
            partitions = psutil.disk_partitions()
            for part in partitions:
                try:
                    usage = psutil.disk_usage(part.mountpoint)
                    self.log_message(f"\n📊 Partisi {part.mountpoint}:")
                    self.log_message(f"💽 Total: {self._bytes_to_gb(usage.total):.2f} GB")
                    self.log_message(f"📌 Digunakan: {self._bytes_to_gb(usage.used):.2f} GB ({usage.percent}%)")
                    self.log_message(f"🆓 Tersedia: {self._bytes_to_gb(usage.free):.2f} GB")
                    
                    if usage.free / usage.total < 0.2:
                        self.log_message("⚠ Peringatan: Ruang disk hampir penuh!")
                        self.log_message("💡 Rekomendasi:")
                        self.log_message("- Hapus file besar yang tidak perlu")
                        self.log_message("- Uninstall program yang tidak digunakan")
                        self.log_message("- Kosongkan recycle bin/trash")
                except Exception as e:
                    self.log_message(f"❌ Gagal memeriksa partisi {part.mountpoint}: {e}")
            
            # Disk cleanup untuk Windows
            if self.os_type == "Windows":
                self.log_message("\n🚀 Menjalankan Disk Cleanup...")
                try:
                    subprocess.run(['cleanmgr', '/sagerun:1'], shell=True, check=False)
                    self.log_message("✅ Disk Cleanup telah dijalankan")
                except Exception as e:
                    self.log_message(f"❌ Gagal menjalankan Disk Cleanup: {e}")
        
        except Exception as e:
            self.log_message(f"❌ Gagal memeriksa penggunaan disk: {e}")

    def check_resource_usage(self):
        """Memeriksa penggunaan CPU, RAM, dan disk"""
        self.log_message("\n[5] Memeriksa penggunaan resource...")
        try:
            # CPU usage
            cpu_percent = psutil.cpu_percent(interval=1)
            cpu_count = psutil.cpu_count(logical=False)
            cpu_threads = psutil.cpu_count(logical=True)
            self.log_message(f"\n🖥️ Penggunaan CPU: {cpu_percent}%")
            self.log_message(f"🔢 Core fisik: {cpu_count}, Thread: {cpu_threads}")
            
            # RAM usage
            mem = psutil.virtual_memory()
            self.log_message(f"\n🧠 Penggunaan RAM: {mem.percent}%")
            self.log_message(f"📌 Digunakan: {self._bytes_to_gb(mem.used):.2f} GB")
            self.log_message(f"🆓 Tersedia: {self._bytes_to_gb(mem.available):.2f} GB / Total: {self._bytes_to_gb(mem.total):.2f} GB")
            
            # Swap usage (jika ada)
            swap = psutil.swap_memory()
            if swap.total > 0:
                self.log_message(f"\n🔀 Penggunaan Swap: {swap.percent}%")
                self.log_message(f"📌 Digunakan: {self._bytes_to_gb(swap.used):.2f} GB / Total: {self._bytes_to_gb(swap.total):.2f} GB")
            
            # Disk activity
            disk = psutil.disk_io_counters()
            self.log_message(f"\n💾 Aktivitas Disk:")
            self.log_message(f"📥 Baca: {self._bytes_to_mb(disk.read_bytes):.2f} MB")
            self.log_message(f"📤 Tulis: {self._bytes_to_mb(disk.write_bytes):.2f} MB")
            
            # Proses yang memakan banyak resource
            self.log_message("\n🔝 5 Proses dengan penggunaan CPU tertinggi:")
            for proc in sorted(psutil.process_iter(['name', 'cpu_percent']), 
                             key=lambda p: p.info['cpu_percent'], reverse=True)[:5]:
                self.log_message(f"- {proc.info['name']}: {proc.info['cpu_percent']}%")
            
            self.log_message("\n🔝 5 Proses dengan penggunaan RAM tertinggi:")
            for proc in sorted(psutil.process_iter(['name', 'memory_info']), 
                             key=lambda p: p.info['memory_info'].rss, reverse=True)[:5]:
                self.log_message(f"- {proc.info['name']}: {self._bytes_to_mb(proc.info['memory_info'].rss):.2f} MB")
            
            # Rekomendasi
            self.log_message("\n💡 Rekomendasi:")
            if mem.percent > 80:
                self.log_message("- Tutup aplikasi yang tidak diperlukan untuk menghemat RAM")
                if self.os_type == "Windows":
                    self.log_message("- Nonaktifkan efek visual: System Properties > Performance Settings > Adjust for best performance")
            if cpu_percent > 80:
                self.log_message("- Kurangi beban CPU dengan menutup aplikasi berat")
                self.log_message("- Cek proses yang menggunakan CPU tinggi di Task Manager")
            if swap.percent > 50 and swap.total > 0:
                self.log_message("- Penggunaan swap tinggi, pertimbangkan untuk menambah RAM")
        
        except Exception as e:
            self.log_message(f"❌ Gagal memeriksa penggunaan resource: {e}")

    def defragment_disk(self):
        """Menjalankan defragmentasi disk (hanya Windows)"""
        if self.os_type != "Windows":
            self.log_message("ℹ Defragmentasi hanya tersedia di Windows")
            return
            
        self.log_message("\n[6] Memeriksa kebutuhan defragmentasi...")
        try:
            result = subprocess.run(['defrag', 'C:', '/a', '/v'], 
                                  capture_output=True, text=True, shell=True)
            self.log_message(result.stdout)
            
            if "You don't need to defragment this volume" not in result.stdout:
                self.log_message("\n🚀 Menjalankan defragmentasi...")
                subprocess.run(['defrag', 'C:', '/u', '/v'], shell=True)
                self.log_message("✅ Defragmentasi selesai dijalankan")
            else:
                self.log_message("✅ Disk tidak perlu defragmentasi")
                
        except Exception as e:
            self.log_message(f"❌ Gagal menjalankan defragmentasi: {e}")

    def update_system(self):
        """Memeriksa update sistem"""
        self.log_message("\n[7] Memeriksa update sistem...")
        try:
            if self.os_type == "Windows":
                self.log_message("\n🚀 Memeriksa Windows Update...")
                subprocess.run(['wuauclt', '/detectnow'], shell=True)
                self.log_message("\n💡 Buka Settings > Update & Security untuk instal update")
            
            elif self.os_type == "Linux":
                self.log_message("\n🚀 Memeriksa update paket...")
                if os.path.exists('/etc/apt/sources.list'):
                    result = subprocess.run(['sudo', 'apt', 'update'], capture_output=True, text=True)
                    self.log_message(result.stdout)
                    self.log_message("\nDaftar update yang tersedia:")
                    result = subprocess.run(['sudo', 'apt', 'list', '--upgradable'], capture_output=True, text=True)
                    self.log_message(result.stdout)
                elif os.path.exists('/etc/yum.conf'):
                    result = subprocess.run(['sudo', 'yum', 'check-update'], capture_output=True, text=True)
                    self.log_message(result.stdout)
                elif os.path.exists('/etc/pacman.conf'):
                    result = subprocess.run(['sudo', 'pacman', '-Syuw'], capture_output=True, text=True)
                    self.log_message(result.stdout)
            
            elif self.os_type == "Darwin":
                self.log_message("\n🚀 Memeriksa update macOS...")
                result = subprocess.run(['softwareupdate', '-l'], capture_output=True, text=True)
                self.log_message(result.stdout)
            
            self.log_message("\n✅ Update sistem telah diperiksa")
        
        except Exception as e:
            self.log_message(f"❌ Gagal memeriksa update sistem: {e}")

    def boost_fps(self):
        """Meningkatkan performa gaming/FPS"""
        self.log_message("\n[8] Meningkatkan performa gaming/FPS...")
        try:
            if self.os_type == "Windows":
                self.log_message("\n🚀 Mengoptimalkan sistem untuk gaming...")
                
                # 1. Set priority untuk proses game
                game_processes = ['csgo.exe', 'dota2.exe', 'valorant.exe', 'fortnite.exe', 
                                'overwatch.exe', 'leagueoflegends.exe', 'game.exe']
                
                self.log_message("\n🔍 Mencari proses game yang sedang berjalan...")
                found = False
                for proc in psutil.process_iter(['name', 'pid']):
                    if proc.info['name'].lower() in [p.lower() for p in game_processes]:
                        try:
                            p = psutil.Process(proc.info['pid'])
                            p.nice(psutil.HIGH_PRIORITY_CLASS)
                            self.log_message(f"✅ Menaikkan prioritas proses {proc.info['name']} (PID: {proc.info['pid']})")
                            found = True
                        except Exception as e:
                            self.log_message(f"❌ Gagal mengubah prioritas {proc.info['name']}: {e}")
                
                if not found:
                    self.log_message("⚠ Tidak ditemukan proses game yang sedang berjalan")
                    self.log_message("💡 Jalankan game terlebih dahulu, lalu coba lagi")
                
                # 2. Nonaktifkan Game Bar dan DVR
                self.log_message("\n🛑 Menonaktifkan Game Bar dan DVR...")
                try:
                    subprocess.run(['reg', 'add', 'HKCU\\System\\GameConfigStore', 
                                  '/v', 'GameDVR_Enabled', '/t', 'REG_DWORD', '/d', '0', '/f'], 
                                 shell=True)
                    subprocess.run(['reg', 'add', 'HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\GameDVR', 
                                  '/v', 'AllowGameDVR', '/t', 'REG_DWORD', '/d', '0', '/f'], 
                                 shell=True)
                    self.log_message("✅ Game Bar dan DVR telah dinonaktifkan")
                except Exception as e:
                    self.log_message(f"❌ Gagal menonaktifkan Game Bar/DVR: {e}")
                
                # 3. Optimasi power plan
                self.log_message("\n⚡ Mengatur power plan ke High Performance...")
                try:
                    subprocess.run(['powercfg', '/setactive', '8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c'], 
                                 shell=True)
                    self.log_message("✅ Power plan diatur ke High Performance")
                except Exception as e:
                    self.log_message(f"❌ Gagal mengatur power plan: {e}")
                    self.log_message("💡 Coba jalankan sebagai Administrator")
                
                # 4. Optimasi NVIDIA/AMD settings (jika ada)
                self.log_message("\n🎮 Optimasi pengaturan GPU...")
                try:
                    # Cek GPU NVIDIA
                    nvidia_smi = shutil.which('nvidia-smi')
                    if nvidia_smi:
                        self.log_message("\n🟢 GPU NVIDIA ditemukan")
                        self.log_message("💡 Untuk optimasi NVIDIA:")
                        self.log_message("1. Buka NVIDIA Control Panel")
                        self.log_message("2. Pilih 'Manage 3D settings'")
                        self.log_message("3. Atur 'Power management mode' ke 'Prefer maximum performance'")
                        self.log_message("4. Atur 'Texture filtering - Quality' ke 'High performance'")
                    
                    # Cek GPU AMD
                    amd_path = os.path.join(os.environ['ProgramFiles'], 'AMD', 'CNext', 'CNext', 'amdow.exe')
                    if os.path.exists(amd_path):
                        self.log_message("\n🔵 GPU AMD ditemukan")
                        self.log_message("💡 Untuk optimasi AMD:")
                        self.log_message("1. Buka AMD Radeon Settings")
                        self.log_message("2. Pilih 'Gaming' > 'Global Settings'")
                        self.log_message("3. Aktifkan 'Radeon Chill' dan atur FPS target")
                        self.log_message("4. Atur 'Power Efficiency' ke 'Off'")
                    
                    if not nvidia_smi and not os.path.exists(amd_path):
                        self.log_message("⚠ Tidak ditemukan pengaturan GPU khusus")
                except Exception as e:
                    self.log_message(f"❌ Gagal memeriksa pengaturan GPU: {e}")
                
                # 5. Nonaktifkan fullscreen optimizations
                self.log_message("\n🖥️ Menonaktifkan fullscreen optimizations...")
                self.log_message("💡 Untuk game tertentu:")
                self.log_message("1. Cari executable game (.exe)")
                self.log_message("2. Klik kanan > Properties")
                self.log_message("3. Tab Compatibility")
                self.log_message("4. Centang 'Disable fullscreen optimizations'")
                self.log_message("5. Centang 'Run this program as an administrator'")
                
                self.log_message("\n✅ Optimasi gaming selesai")
                self.log_message("💡 Restart game untuk melihat perubahan")
            
            elif self.os_type == "Linux":
                self.log_message("\n🐧 Optimasi gaming di Linux:")
                self.log_message("1. Gunakan GPU driver proprietary (NVIDIA/AMD)")
                self.log_message("2. Install gamemode: sudo apt install gamemode")
                self.log_message("3. Jalankan game dengan: gamemoderun %command%")
                self.log_message("4. Gunakan kernel low-latency: sudo apt install linux-lowlatency")
                self.log_message("5. Atur CPU governor ke performance:")
                self.log_message("   sudo cpupower frequency-set -g performance")
            
            elif self.os_type == "Darwin":
                self.log_message("\n🍏 Optimasi gaming di macOS:")
                self.log_message("1. Tutup semua aplikasi lain sebelum gaming")
                self.log_message("2. Aktifkan 'Increase contrast' di Accessibility settings")
                self.log_message("3. Nonaktifkan transparency effects di System Preferences > Accessibility")
                self.log_message("4. Gunakan resolusi native display untuk performa terbaik")
            
        except Exception as e:
            self.log_message(f"❌ Gagal melakukan optimasi FPS: {e}")

    def word_to_pdf(self):
        """Konversi Word ke PDF"""
        try:
            file_path = filedialog.askopenfilename(
                title="Pilih File Word",
                filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")]
            )
            
            if not file_path:
                return
                
            pdf_path = os.path.splitext(file_path)[0] + ".pdf"
            
            if self.os_type == "Windows":
                # Gunakan Microsoft Word untuk konversi
                self.log_message(f"\nMengkonversi {file_path} ke PDF...")
                word = client.CreateObject("Word.Application")
                doc = word.Documents.Open(file_path)
                doc.SaveAs(pdf_path, FileFormat=17)
                doc.Close()
                word.Quit()
            else:
                # Fallback untuk sistem non-Windows (membutuhkan unoconv)
                self.log_message(f"\nMengkonversi menggunakan unoconv...")
                subprocess.run(['unoconv', '-f', 'pdf', file_path], check=True)
            
            self.log_message(f"✅ Berhasil dikonversi ke: {pdf_path}")
            messagebox.showinfo("Sukses", f"File berhasil dikonversi ke:\n{pdf_path}")
            
        except Exception as e:
            self.log_message(f"❌ Gagal konversi: {e}")
            messagebox.showerror("Error", f"Gagal mengkonversi file:\n{str(e)}")

    def pdf_to_word(self):
        """Konversi PDF ke Word"""
        try:
            file_path = filedialog.askopenfilename(
                title="Pilih File PDF",
                filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")]
            )
            
            if not file_path:
                return
                
            word_path = os.path.splitext(file_path)[0] + ".docx"
            
            self.log_message(f"\nMengkonversi {file_path} ke Word...")
            cv = Converter(file_path)
            cv.convert(word_path)
            cv.close()
            
            self.log_message(f"✅ Berhasil dikonversi ke: {word_path}")
            messagebox.showinfo("Sukses", f"File berhasil dikonversi ke:\n{word_path}")
            
        except Exception as e:
            self.log_message(f"❌ Gagal konversi: {e}")
            messagebox.showerror("Error", f"Gagal mengkonversi file:\n{str(e)}")

    def quick_clean(self):
        """Pembersihan cepat (temp files + browser cache)"""
        self.log_message("\n🚀 Memulai pembersihan cepat...")
        self.clear_temp_files()
        self.clean_browser_cache()
        self.log_message("✅ Pembersihan cepat selesai")
        messagebox.showinfo("Sukses", "Pembersihan cepat selesai!")

    def show_about(self):
        """Menampilkan informasi tentang aplikasi"""
        about_text = (
            "Xboy Optimizer & Converter\n"
            "Versi 1.0\n\n"
            "Aplikasi untuk optimasi sistem dan konversi dokumen\n\n"
            "Fitur:\n"
            "- Pembersihan file temporary\n"
            "- Pembersihan cache browser\n"
            "- Optimasi program startup\n"
            "- Analisis penggunaan resource\n"
            "- Konversi Word ke PDF\n"
            "- Konversi PDF ke Word\n"
        )
        messagebox.showinfo("Tentang", about_text)

    def _bytes_to_gb(self, bytes_val: int) -> float:
        """Konversi bytes ke GB"""
        return bytes_val / (1024 ** 3)

    def _bytes_to_mb(self, bytes_val: int) -> float:
        """Konversi bytes ke MB"""
        return bytes_val / (1024 ** 2)

if __name__ == "__main__":
    root = Tk()
    app = XboyOptimizer(root)
    root.mainloop()