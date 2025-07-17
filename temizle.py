import os
import sys
import shutil
import winreg
import tempfile
import glob
import subprocess
import threading
import time
from pathlib import Path
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
from tkinter import font as tkFont

class ModernOfficeCleanerGUI:
    def __init__(self, root):
        self.root = root
        self.setup_window()
        self.setup_variables()
        self.setup_styles()
        self.create_widgets()
        self.center_window()
        
        # Sistem bilgileri
        self.user_profile = os.environ.get('USERPROFILE', '')
        self.appdata = os.environ.get('APPDATA', '')
        self.localappdata = os.environ.get('LOCALAPPDATA', '')
        self.cleaned_items = []
        self.is_cleaning = False
        
    def setup_window(self):
        """Pencere ayarlarını yapılandır"""
        self.root.title("Office Geçmiş ve Sistem Temizleyici")
        self.root.geometry("800x600")
        self.root.minsize(700, 550)
        self.root.configure(bg='#f0f0f0')
        
        # İkon ayarla (varsa)
        try:
            self.root.iconbitmap('cleaner.ico')
        except:
            pass
    
    def setup_variables(self):
        """Tkinter değişkenlerini ayarla"""
        self.var_recent_docs = tk.BooleanVar(value=True)
        self.var_office_history = tk.BooleanVar(value=True)
        self.var_temp_files = tk.BooleanVar(value=True)
        self.var_browser_cache = tk.BooleanVar(value=False)
        self.var_recycle_bin = tk.BooleanVar(value=False)
        self.var_windows_update = tk.BooleanVar(value=False)
        self.var_system_logs = tk.BooleanVar(value=False)
        self.var_prefetch = tk.BooleanVar(value=False)
        
        self.progress_var = tk.StringVar(value="Hazır")
        self.progress_percent = tk.DoubleVar()
        
    def setup_styles(self):
        """Modern stilleri ayarla"""
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        # Renkler
        self.colors = {
            'primary': '#2c3e50',
            'secondary': '#3498db',
            'success': '#27ae60',
            'danger': '#e74c3c',
            'warning': '#f39c12',
            'light': '#ecf0f1',
            'dark': '#2c3e50',
            'white': '#ffffff'
        }
        
        # Özel stiller
        self.style.configure('Title.TLabel', 
                           font=('Segoe UI', 16, 'bold'),
                           foreground=self.colors['primary'])
        
        self.style.configure('Subtitle.TLabel',
                           font=('Segoe UI', 10),
                           foreground=self.colors['dark'])
        
        self.style.configure('Primary.TButton',
                           font=('Segoe UI', 10, 'bold'),
                           foreground=self.colors['white'])
        
        self.style.map('Primary.TButton',
                      background=[('active', self.colors['primary']),
                                ('!active', self.colors['secondary'])])
        
        self.style.configure('Success.TButton',
                           font=('Segoe UI', 10, 'bold'),
                           foreground=self.colors['white'])
        
        self.style.map('Success.TButton',
                      background=[('active', '#229954'),
                                ('!active', self.colors['success'])])
        
        self.style.configure('Danger.TButton',
                           font=('Segoe UI', 10, 'bold'),
                           foreground=self.colors['white'])
        
        self.style.map('Danger.TButton',
                      background=[('active', '#c0392b'),
                                ('!active', self.colors['danger'])])
    
    def create_widgets(self):
        """Widget'ları oluştur"""
        # Ana container
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Grid ağırlıkları
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Başlık
        title_frame = ttk.Frame(main_frame)
        title_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 20))
        
        ttk.Label(title_frame, text="🧹 Office Geçmiş ve Sistem Temizleyici", 
                 style='Title.TLabel').pack(side=tk.LEFT)
        
        # Yönetici durumu
        admin_status = "👑 Yönetici" if self.is_admin() else "⚠️ Normal Kullanıcı"
        ttk.Label(title_frame, text=admin_status, 
                 style='Subtitle.TLabel').pack(side=tk.RIGHT)
        
        # Sol panel - Seçenekler
        left_frame = ttk.LabelFrame(main_frame, text="Temizlik Seçenekleri", padding="15")
        left_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 10))
        
        # Sağ panel - Log ve İlerleme
        right_frame = ttk.LabelFrame(main_frame, text="İşlem Durumu", padding="15")
        right_frame.grid(row=1, column=1, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Grid ağırlıkları
        main_frame.rowconfigure(1, weight=1)
        left_frame.columnconfigure(0, weight=1)
        right_frame.columnconfigure(0, weight=1)
        right_frame.rowconfigure(1, weight=1)
        
        self.create_left_panel(left_frame)
        self.create_right_panel(right_frame)
        
        # Alt panel - Butonlar
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(20, 0))
        
        self.create_buttons(button_frame)
    
    def create_left_panel(self, parent):
        """Sol panel - seçenekler"""
        # Office seçenekleri
        office_frame = ttk.LabelFrame(parent, text="Office Temizlik", padding="10")
        office_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Checkbutton(office_frame, text="📄 Son Kullanılan Belgeler", 
                       variable=self.var_recent_docs).pack(anchor=tk.W, pady=2)
        ttk.Checkbutton(office_frame, text="📈 Office Uygulama Geçmişi", 
                       variable=self.var_office_history).pack(anchor=tk.W, pady=2)
        
        # Sistem seçenekleri
        system_frame = ttk.LabelFrame(parent, text="Sistem Temizlik", padding="10")
        system_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Checkbutton(system_frame, text="🗂️ Geçici Dosyalar", 
                       variable=self.var_temp_files).pack(anchor=tk.W, pady=2)
        ttk.Checkbutton(system_frame, text="🌐 Tarayıcı Önbelleği", 
                       variable=self.var_browser_cache).pack(anchor=tk.W, pady=2)
        ttk.Checkbutton(system_frame, text="🗑️ Geri Dönüşüm Kutusu", 
                       variable=self.var_recycle_bin).pack(anchor=tk.W, pady=2)
        
        # Gelişmiş seçenekler
        advanced_frame = ttk.LabelFrame(parent, text="Gelişmiş Temizlik", padding="10")
        advanced_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Checkbutton(advanced_frame, text="🔄 Windows Update Önbelleği", 
                       variable=self.var_windows_update).pack(anchor=tk.W, pady=2)
        ttk.Checkbutton(advanced_frame, text="📋 Sistem Logları", 
                       variable=self.var_system_logs).pack(anchor=tk.W, pady=2)
        ttk.Checkbutton(advanced_frame, text="⚡ Prefetch Dosyaları", 
                       variable=self.var_prefetch).pack(anchor=tk.W, pady=2)
        
        # Hızlı seçim butonları
        quick_frame = ttk.Frame(parent)
        quick_frame.pack(fill=tk.X, pady=(10, 0))
        
        ttk.Button(quick_frame, text="✓ Hepsini Seç", 
                  command=self.select_all).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(quick_frame, text="✗ Hiçbirini Seçme", 
                  command=self.select_none).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(quick_frame, text="💼 Sadece Office", 
                  command=self.select_office_only).pack(side=tk.LEFT)
    
    def create_right_panel(self, parent):
        """Sağ panel - log ve ilerleme"""
        # İlerleme durumu
        progress_frame = ttk.Frame(parent)
        progress_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(progress_frame, text="Durum:").pack(side=tk.LEFT)
        self.status_label = ttk.Label(progress_frame, textvariable=self.progress_var, 
                                     style='Subtitle.TLabel')
        self.status_label.pack(side=tk.LEFT, padx=(5, 0))
        
        # İlerleme çubuğu
        self.progress_bar = ttk.Progressbar(parent, variable=self.progress_percent, 
                                          maximum=100, length=300)
        self.progress_bar.pack(fill=tk.X, pady=(0, 10))
        
        # Log alanı
        log_frame = ttk.Frame(parent)
        log_frame.pack(fill=tk.BOTH, expand=True)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, 
                                                 height=15, 
                                                 font=('Consolas', 9),
                                                 wrap=tk.WORD)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Log renklendirme
        self.log_text.tag_config('success', foreground=self.colors['success'])
        self.log_text.tag_config('error', foreground=self.colors['danger'])
        self.log_text.tag_config('warning', foreground=self.colors['warning'])
        self.log_text.tag_config('info', foreground=self.colors['secondary'])
        
        # Başlangıç mesajı
        self.log_message("Office Geçmiş ve Sistem Temizleyici hazır", "info")
        if not self.is_admin():
            self.log_message("⚠️ Yönetici yetkisi olmadan bazı işlemler yapılamayabilir", "warning")
    
    def create_buttons(self, parent):
        """Butonları oluştur"""
        # Sol taraf - ana butonlar
        left_buttons = ttk.Frame(parent)
        left_buttons.pack(side=tk.LEFT)
        
        self.start_button = ttk.Button(left_buttons, text="🚀 Temizliği Başlat", 
                                      style='Success.TButton',
                                      command=self.start_cleaning)
        self.start_button.pack(side=tk.LEFT, padx=(0, 10))
        
        self.stop_button = ttk.Button(left_buttons, text="⏹️ Durdur", 
                                     style='Danger.TButton',
                                     command=self.stop_cleaning,
                                     state=tk.DISABLED)
        self.stop_button.pack(side=tk.LEFT, padx=(0, 10))
        
        # Sağ taraf - yardımcı butonlar
        right_buttons = ttk.Frame(parent)
        right_buttons.pack(side=tk.RIGHT)
        
        ttk.Button(right_buttons, text="📋 Logu Temizle", 
                  command=self.clear_log).pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Button(right_buttons, text="💾 Logu Kaydet", 
                  command=self.save_log).pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Button(right_buttons, text="ℹ️ Hakkında", 
                  command=self.show_about).pack(side=tk.LEFT)
    
    def center_window(self):
        """Pencereyi ekranın ortasına yerleştir"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
    
    def log_message(self, message, level="info"):
        """Log mesajı ekle"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        formatted_message = f"[{timestamp}] {message}\n"
        
        self.log_text.insert(tk.END, formatted_message, level)
        self.log_text.see(tk.END)
        self.root.update_idletasks()
    
    def update_progress(self, percent, status):
        """İlerleme durumunu güncelle"""
        self.progress_percent.set(percent)
        self.progress_var.set(status)
        self.root.update_idletasks()
    
    def select_all(self):
        """Tüm seçenekleri seç"""
        for var in [self.var_recent_docs, self.var_office_history, self.var_temp_files,
                   self.var_browser_cache, self.var_recycle_bin, self.var_windows_update,
                   self.var_system_logs, self.var_prefetch]:
            var.set(True)
    
    def select_none(self):
        """Hiçbir seçeneği seçme"""
        for var in [self.var_recent_docs, self.var_office_history, self.var_temp_files,
                   self.var_browser_cache, self.var_recycle_bin, self.var_windows_update,
                   self.var_system_logs, self.var_prefetch]:
            var.set(False)
    
    def select_office_only(self):
        """Sadece Office seçeneklerini seç"""
        self.select_none()
        self.var_recent_docs.set(True)
        self.var_office_history.set(True)
    
    def clear_log(self):
        """Log alanını temizle"""
        self.log_text.delete(1.0, tk.END)
        self.log_message("Log temizlendi", "info")
    
    def save_log(self):
        """Logu dosyaya kaydet"""
        from tkinter import filedialog
        filename = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
            title="Log Dosyasını Kaydet"
        )
        if filename:
            try:
                with open(filename, 'w', encoding='utf-8') as f:
                    f.write(self.log_text.get(1.0, tk.END))
                self.log_message(f"Log kaydedildi: {filename}", "success")
            except Exception as e:
                self.log_message(f"Log kaydedilirken hata: {e}", "error")
    
    def show_about(self):
        """Hakkında penceresi"""
        messagebox.showinfo("Hakkında", 
                           "Office Geçmiş ve Sistem Temizleyici\n"
                           "Önder AKÖZ Sürüm: 2.0\n"
                           "Modern GUI ile Windows sistem temizliği\n\n"
                           "Özellikler:\n"
                           "• Office dosya geçmişi temizliği\n"
                           "• Sistem geçici dosyalarını temizleme\n"
                           "• Tarayıcı önbellek temizliği\n"
                           "• Gerçek zamanlı ilerleme takibi\n"
                           "• Detaylı log raporları")
    
    def start_cleaning(self):
        """Temizlik işlemini başlat"""
        if self.is_cleaning:
            return
        
        # Seçili işlemleri kontrol et
        selected_tasks = []
        if self.var_recent_docs.get():
            selected_tasks.append("recent_docs")
        if self.var_office_history.get():
            selected_tasks.append("office_history")
        if self.var_temp_files.get():
            selected_tasks.append("temp_files")
        if self.var_browser_cache.get():
            selected_tasks.append("browser_cache")
        if self.var_recycle_bin.get():
            selected_tasks.append("recycle_bin")
        if self.var_windows_update.get():
            selected_tasks.append("windows_update")
        if self.var_system_logs.get():
            selected_tasks.append("system_logs")
        if self.var_prefetch.get():
            selected_tasks.append("prefetch")
        
        if not selected_tasks:
            messagebox.showwarning("Uyarı", "Lütfen en az bir temizlik seçeneği seçin!")
            return
        
        # Onay al
        if not messagebox.askyesno("Onay", 
                                  f"{len(selected_tasks)} temizlik işlemi başlatılacak.\n"
                                  "Devam etmek istiyor musunuz?"):
            return
        
        # UI'yi güncelle
        self.is_cleaning = True
        self.start_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)
        
        # Temizlik thread'ini başlat
        self.cleaning_thread = threading.Thread(target=self.run_cleaning, args=(selected_tasks,))
        self.cleaning_thread.daemon = True
        self.cleaning_thread.start()
    
    def stop_cleaning(self):
        """Temizlik işlemini durdur"""
        self.is_cleaning = False
        self.log_message("Temizlik işlemi kullanıcı tarafından durduruldu", "warning")
        self.cleanup_ui()
    
    def cleanup_ui(self):
        """UI'yi temizlik sonrası duruma getir"""
        self.is_cleaning = False
        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
        self.update_progress(0, "Hazır")
    
    def run_cleaning(self, tasks):
        """Temizlik işlemlerini çalıştır"""
        total_tasks = len(tasks)
        completed = 0
        
        self.log_message("🚀 Temizlik işlemi başlatıldı", "info")
        
        task_functions = {
            "recent_docs": self.clean_recent_documents,
            "office_history": self.clean_office_history,
            "temp_files": self.clean_temp_files,
            "browser_cache": self.clean_browser_cache,
            "recycle_bin": self.clean_recycle_bin,
            "windows_update": self.clean_windows_update_cache,
            "system_logs": self.clean_system_logs,
            "prefetch": self.clean_prefetch
        }
        
        for task in tasks:
            if not self.is_cleaning:
                break
            
            if task in task_functions:
                try:
                    task_functions[task]()
                    completed += 1
                    progress = (completed / total_tasks) * 100
                    self.update_progress(progress, f"Tamamlanan: {completed}/{total_tasks}")
                except Exception as e:
                    self.log_message(f"Hata: {task} - {e}", "error")
        
        if self.is_cleaning:
            self.log_message("✅ Tüm temizlik işlemleri tamamlandı!", "success")
            self.update_progress(100, "Tamamlandı")
            messagebox.showinfo("Başarılı", "Temizlik işlemleri başarıyla tamamlandı!")
        
        self.cleanup_ui()
    
    def clean_recycle_bin(self):
        """Geri dönüşüm kutusunu temizle"""
        self.log_message("🗑️ Geri dönüşüm kutusu temizleniyor...", "info")

        # Yöntem 1: PowerShell komutu
        try:
            subprocess.run([
                "powershell", "-ExecutionPolicy", "Bypass", "-Command",
                "Clear-RecycleBin -Force -Confirm:$false"
            ], check=True, capture_output=True, timeout=30)
            self.log_message("✓ Geri dönüşüm kutusu temizlendi (PowerShell)", "success")
            return
        except Exception as e:
            self.log_message(f"⚠️ PowerShell metodu başarısız: {str(e)[:100]}", "warning")

        # Yöntem 2: CMD komutu
        try:
            subprocess.run([
                "cmd", "/c", "rd /s /q %systemdrive%\\$Recycle.Bin"
            ], check=True, capture_output=True, timeout=30)
            self.log_message("✓ Geri dönüşüm kutusu temizlendi (CMD)", "success")
            return
        except Exception as e:
            self.log_message(f"⚠️ CMD metodu başarısız: {str(e)[:100]}", "warning")

        # Yöntem 3: Python ile manuel temizlik
        try:
            import glob
            recycle_paths = []

            # Tüm sürücülerde $Recycle.Bin klasörlerini bul
            for drive in ['C:', 'D:', 'E:', 'F:', 'G:', 'H:', 'I:', 'J:', 'K:', 'L:', 'M:', 'N:', 'O:', 'P:', 'Q:', 'R:', 'S:', 'T:', 'U:', 'V:', 'W:', 'X:', 'Y:', 'Z:']:
                recycle_path = os.path.join(drive, os.sep, '$Recycle.Bin')
                if os.path.exists(recycle_path):
                    recycle_paths.append(recycle_path)

            cleaned_items = 0
            for recycle_path in recycle_paths:
                try:
                    for root, dirs, files in os.walk(recycle_path):
                        # Dosyaları sil
                        for file in files: # Bu 'for' döngüsünün gövdesi için girinti eksikti.
                            try:
                                file_path = os.path.join(root, file)
                                os.remove(file_path)
                                cleaned_items += 1
                            except Exception: # Mümkünse belirli bir istisna yakalayın veya loglayın
                                continue

                        # Boş klasörleri sil
                        for dir_name in dirs: # Yerleşik olanla çakışmaması için 'dir' adı 'dir_name' olarak değiştirildi. Bu 'for' döngüsünün gövdesi için girinti eksikti.
                            try:
                                dir_path = os.path.join(root, dir_name)
                                if not os.listdir(dir_path):  # Boşsa
                                    os.rmdir(dir_path)
                                    cleaned_items += 1
                            except Exception: # Mümkünse belirli bir istisna yakalayın veya loglayın
                                continue
                except Exception: # Mümkünse belirli bir istisna yakalayın veya loglayın
                    continue

            if cleaned_items > 0:
                self.log_message(f"✓ Geri dönüşüm kutusu temizlendi ({cleaned_items} öğe - Python)", "success")
            else:
                self.log_message("ℹ️ Geri dönüşüm kutusu zaten boş", "info")

        except Exception as e:
            self.log_message(f"✗ Geri dönüşüm kutusu temizlenirken hata: {str(e)[:100]}", "error")
        # clean_recycle_bin fonksiyonunun geri kalanı (Yöntem 4) burada, doğru şekilde girintilenmiş olarak devam etmelidir.
        # ... (clean_recycle_bin fonksiyonunun geri kalanı)

    def safe_delete(self, path, item_name):
        """Güvenli silme işlemi"""
        try:
            if os.path.exists(path):
                if os.path.isdir(path):
                    shutil.rmtree(path)
                else:
                    os.remove(path)
                self.log_message(f"✓ {item_name} temizlendi", "success")
                return True
            else:
                self.log_message(f"⚠️ {item_name} bulunamadı", "warning")
                return False
        except Exception as e:
            self.log_message(f"✗ {item_name} temizlenirken hata: {e}", "error")
            return False
    
    def clean_recent_documents(self):
        """Son kullanılan belgeler listesini temizle"""
        self.log_message("📄 Son kullanılan belgeler temizleniyor...", "info")
        
        recent_folder = os.path.join(self.appdata, "Microsoft", "Windows", "Recent")
        if os.path.exists(recent_folder):
            try:
                for file in os.listdir(recent_folder):
                    if not self.is_cleaning:
                        break
                    file_path = os.path.join(recent_folder, file)
                    if os.path.isfile(file_path):
                        os.remove(file_path)
                self.log_message("✓ Recent klasörü temizlendi", "success")
            except Exception as e:
                self.log_message(f"✗ Recent klasörü temizlenirken hata: {e}", "error")
    
    def clean_office_history(self):
        """Office uygulamalarının geçmişini temizle"""
        self.log_message("📈 Office geçmişi temizleniyor...", "info")
        
        # Excel için özel temizlik işlemleri
        self.clean_excel_recent_files()
        
        # Word için özel temizlik işlemleri
        self.clean_word_recent_files()
        
        # Office 365 cloud temizliği
        self.clean_office365_cloud_history()
        
        # Diğer Office uygulamaları için genel temizlik
        office_apps = {
            "PowerPoint": [
                r"Software\Microsoft\Office\16.0\PowerPoint\User MRU",
                r"Software\Microsoft\Office\15.0\PowerPoint\User MRU",
                r"Software\Microsoft\Office\14.0\PowerPoint\User MRU",
                r"Software\Microsoft\Office\16.0\PowerPoint\File MRU",
                r"Software\Microsoft\Office\15.0\PowerPoint\File MRU",
                r"Software\Microsoft\Office\14.0\PowerPoint\File MRU",
                r"Software\Microsoft\Office\16.0\PowerPoint\Recent Files",
                r"Software\Microsoft\Office\15.0\PowerPoint\Recent Files",
                r"Software\Microsoft\Office\14.0\PowerPoint\Recent Files",
                r"Software\Microsoft\Office\16.0\PowerPoint\Web Service Cache",
                r"Software\Microsoft\Office\16.0\PowerPoint\SharePoint",
                r"Software\Microsoft\Office\16.0\PowerPoint\OneDrive",
            ],
            "Access": [
                r"Software\Microsoft\Office\16.0\Access\User MRU",
                r"Software\Microsoft\Office\15.0\Access\User MRU",
                r"Software\Microsoft\Office\14.0\Access\User MRU",
                r"Software\Microsoft\Office\16.0\Access\File MRU",
                r"Software\Microsoft\Office\15.0\Access\File MRU",
                r"Software\Microsoft\Office\14.0\Access\File MRU",
            ]
        }
        
        for app_name, reg_paths in office_apps.items():
            if not self.is_cleaning:
                break
            app_cleaned = False
            cleaned_count = 0
            
            for reg_path in reg_paths:
                try:
                    # Anahtarı açmaya çalış
                    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, reg_path, 0, winreg.KEY_ALL_ACCESS) as key:
                        # Tüm alt anahtarları sil
                        subkeys_to_delete = []
                        try:
                            i = 0
                            while True:
                                subkey = winreg.EnumKey(key, i)
                                subkeys_to_delete.append(subkey)
                                i += 1
                        except WindowsError:
                            pass
                        
                        # Subkey'leri sil
                        for subkey in subkeys_to_delete:
                            try:
                                winreg.DeleteKey(key, subkey)
                                cleaned_count += 1
                            except:
                                pass
                        
                        # Değerleri de sil
                        values_to_delete = []
                        try:
                            i = 0
                            while True:
                                value_name, value_data, _ = winreg.EnumValue(key, i)
                                # Recent files ile ilgili değerleri tespit et
                                if any(keyword in value_name.lower() for keyword in ['recent', 'mru', 'file', 'path', 'document']):
                                    values_to_delete.append(value_name)
                                elif isinstance(value_data, str) and any(ext in value_data.lower() for ext in ['.ppt', '.pptx', '.pptm', '.mdb', '.accdb']):
                                    values_to_delete.append(value_name)
                                # SharePoint/OneDrive dosya referansları
                                elif isinstance(value_data, str) and any(keyword in value_data.lower() for keyword in ['sharepoint', 'onedrive', 'https://', 'my.sharepoint.com']):
                                    values_to_delete.append(value_name)
                                i += 1
                        except WindowsError:
                            pass
                        
                        for value_name in values_to_delete:
                            try:
                                winreg.DeleteValue(key, value_name)
                                cleaned_count += 1
                            except:
                                pass
                        
                        app_cleaned = True
                except Exception as e:
                    continue
            
            if app_cleaned and cleaned_count > 0:
                self.log_message(f"✓ {app_name} geçmişi temizlendi ({cleaned_count} kayıt)", "success")
            elif app_cleaned:
                self.log_message(f"✓ {app_name} geçmişi temizlendi", "success")
    
    def clean_office365_cloud_history(self):
        """Office 365 cloud geçmişini temizle"""
        self.log_message("☁️ Office 365 cloud geçmişi temizleniyor...", "info")
        
        # Office 365 cloud cache konumları
        cloud_cache_locations = [
            # Microsoft Graph cache
            os.path.join(self.localappdata, "Microsoft", "Office", "16.0", "ClientTelemetry"),
            os.path.join(self.localappdata, "Microsoft", "Office", "16.0", "RoamingOfficeData"),
            os.path.join(self.localappdata, "Microsoft", "Office", "16.0", "WebServiceCache"),
            
            # OneDrive integration cache
            os.path.join(self.localappdata, "Microsoft", "OneDrive", "cache"),
            os.path.join(self.localappdata, "Microsoft", "OneDrive", "logs"),
            os.path.join(self.appdata, "Microsoft", "OneDrive", "settings"),
            
            # SharePoint cache
            os.path.join(self.localappdata, "Microsoft", "SharePoint Designer"),
            os.path.join(self.appdata, "Microsoft", "SharePoint"),
            
            # Teams integration (Office entegrasyonu için)
            os.path.join(self.appdata, "Microsoft", "Teams", "Application Cache"),
            os.path.join(self.appdata, "Microsoft", "Teams", "Cache"),
            
            # Office roaming settings
            os.path.join(self.localappdata, "Microsoft", "Office", "16.0", "roaming"),
            os.path.join(self.appdata, "Microsoft", "Office", "16.0", "roaming"),
        ]
        
        cleaned_files = 0
        for cache_location in cloud_cache_locations:
            if not self.is_cleaning:
                break
            if os.path.exists(cache_location):
                try:
                    for root, dirs, files in os.walk(cache_location):
                        for file in files:
                            # Cloud ile ilgili cache dosyalarını temizle
                            if any(keyword in file.lower() for keyword in [
                                'recent', 'mru', 'cache', 'temp', 'log',
                                'sharepoint', 'onedrive', 'teams', 'graph',
                                '.json', '.xml', '.tmp', '.log'
                            ]):
                                file_path = os.path.join(root, file)
                                try:
                                    os.remove(file_path)
                                    cleaned_files += 1
                                except:
                                    pass
                except Exception as e:
                    continue
        
        if cleaned_files > 0:
            self.log_message(f"✓ Office 365 cloud cache temizlendi ({cleaned_files} dosya)", "success")
        
        # Registry'deki Office 365 cloud kayıtları
        office365_registry_locations = [
            r"Software\Microsoft\Office\16.0\Common\Roaming",
            r"Software\Microsoft\Office\16.0\Common\Internet",
            r"Software\Microsoft\Office\16.0\Common\Identity",
            r"Software\Microsoft\Office\16.0\Common\Experiment",
            r"Software\Microsoft\OneDrive",
            r"Software\Microsoft\SharePoint",
        ]
        
        registry_cleaned = 0
        for reg_location in office365_registry_locations:
            if not self.is_cleaning:
                break
            try:
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, reg_location, 0, winreg.KEY_ALL_ACCESS) as key:
                    # Recent ve cache ile ilgili değerleri temizle
                    values_to_delete = []
                    try:
                        i = 0
                        while True:
                            value_name, value_data, _ = winreg.EnumValue(key, i)
                            if any(keyword in value_name.lower() for keyword in ['recent', 'cache', 'temp', 'mru']):
                                values_to_delete.append(value_name)
                            elif isinstance(value_data, str) and any(keyword in value_data.lower() for keyword in ['sharepoint', 'onedrive', 'graph.microsoft.com']):
                                values_to_delete.append(value_name)
                            i += 1
                    except WindowsError:
                        pass
                    
                    for value_name in values_to_delete:
                        try:
                            winreg.DeleteValue(key, value_name)
                            registry_cleaned += 1
                        except:
                            pass
            except Exception as e:
                continue
        
        if registry_cleaned > 0:
            self.log_message(f"✓ Office 365 registry kayıtları temizlendi ({registry_cleaned} kayıt)", "success")
    
    def clean_word_recent_files(self):
        """Word'ün son dosyalar listesini özel olarak temizle"""
        self.log_message("📝 Word son dosyalar listesi temizleniyor...", "info")
        
        # Word'ün TÜMU konumlardaki recent file kayıtları
        word_locations = [
            # Office 2016/2019/365 (16.0) - Kapsamlı liste
            r"Software\Microsoft\Office\16.0\Word\Recent Files",
            r"Software\Microsoft\Office\16.0\Word\User MRU",
            r"Software\Microsoft\Office\16.0\Word\File MRU",
            r"Software\Microsoft\Office\16.0\Word\Place MRU",
            r"Software\Microsoft\Office\16.0\Word\Security\Trusted Documents",
            r"Software\Microsoft\Office\16.0\Word\Options",
            r"Software\Microsoft\Office\16.0\Word\Data",
            r"Software\Microsoft\Office\16.0\Common\Open Find\Microsoft Office Word\Settings",
            r"Software\Microsoft\Office\16.0\Common\General",
            
            # Office 365 Cloud/SharePoint kayıtları
            r"Software\Microsoft\Office\16.0\Word\Web Service Cache",
            r"Software\Microsoft\Office\16.0\Word\SharePoint",
            r"Software\Microsoft\Office\16.0\Word\OneDrive",
            r"Software\Microsoft\Office\16.0\Common\Internet",
            r"Software\Microsoft\Office\16.0\Common\Roaming",
            r"Software\Microsoft\Office\16.0\Common\Identity",
            
            # Office 2013 (15.0)
            r"Software\Microsoft\Office\15.0\Word\Recent Files",
            r"Software\Microsoft\Office\15.0\Word\User MRU",
            r"Software\Microsoft\Office\15.0\Word\File MRU",
            r"Software\Microsoft\Office\15.0\Word\Place MRU",
            r"Software\Microsoft\Office\15.0\Word\Security\Trusted Documents",
            r"Software\Microsoft\Office\15.0\Word\Options",
            r"Software\Microsoft\Office\15.0\Word\Data",
            r"Software\Microsoft\Office\15.0\Common\Open Find\Microsoft Office Word\Settings",
            r"Software\Microsoft\Office\15.0\Word\Web Service Cache",
            r"Software\Microsoft\Office\15.0\Word\SharePoint",
            
            # Office 2010 (14.0)
            r"Software\Microsoft\Office\14.0\Word\Recent Files",
            r"Software\Microsoft\Office\14.0\Word\User MRU",
            r"Software\Microsoft\Office\14.0\Word\File MRU",
            r"Software\Microsoft\Office\14.0\Word\Place MRU",
            r"Software\Microsoft\Office\14.0\Word\Security\Trusted Documents",
            r"Software\Microsoft\Office\14.0\Word\Options",
            r"Software\Microsoft\Office\14.0\Word\Data",
            
            # Eski Office sürümleri
            r"Software\Microsoft\Office\12.0\Word\Recent Files",
            r"Software\Microsoft\Office\11.0\Word\Recent Files",
        ]
        
        cleaned_count = 0
        
        # Registry temizliği
        for reg_path in word_locations:
            if not self.is_cleaning:
                break
            try:
                # Registry anahtarını aç
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, reg_path, 0, winreg.KEY_ALL_ACCESS) as key:
                    # Tüm değerleri listele ve sil
                    values_to_delete = []
                    try:
                        i = 0
                        while True:
                            value_name, value_data, _ = winreg.EnumValue(key, i)
                            # Recent files ile ilgili değerleri tespit et
                            if any(keyword in value_name.lower() for keyword in ['recent', 'mru', 'file', 'path', 'document']):
                                values_to_delete.append(value_name)
                            elif isinstance(value_data, str) and any(ext in value_data.lower() for ext in ['.docx', '.doc', '.docm', '.dot', '.dotx']):
                                values_to_delete.append(value_name)
                            # SharePoint/OneDrive dosya referansları
                            elif isinstance(value_data, str) and any(keyword in value_data.lower() for keyword in ['sharepoint', 'onedrive', 'https://', 'my.sharepoint.com']):
                                values_to_delete.append(value_name)
                            i += 1
                    except WindowsError:
                        pass
                    
                    # Değerleri sil
                    for value_name in values_to_delete:
                        try:
                            winreg.DeleteValue(key, value_name)
                            cleaned_count += 1
                        except:
                            pass
                    
                    # Alt anahtarları da sil
                    subkeys_to_delete = []
                    try:
                        i = 0
                        while True:
                            subkey = winreg.EnumKey(key, i)
                            subkeys_to_delete.append(subkey)
                            i += 1
                    except WindowsError:
                        pass
                    
                    for subkey in subkeys_to_delete:
                        try:
                            self.delete_registry_key_recursive(winreg.HKEY_CURRENT_USER, f"{reg_path}\\{subkey}")
                            cleaned_count += 1
                        except:
                            pass
                            
            except Exception as e:
                continue
        
        # Word dosya geçmişini de temizle
        self.clean_word_file_history()
        
        # Word OneDrive/SharePoint cache temizliği
        self.clean_word_cloud_cache()
        
        if cleaned_count > 0:
            self.log_message(f"✓ Word'den {cleaned_count} recent file kaydı temizlendi", "success")
        else:
            self.log_message("⚠️ Word recent files bulunamadı veya zaten temiz", "warning")
    
    def clean_word_cloud_cache(self):
        """Word'ün OneDrive/SharePoint cache dosyalarını temizle"""
        self.log_message("☁️ Word cloud cache temizleniyor...", "info")
        
        # Word OneDrive cache konumları
        word_cache_paths = [
            os.path.join(self.localappdata, "Microsoft", "Office", "16.0", "Wef"),
            os.path.join(self.localappdata, "Microsoft", "Office", "15.0", "Wef"),
            os.path.join(self.localappdata, "Microsoft", "Office", "14.0", "Wef"),
            os.path.join(self.appdata, "Microsoft", "Office", "16.0", "roaming"),
            os.path.join(self.appdata, "Microsoft", "Office", "15.0", "roaming"),
            os.path.join(self.localappdata, "Microsoft", "Office", "16.0", "roaming"),
            os.path.join(self.localappdata, "Microsoft", "Office", "15.0", "roaming"),
        ]
        
        for cache_path in word_cache_paths:
            if not self.is_cleaning:
                break
            if os.path.exists(cache_path):
                try:
                    for root, dirs, files in os.walk(cache_path):
                        for file in files:
                            if any(keyword in file.lower() for keyword in ['recent', 'mru', 'cache', 'word', '.json', '.xml']):
                                file_path = os.path.join(root, file)
                                try:
                                    os.remove(file_path)
                                    self.log_message(f"✓ Word cloud cache temizlendi: {file}", "success")
                                except:
                                    pass
                except Exception as e:
                    continue
    
    def clean_word_file_history(self):
        """Word'ün dosya geçmişini AppData'dan temizle"""
        try:
            # Word'ün yerel ayar dosyaları
            word_roaming_path = os.path.join(self.appdata, "Microsoft", "Word")
            word_local_path = os.path.join(self.localappdata, "Microsoft", "Office")
            
            # Roaming Word klasörü
            if os.path.exists(word_roaming_path):
                for file in os.listdir(word_roaming_path):
                    if file.endswith('.officeUI') or 'recent' in file.lower():
                        file_path = os.path.join(word_roaming_path, file)
                        try:
                            os.remove(file_path)
                            self.log_message(f"✓ Word config dosyası temizlendi: {file}", "success")
                        except:
                            pass
            
            # Local Word cache
            if os.path.exists(word_local_path):
                for root, dirs, files in os.walk(word_local_path):
                    for file in files:
                        if any(keyword in file.lower() for keyword in ['recent', 'mru', 'cache']) and 'word' in root.lower():
                            file_path = os.path.join(root, file)
                            try:
                                os.remove(file_path)
                                self.log_message(f"✓ Word cache dosyası temizlendi: {file}", "success")
                            except:
                                pass
        except Exception as e:
            self.log_message(f"✗ Word dosya geçmişi temizlenirken hata: {e}", "error")
    
    def clean_excel_recent_files(self):
        """Excel'in son dosyalar listesini özel olarak temizle"""
        self.log_message("📊 Excel son dosyalar listesi temizleniyor...", "info")
        
        # Excel'in TÜMU konumlardaki recent file kayıtları
        excel_locations = [
            # Office 2016/2019/365 (16.0) - Kapsamlı liste
            r"Software\Microsoft\Office\16.0\Excel\Recent Files",
            r"Software\Microsoft\Office\16.0\Excel\User MRU",
            r"Software\Microsoft\Office\16.0\Excel\File MRU",
            r"Software\Microsoft\Office\16.0\Excel\Place MRU",
            r"Software\Microsoft\Office\16.0\Excel\Security\Trusted Documents",
            r"Software\Microsoft\Office\16.0\Excel\Options",
            r"Software\Microsoft\Office\16.0\Common\Open Find\Microsoft Office Excel\Settings",
            r"Software\Microsoft\Office\16.0\Common\General",
            
            # Office 365 Cloud/SharePoint kayıtları
            r"Software\Microsoft\Office\16.0\Excel\Web Service Cache",
            r"Software\Microsoft\Office\16.0\Excel\SharePoint",
            r"Software\Microsoft\Office\16.0\Excel\OneDrive",
            r"Software\Microsoft\Office\16.0\Common\Internet",
            r"Software\Microsoft\Office\16.0\Common\Roaming",
            r"Software\Microsoft\Office\16.0\Common\Identity",
            
            # Office 2013 (15.0)
            r"Software\Microsoft\Office\15.0\Excel\Recent Files",
            r"Software\Microsoft\Office\15.0\Excel\User MRU",
            r"Software\Microsoft\Office\15.0\Excel\File MRU",
            r"Software\Microsoft\Office\15.0\Excel\Place MRU",
            r"Software\Microsoft\Office\15.0\Excel\Security\Trusted Documents",
            r"Software\Microsoft\Office\15.0\Excel\Options",
            r"Software\Microsoft\Office\15.0\Common\Open Find\Microsoft Office Excel\Settings",
            r"Software\Microsoft\Office\15.0\Excel\Web Service Cache",
            r"Software\Microsoft\Office\15.0\Excel\SharePoint",
            
            # Office 2010 (14.0)
            r"Software\Microsoft\Office\14.0\Excel\Recent Files",
            r"Software\Microsoft\Office\14.0\Excel\User MRU",
            r"Software\Microsoft\Office\14.0\Excel\File MRU",
            r"Software\Microsoft\Office\14.0\Excel\Place MRU",
            r"Software\Microsoft\Office\14.0\Excel\Security\Trusted Documents",
            r"Software\Microsoft\Office\14.0\Excel\Options",
            
            # Eski Office sürümleri
            r"Software\Microsoft\Office\12.0\Excel\Recent Files",
            r"Software\Microsoft\Office\11.0\Excel\Recent Files",
        ]
        
        cleaned_count = 0
        
        # Registry temizliği
        for reg_path in excel_locations:
            if not self.is_cleaning:
                break
            try:
                # Registry anahtarını aç
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, reg_path, 0, winreg.KEY_ALL_ACCESS) as key:
                    # Tüm değerleri listele ve sil
                    values_to_delete = []
                    try:
                        i = 0
                        while True:
                            value_name, value_data, _ = winreg.EnumValue(key, i)
                            # Recent files ile ilgili değerleri tespit et
                            if any(keyword in value_name.lower() for keyword in ['recent', 'mru', 'file', 'path', 'document']):
                                values_to_delete.append(value_name)
                            elif isinstance(value_data, str) and any(ext in value_data.lower() for ext in ['.xlsx', '.xls', '.xlsm', '.xlsb']):
                                values_to_delete.append(value_name)
                            # SharePoint/OneDrive dosya referansları
                            elif isinstance(value_data, str) and any(keyword in value_data.lower() for keyword in ['sharepoint', 'onedrive', 'https://', 'my.sharepoint.com']):
                                values_to_delete.append(value_name)
                            i += 1
                    except WindowsError:
                        pass
                    
                    # Değerleri sil
                    for value_name in values_to_delete:
                        try:
                            winreg.DeleteValue(key, value_name)
                            cleaned_count += 1
                        except:
                            pass
                    
                    # Alt anahtarları da sil
                    subkeys_to_delete = []
                    try:
                        i = 0
                        while True:
                            subkey = winreg.EnumKey(key, i)
                            subkeys_to_delete.append(subkey)
                            i += 1
                    except WindowsError:
                        pass
                    
                    for subkey in subkeys_to_delete:
                        try:
                            self.delete_registry_key_recursive(winreg.HKEY_CURRENT_USER, f"{reg_path}\\{subkey}")
                            cleaned_count += 1
                        except:
                            pass
                            
            except Exception as e:
                continue
        
        # Excel dosya geçmişini de temizle
        self.clean_excel_file_history()
        
        # Excel OneDrive/SharePoint cache temizliği
        self.clean_excel_cloud_cache()
        
        if cleaned_count > 0:
            self.log_message(f"✓ Excel'den {cleaned_count} recent file kaydı temizlendi", "success")
        else:
            self.log_message("⚠️ Excel recent files bulunamadı veya zaten temiz", "warning")
    
    def clean_excel_cloud_cache(self):
        """Excel'in OneDrive/SharePoint cache dosyalarını temizle"""
        self.log_message("☁️ Excel cloud cache temizleniyor...", "info")
        
        # OneDrive cache konumları
        onedrive_cache_paths = [
            os.path.join(self.localappdata, "Microsoft", "OneDrive", "logs"),
            os.path.join(self.localappdata, "Microsoft", "OneDrive", "settings"),
            os.path.join(self.appdata, "Microsoft", "OneDrive", "logs"),
            os.path.join(self.appdata, "Microsoft", "SharePoint"),
            os.path.join(self.localappdata, "Microsoft", "Office", "16.0", "Wef"),
            os.path.join(self.localappdata, "Microsoft", "Office", "15.0", "Wef"),
            os.path.join(self.localappdata, "Microsoft", "Office", "14.0", "Wef"),
        ]
        
        # Office 365 roaming settings
        office365_paths = [
            os.path.join(self.localappdata, "Microsoft", "Office", "16.0", "roaming"),
            os.path.join(self.localappdata, "Microsoft", "Office", "15.0", "roaming"),
            os.path.join(self.appdata, "Microsoft", "Office", "16.0", "roaming"),
            os.path.join(self.appdata, "Microsoft", "Office", "15.0", "roaming"),
        ]
        
        all_cache_paths = onedrive_cache_paths + office365_paths
        
        for cache_path in all_cache_paths:
            if not self.is_cleaning:
                break
            if os.path.exists(cache_path):
                try:
                    for root, dirs, files in os.walk(cache_path):
                        for file in files:
                            if any(keyword in file.lower() for keyword in ['recent', 'mru', 'cache', 'excel', '.json', '.xml']):
                                file_path = os.path.join(root, file)
                                try:
                                    os.remove(file_path)
                                    self.log_message(f"✓ Excel cloud cache temizlendi: {file}", "success")
                                except:
                                    pass
                except Exception as e:
                    continue
    
    def clean_excel_file_history(self):
        """Excel'in dosya geçmişini AppData'dan temizle"""
        try:
            # Excel'in yerel ayar dosyaları
            excel_roaming_path = os.path.join(self.appdata, "Microsoft", "Excel")
            excel_local_path = os.path.join(self.localappdata, "Microsoft", "Office")
            
            # Roaming Excel klasörü
            if os.path.exists(excel_roaming_path):
                for file in os.listdir(excel_roaming_path):
                    if file.endswith('.officeUI') or 'recent' in file.lower():
                        file_path = os.path.join(excel_roaming_path, file)
                        try:
                            os.remove(file_path)
                            self.log_message(f"✓ Excel config dosyası temizlendi: {file}", "success")
                        except:
                            pass
            
            # Local Excel cache
            if os.path.exists(excel_local_path):
                for root, dirs, files in os.walk(excel_local_path):
                    for file in files:
                        if any(keyword in file.lower() for keyword in ['recent', 'mru', 'cache']):
                            file_path = os.path.join(root, file)
                            try:
                                os.remove(file_path)
                                self.log_message(f"✓ Excel cache dosyası temizlendi: {file}", "success")
                            except:
                                pass
        except Exception as e:
            self.log_message(f"✗ Excel dosya geçmişi temizlenirken hata: {e}", "error")
    
    def delete_registry_key_recursive(self, hive, key_path):
        """Registry anahtarını alt anahtarlarıyla birlikte özyinelemeli sil"""
        try:
            with winreg.OpenKey(hive, key_path, 0, winreg.KEY_ALL_ACCESS) as key:
                # Alt anahtarları listele
                subkeys = []
                try:
                    i = 0
                    while True:
                        subkey = winreg.EnumKey(key, i)
                        subkeys.append(subkey)
                        i += 1
                except WindowsError:
                    pass
                
                # Alt anahtarları özyinelemeli sil
                for subkey in subkeys:
                    self.delete_registry_key_recursive(hive, f"{key_path}\\{subkey}")
                
            # Ana anahtarı sil
            winreg.DeleteKey(hive, key_path)
            
        except Exception as e:
            pass
    
    def clean_temp_files(self):
        """Geçici dosyaları temizle"""
        self.log_message("🗂️ Geçici dosyalar temizleniyor...", "info")
        
        temp_locations = [
            tempfile.gettempdir(),
            os.path.join(os.environ.get('WINDIR', ''), 'Temp'),
            os.path.join(self.localappdata, 'Temp'),
        ]
        
        for temp_dir in temp_locations:
            if not self.is_cleaning:
                break
            if os.path.exists(temp_dir):
                try:
                    for item in os.listdir(temp_dir):
                        if not self.is_cleaning:
                            break
                        item_path = os.path.join(temp_dir, item)
                        try:
                            if os.path.isdir(item_path):
                                shutil.rmtree(item_path)
                            else:
                                os.remove(item_path)
                        except:
                            continue
                    self.log_message(f"✓ Temp klasörü temizlendi: {temp_dir}", "success")
                except Exception as e:
                    self.log_message(f"✗ Temp klasörü temizlenirken hata: {e}", "error")
    
    def clean_browser_cache(self):
        """Tarayıcı önbelleğini temizle"""
        self.log_message("🌐 Tarayıcı önbelleği temizleniyor...", "info")
        
        # Chrome
        chrome_cache = os.path.join(self.localappdata, "Google", "Chrome", "User Data", "Default", "Cache")
        self.safe_delete(chrome_cache, "Chrome Cache")
        
        # Firefox
        firefox_profiles = os.path.join(self.appdata, "Mozilla", "Firefox", "Profiles")
        if os.path.exists(firefox_profiles):
            for profile in os.listdir(firefox_profiles):
                if not self.is_cleaning:
                    break
                cache_path = os.path.join(firefox_profiles, profile, "cache2")
                self.safe_delete(cache_path, "Firefox Cache")
        
        # Edge
        edge_cache = os.path.join(self.localappdata, "Microsoft", "Edge", "User Data", "Default", "Cache")
        self.safe_delete(edge_cache, "Edge Cache")
    
        for file in files:
                                try:
                                    file_path = os.path.join(root, file)
                                    os.remove(file_path)
                                    cleaned_items += 1
                                except:
                                    continue
                            
                            # Boş klasörleri sil
    
    
    def clean_windows_update_cache(self):
        """Windows Update önbelleğini temizle"""
        self.log_message("🔄 Windows Update önbelleği temizleniyor...", "info")
        
        update_cache = os.path.join(os.environ.get('WINDIR', ''), 'SoftwareDistribution', 'Download')
        self.safe_delete(update_cache, "Windows Update Cache")
    
    def clean_system_logs(self):
        """Sistem loglarını temizle"""
        self.log_message("📋 Sistem logları temizleniyor...", "info")
        
        log_locations = [
            os.path.join(os.environ.get('WINDIR', ''), 'Logs'),
            os.path.join(os.environ.get('WINDIR', ''), 'Temp'),
        ]
        
        for log_dir in log_locations:
            if not self.is_cleaning:
                break
            if os.path.exists(log_dir):
                try:
                    for file in glob.glob(os.path.join(log_dir, '*.log')):
                        if not self.is_cleaning:
                            break
                        self.safe_delete(file, "Log dosyası")
                except Exception as e:
                    self.log_message(f"✗ Log dosyaları temizlenirken hata: {e}", "error")
    
    def clean_prefetch(self):
        """Prefetch dosyalarını temizle"""
        self.log_message("⚡ Prefetch dosyaları temizleniyor...", "info")
        
        prefetch_dir = os.path.join(os.environ.get('WINDIR', ''), 'Prefetch')
        if os.path.exists(prefetch_dir):
            try:
                for file in os.listdir(prefetch_dir):
                    if not self.is_cleaning:
                        break
                    if file.endswith('.pf'):
                        file_path = os.path.join(prefetch_dir, file)
                        self.safe_delete(file_path, "Prefetch dosyası")
            except Exception as e:
                self.log_message(f"✗ Prefetch dosyaları temizlenirken hata: {e}", "error")
    
    def clean_windows_update_cache(self):
        """Windows Update önbelleğini temizle"""
        self.log_message("🔄 Windows Update önbelleği temizleniyor...", "info")
        update_cache = os.path.join(os.environ.get('WINDIR', ''), 'SoftwareDistribution', 'Download')
        self.safe_delete(update_cache, "Windows Update Cache")
    
    def clean_system_logs(self):
        """Sistem loglarını temizle"""
        self.log_message("📋 Sistem logları temizleniyor...", "info")
        
        log_locations = [
            os.path.join(os.environ.get('WINDIR', ''), 'Logs'),
            os.path.join(os.environ.get('WINDIR', ''), 'Temp'),
        ]
        
        for log_dir in log_locations:
            if not self.is_cleaning:
                break
            if os.path.exists(log_dir):
                try:
                    for file in glob.glob(os.path.join(log_dir, '*.log')):
                        if not self.is_cleaning:
                            break
                        self.safe_delete(file, "Log dosyası")
                except Exception as e:
                    self.log_message(f"✗ Log dosyaları temizlenirken hata: {e}", "error")
    
    def clean_prefetch(self):
        """Prefetch dosyalarını temizle"""
        self.log_message("⚡ Prefetch dosyaları temizleniyor...", "info")
        
        prefetch_dir = os.path.join(os.environ.get('WINDIR', ''), 'Prefetch')
        if os.path.exists(prefetch_dir):
            try:
                for file in os.listdir(prefetch_dir):
                    if not self.is_cleaning:
                        break
                    if file.endswith('.pf'):
                        file_path = os.path.join(prefetch_dir, file)
                        self.safe_delete(file_path, "Prefetch dosyası")
            except Exception as e:
                self.log_message(f"✗ Prefetch dosyaları temizlenirken hata: {e}", "error")
    
    def is_admin(self):
        """Yönetici yetkisi kontrolü"""
        try:
            return os.getuid() == 0
        except AttributeError:
            import ctypes
            return ctypes.windll.shell32.IsUserAnAdmin() != 0

def main():
    """Ana fonksiyon"""
    # Tkinter başlatma
    root = tk.Tk()
    
    # Uygulama başlatma
    app = ModernOfficeCleanerGUI(root)
    
    # Ana döngü
    root.mainloop()

if __name__ == "__main__":
    main()
