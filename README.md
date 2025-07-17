
## Office History and System Cleaner

<img width="810" height="641" alt="image" src="https://github.com/user-attachments/assets/5687768c-193f-4dad-a541-93d1b225486d" />


A robust and user-friendly Python-based cleaner application designed for Windows systems to clear Office application history, temporary files, browser caches, and other system junk. This tool aims to enhance privacy and free up disk space with a modern graphical interface.

### Features

  * **Office Application History Cleaning:**
      * Clears recent file lists for Word, Excel, PowerPoint, and Access.
      * Removes Office 365 cloud cache and related registry entries.
      * Specific handling for Word and Excel temporary and cache files.
  * **System Junk Cleaning:**
      * Deletes temporary files from various system locations.
      * Clears browser caches for Chrome, Firefox, and Edge.
      * Empties the Recycle Bin (using multiple methods including PowerShell, CMD, and Python's `os.remove`/`shutil.rmtree`).
      * Cleans Windows Update cache.
      * Removes system log files.
      * Deletes Prefetch files.
  * **Modern GUI:**
      * Intuitive and easy-to-use interface built with Tkinter.
      * Real-time cleaning progress display.
      * Detailed logging of cleaning operations.
      * Quick selection buttons (Select All, Select None, Office Only).
      * Administrator privilege check.

### Screenshots

*(You can add screenshots here once your GUI is stable.)*
Example:

### Installation

1.  **Clone the repository:**
    ```bash
    git https://github.com/onder7/Office-History-and-System-Cleaner.git
    cd Office-System-Cleaner
    ```
2.  **Install dependencies:**
    ```bash
    pip install pyinstaller # Only needed if you plan to create an .exe
    ```
    *Note: Most required modules (`os`, `sys`, `shutil`, `winreg`, `tempfile`, `glob`, `subprocess`, `threading`, `time`, `pathlib`, `datetime`, `tkinter`) are part of Python's standard library and do not require separate installation.*

### Usage

1.  **Run the script directly:**

    ```bash
    python temizle.py
    ```

    The graphical interface will appear. Select the cleaning options you want and click "Start Cleaning".

2.  **Run the executable (if built):**
    If you've built the `.exe` file (see "Building an Executable" below), simply run `temizle.exe` from the `dist/` folder.

      * **Important:** For full functionality, especially for system-level cleaning like Recycle Bin, Windows Update Cache, and System Logs, it is recommended to **run the application as an administrator**.

### Building an Executable (EXE)

You can convert the Python script into a standalone executable (`.exe`) file for easier distribution without requiring a Python installation on the target machine. We recommend using `PyInstaller`.

1.  **Install PyInstaller:**
    ```bash
    pip install pyinstaller
    ```
2.  **Navigate to the script directory:**
    ```bash
    cd C:\python # Or wherever your temizle.py is located
    ```
3.  **Build the executable:**
    For a single, console-less executable with an icon (if you have `cleaner.ico` in the same directory):
    ```bash
    pyinstaller --onefile --noconsole --icon=cleaner.ico temizle.py
    ```
    If you don't have an icon, just omit `--icon=cleaner.ico`:
    ```bash
    pyinstaller --onefile --noconsole temizle.py
    ```
    The executable will be generated in the `dist/` folder.

### Contributing

Feel free to fork the repository, open issues, or submit pull requests.

### License

This project is licensed under the MIT License - see the [LICENSE](https://www.google.com/search?q=LICENSE) file for details.

-----

# Türkçe README.md

## Office Geçmiş ve Sistem Temizleyici

<img width="810" height="641" alt="image" src="https://github.com/user-attachments/assets/f1c0e4db-0f2f-4546-94b1-96428b6a238a" />


Windows sistemleri için tasarlanmış, Office uygulama geçmişini, geçici dosyaları, tarayıcı önbelleklerini ve diğer sistem çöplüklerini temizlemek için sağlam ve kullanıcı dostu Python tabanlı bir temizleyici uygulamasıdır. Bu araç, modern bir grafik arayüzle gizliliği artırmayı ve disk alanını boşaltmayı amaçlamaktadır.

### Özellikler

  * **Office Uygulama Geçmişi Temizliği:**
      * Word, Excel, PowerPoint ve Access için son kullanılanlar listelerini temizler.
      * Office 365 bulut önbelleğini ve ilgili kayıt defteri girişlerini kaldırır.
      * Word ve Excel geçici ve önbellek dosyaları için özel işlem yapar.
  * **Sistem Gereksiz Dosya Temizliği:**
      * Çeşitli sistem konumlarındaki geçici dosyaları siler.
      * Chrome, Firefox ve Edge tarayıcı önbelleklerini temizler.
      * Geri Dönüşüm Kutusunu boşaltır (PowerShell, CMD ve Python'ın `os.remove`/`shutil.rmtree` dahil birden çok yöntem kullanarak).
      * Windows Update önbelleğini temizler.
      * Sistem günlük dosyalarını kaldırır.
      * Prefetch dosyalarını siler.
  * **Modern Grafik Arayüz (GUI):**
      * Tkinter ile oluşturulmuş sezgisel ve kullanımı kolay arayüz.
      * Gerçek zamanlı temizlik ilerleme göstergesi.
      * Temizlik operasyonlarının detaylı günlük kaydı.
      * Hızlı seçim butonları (Tümünü Seç, Hiçbirini Seçme, Sadece Office).
      * Yönetici yetkisi kontrolü.

### Ekran Görüntüleri

*(GUI'niz sabit olduğunda buraya ekran görüntüleri ekleyebilirsiniz.)*
Örnek:

### Kurulum

1.  **Depoyu klonlayın:**
    ```bash
    https://github.com/onder7/Office-History-and-System-Cleaner.git
    cd Office-System-Cleaner
    ```
2.  **Bağımlılıkları yükleyin:**
    ```bash
    pip install pyinstaller # Sadece .exe oluşturmayı planlıyorsanız gereklidir
    ```
    *Not: Gerekli modüllerin çoğu (`os`, `sys`, `shutil`, `winreg`, `tempfile`, `glob`, `subprocess`, `threading`, `time`, `pathlib`, `datetime`, `tkinter`) Python'ın standart kütüphanesinin bir parçasıdır ve ayrı kurulum gerektirmez.*

### Kullanım

1.  **Betiği doğrudan çalıştırın:**

    ```bash
    python temizle.py
    ```

    Grafik arayüz açılacaktır. İstediğiniz temizlik seçeneklerini seçin ve "Temizliği Başlat" düğmesine tıklayın.

2.  **Çalıştırılabilir dosyayı (EXE) çalıştırın (oluşturulduysa):**
    Eğer `.exe` dosyasını oluşturduysanız (aşağıdaki "Çalıştırılabilir Dosya Oluşturma" bölümüne bakın), `dist/` klasöründen `temizle.exe` dosyasını çalıştırmanız yeterlidir.

      * **Önemli:** Geri Dönüşüm Kutusu, Windows Update Önbelleği ve Sistem Logları gibi sistem düzeyinde temizlikler için, uygulamanın **yönetici olarak çalıştırılması önerilir**.

### Çalıştırılabilir Dosya (EXE) Oluşturma

Python betiğini, hedef makinede Python kurulumu gerektirmeden daha kolay dağıtım için bağımsız bir çalıştırılabilir (`.exe`) dosyasına dönüştürebilirsiniz. `PyInstaller` kullanmanızı öneririz.

1.  **PyInstaller'ı yükleyin:**
    ```bash
    pip install pyinstaller
    ```
2.  **Betik dizinine gidin:**
    ```bash
    cd C:\python # Veya temizle.py dosyanızın bulunduğu dizin
    ```
3.  **Çalıştırılabilir dosyayı oluşturun:**
    Tek, konsol penceresiz ve ikonlu bir çalıştırılabilir dosya için (eğer `cleaner.ico` dosyanız aynı dizindeyse):
    ```bash
    pyinstaller --onefile --noconsole --icon=cleaner.ico temizle.py
    ```
    Eğer bir ikonunuz yoksa, sadece `--icon=cleaner.ico` kısmını çıkarın:
    ```bash
    pyinstaller --onefile --noconsole temizle.py
    ```
    Çalıştırılabilir dosya `dist/` klasöründe oluşturulacaktır.

### Katkıda Bulunma

Depoyu çatallamaktan (fork), sorunlar açmaktan (issues) veya çekme istekleri (pull requests) göndermekten çekinmeyin.

### Lisans

Bu proje MIT Lisansı altında lisanslanmıştır - ayrıntılar için [LICENSE](https://www.google.com/search?q=LICENSE) dosyasına bakın.

