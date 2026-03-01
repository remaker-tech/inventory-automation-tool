# Inventory Handover Automation 🏢📄

Bu layihə ofis daxili inventarların təhvil-təslim prosesini tam avtomatlaşdırmaq üçün yazılmış bir Python skriptidir. Skript Excel bazasından məlumatları oxuyaraq hər bir əməkdaş üçün avtomatik olaraq Word (`.docx`) və PDF formatında Təhvil-Təslim aktları formalaşdırır.

## 🚀 Üstünlükləri
* **Kütləvi Emal:** Excel-dəki (A-dan F sütununa qədər) siyahını oxuyur və şəxslərə görə qruplaşdırır.
* **Avtomatik Doldurma:** Şablon `template2.docx` faylındakı cədvəli və məlumatları dinamik olaraq doldurur.
* **Dinamik Loqo və Başlıqlar:** Hər bir sənədə avtomatik loqo yerləşdirir və başlıqları tənzimləyir.
* **PDF Konvertasiyası:** Word faylları yaradıldıqdan sonra avtomatik olaraq PDF-ə çevrilir.
* **Fayl Adlandırılması:** Yaradılan fayllar təhkim olunan şəxslərin adı ilə (məs: `Ad_Soyad.pdf`) adlandırılır, bu da axtarışı çox asanlaşdırır.

## 🛠️ İstifadə Olunan Texnologiyalar
* **Python**
* `pandas` - Excel məlumatlarının idarə edilməsi üçün.
* `python-docx` - Word şablonlarının oxunması və dəyişdirilməsi üçün.
* `docx2pdf` - Word fayllarının PDF formatına çevrilməsi üçün.

## 📂 Necə İstifadə Etməli?

1. Layihəni kompüterinizə yükləyin.
2. Lazımi kitabxanaları quraşdırın:
   ```bash
   pip install -r requirements.txt

3. list.xlsx faylına məlumatları daxil edin.

4. baslat.bat faylına iki dəfə klikləyərək proqramı işə salın (və ya terminalda python tekko2.py yazın).

5. Hazır sənədlər avtomatik yaranan output qovluğunda yerləşəcək.

   Qeyd: docx2pdf kitabxanasının işləməsi üçün kompüterdə Microsoft Word quraşdırılmış olmalıdır.