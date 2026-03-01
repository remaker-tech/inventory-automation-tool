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

## 📂 İstifadə Qaydası
Bu proqramdan istifadə etmək üçün aşağıdakı addımları izləyin:

Yükləyin və ZIP-dən çıxarın: Sağ yuxarıdakı yaşıl Code düyməsindən "Download ZIP" seçərək layihəni endirin. Endirilən ZIP faylının içindəki bütün faylları yeni və boş bir qovluğa (məsələn: Tehvil_Teslim_Projesi) köçürün.

Vacib: Skriptin düzgün işləməsi üçün bütün fayllar mütləq eyni qovluqda olmalıdır.

Kitabxanaları quraşdırın: Terminalda (CMD) aşağıdakı əmri yazaraq lazımi kitabxanaları yükləyin (yalnız bir dəfə):

Bash (cmd) ----
[pip install -r requirements.txt]
Məlumatları daxil edin: list.xlsx faylını açın, inventar məlumatlarını daxil edin və faylı yadda saxlayıb bağlayın.

İşə salın: Qovluqdakı baslat.bat faylına iki dəfə klikləyin (və ya terminalda python tekko2.py yazın).

✅ Proses bitdikdə, hazırlanan bütün Word və PDF sənədləri avtomatik olaraq output qovluğunda görünəcək.

⚠️ Vacib Qeydlər
Texniki Tələblər: Proqramın işləməsi üçün kompüterinizdə Python 3.x və Microsoft Word proqramlarının quraşdırılması mütləqdir.


PDF Konvertasiyası: docx2pdf kitabxanası Word proqramının arxa planda açılmasını tələb edir, buna görə də konvertasiya zamanı Word proqramını tam bağlamamağınız tövsiyə olunur.
