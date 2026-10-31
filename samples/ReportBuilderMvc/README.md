# Spreadox Report Studio (MVC Sample)

Bu proje, **Spreadox** kütüphanesini kullanarak dinamik veri üretimi, interaktif grafik tipi yapılandırması, Excel ve PDF rapor oluşturma, üstbilgi/altbilgi (Header/Footer) özelleştirme ve özel font desteği sunan bir ASP.NET Core MVC örnek uygulamasıdır.

## Özellikler

1. **Dinamik Veri Üretimi (Random Data Generator)**
   - İstenilen sayıda veri başlığı üretme (örn: 10 başlık, 5 başlık, 15 başlık vb.).
   - Her başlık için otomatik kategoriler (Jan..Dec) ve gerçekçi değerler.
   - Üretilen verileri doğrudan arayüzden düzenleyebilme (Başlık adı, kategoriler, değerler, renk).

2. **Gelişmiş Grafik Seçenekleri (Chart Types)**
   - Her veri başlığı için bağımsız grafik tipi seçimi:
     - `ColumnClustered`, `ColumnStacked`
     - `BarClustered`, `BarStacked`
     - `Line`, `LineMarker`, `LineStacked`
     - `Area`, `AreaStacked`
     - `Pie`, `Doughnut`
     - `Scatter`, `Radar`, `Combo`
   - Her grafik serisi için özel HEX / ColorPicker renk seçimi.

3. **Sayfa Düzeni (Layout Options)**
   - **Ayrı Sayfalarda (Separate Pages / Sheets):** Her grafik kendi PDF sayfasında veya ayrı Excel çalışma sayfasında yer alır, otomatik özet istatistikler (Min, Max, Mean, Sum) eklenir.
   - **Aynı Sayfada (Single Page / Sheet):** Tüm grafikler tek bir sayfada sıralı olarak konumlandırılır.

4. **Header (Üstbilgi) Düzenleme Desteği**
   - Header Başlığı (Title)
   - Alt Açıklama Metni (Subtitle / Text)
   - Header Logosu / Resim Ekleme (Aspect-fit PNG, JPEG)
   - Yazı Tipi (Font Family) seçimi
   - Yazı Boyutu (Font Size) ve Kalın (Bold) seçeneği
   - Başlık Rengi seçimi (Color Picker)

5. **Footer (Altbilgi) Düzenleme Desteği**
   - İstenilen Özel Metin (Footer Text)
   - Sayfa Numarası Formatı (`{0}` geçerli sayfa, `{1}` toplam sayfa - örn: `Page {0} of {1}`)
   - Footer Logosu / Küçük Resim (PNG, JPEG)
   - Yazı Tipi (Font Family) seçimi

6. **Özel Font Yükleme Desteği (Custom TTF Font Registration)**
   - Kütüphaneye eklenen `PdfFontRegistry` ve `PdfFontResolver` sayesinde sisteme çalışma zamanında `.ttf` TrueType font dosyası yüklenebilir.
   - Yüklenen fontlar anında Header ve Footer yazı tipi listelerine eklenir ve PDF üretiminde kullanılır.

7. **Çift Çıktı Desteği (Excel & PDF)**
   - **PDF:** Spreadox'un yüksek performanslı vektör tabanlı A4 Landscape PDF motoru ile vektörel grafikler, kapak sayfası, header/footer ve özet istatistikler.
   - **Excel:** Yerel OOXML DrawingML grafikleri, hücre stilleri, kenarlıklar, renklendirme ve dondurulmuş bölmeler (freeze panes).

## Çalıştırma

Projeyi çalıştırmak için proje dizininde:

```bash
cd samples/ReportBuilderMvc
dotnet run
```

Tarayıcınızda açın:
```
http://localhost:5000 veya konsolda belirtilen port
```

