# FONET Kadın Doğum Konsültasyon Tarayıcı

FONET HBYS ameliyat listesindeki kadın ve erkek tüm kayıtları ExtJS veri deposundan toplu okur ve HBYS'nin kendi konsültasyon servisini kullanarak kadın doğum konsültasyonlarını tarar.

## Özellikler

- Ekranda açık ameliyat listesini tek seferde toplar; kullanıcı `Tümü`, `Kadın` veya `Erkek` seçebilir.
- Hasta geliş kimliklerini otomatik çıkarır.
- Konsültasyon servisini kadın ve erkek tüm ameliyat kayıtları için kontrollü paralel sorgular.
- Sonuçta yalnızca Kadın Hastalıkları ve Doğum / Kadın Doğum / Jinekoloji / Obstetri / Perinatoloji konsültasyonu bulunan kayıtları listeler.
- İstem nedeni, konsültasyon yanıtı, tarih, hasta ve ameliyat bilgilerini gösterir.
- Bulunan sonuçların tamamını ameliyat tarihine göre yeniden eskiye gösterir; ekranda 100 kayıt sınırı uygulamaz.
- İstem veya yanıtta `gebe` geçen konsültasyonları sarı renkle işaretleyip listenin en üstünde ayrı gruplar; CSV'ye `Gebe İfadesi` sütunu ekler.
- Sonuç satırlarında TC kimlik numarası ve telefon gösterir; ameliyat kaydında eksikse sonucu bulunan hastanın HBYS hasta/sevk bilgisinden tamamlamayı dener.
- Duraklatma, devam etme, durdurma ve CSV dışa aktarma sunar.
- Verileri yalnızca tarayıcı belleğinde tutar; başka bir sunucuya göndermez.

## Kullanım

1. HBYS'de Ameliyat listesini açın ve sorguyu çalıştırın.
2. `fonet-kadin-dogum-tarayici.js` dosyasını, FONET Canlı Vizit'i çalıştırdığınız yöntemle HBYS sayfasının ana JavaScript bağlamında çalıştırın.
3. Açılan panelde **Listeyi Bul** düğmesine basın.
4. Bulunan kayıt sayısı doğruysa **Taramayı Başlat** düğmesine basın.
5. Tarama tamamlanınca **CSV İndir** ile sonucu alın.

## Güvenlik

Bu araç yalnızca `hbys.bursa.yerel` adresindeki mevcut oturumla çalışır. Hasta verileri GitHub'a veya başka bir dış sisteme gönderilmez. Üretilen CSV dosyasını kurum politikalarına uygun saklayın.
