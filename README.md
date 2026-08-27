# FONET Kadın Doğum Konsültasyon Tarayıcı

FONET HBYS ameliyat listesindeki kayıtları ExtJS veri deposundan toplu okur ve HBYS'nin kendi konsültasyon servisini kullanarak kadın doğum konsültasyonlarını tarar.

## Özellikler

- Ekranda açık ameliyat listesini tek seferde toplar ve yalnızca kadın hastaları ayırır.
- Cinsiyeti ayrı alanın yanında `Yaş / Cinsiyet` gibi birleşik HBYS değerlerinden de tanır.
- Cinsiyet ameliyat tablosunda yoksa hasta/sevk servisinden; bu da mümkün değilse satırı seçerek üstteki gerçek `Yaş / Cinsiyet` alanından okur.
- Hasta geliş kimliklerini otomatik çıkarır.
- Konsültasyon servisini yalnızca kadın hastalar için kontrollü paralel sorgular.
- Sonuçta yalnızca Kadın Hastalıkları ve Doğum / Kadın Doğum / Jinekoloji konsültasyonu bulunan kadınları listeler.
- İstem nedeni, konsültasyon yanıtı, tarih, hasta ve ameliyat bilgilerini gösterir.
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
