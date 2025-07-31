# fish-feed-tracking
Akvaryum İçindeki Nesnelerin Gerçek Zamanlı Tespiti
# Aquarium Object Tracking

## Proje Özeti  
Kamera aracılığıyla akvaryum içindeki kırmızı (balıklar) ve mavi (yemler) nesneler gerçek zamanlı olarak tespit edilmektedir. HSV renk uzayında belirlenen aralıklarla maskeleme yapılarak nesneler ayrıştırılmıştır. Nesnelerin konumları, boyutları ve hareket mesafeleri hesaplanıp, hareket etmeyen nesneler için sesli ve görsel uyarı sistemi uygulanmıştır. Tespit edilen veriler saniyelik olarak Excel dosyasına kaydedilmektedir.

## Teknik Detaylar  
- Kırmızı ve mavi nesneler için HSV renk aralıkları ile maskeleme  
- Kontur analizi ile nesne koordinatları ve boyutlarının çıkarılması  
- Nesnelerin hareket mesafelerinin gerçek dünya birimi (cm) olarak hesaplanması  
- Hareket etmeyen nesneler için uyarı (görsel ve sesli)  
- Nesnelerin kesişiminde “TEHLİKE KAÇ” uyarısı  
- Verilerin gerçek zamanlı Excel kaydı (zaman, nesne numarası, koordinatlar, boyutlar, hareket mesafesi)  
- OpenCV ile nesnelerin görsel olarak işaretlenmesi ve etiketlenmesi  

## Kullanılan Teknolojiler  
- Python  
- Basler Kamera ve pypylon kütüphanesi  
- OpenCV  
- Numpy  
- Win32com (Excel kaydı için)  
- Winsound (sesli uyarı için)  

## Kurulum  
```bash
pip install pypylon opencv-python numpy pywin32
