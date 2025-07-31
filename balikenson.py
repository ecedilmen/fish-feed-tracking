from pypylon import pylon
import cv2 as cv
import numpy as np
import time
import os
import pythoncom
import win32com.client as win32
import math
import winsound  # Windows ses için

# Excel dosyası hazırlığı
dosya_adi = os.path.abspath("kirmizi_blob_kayitlari.xlsx")

pythoncom.CoInitialize()  # COM başlat (thread güvenliği için)

excel = win32.gencache.EnsureDispatch('Excel.Application')
excel.Visible = True

if os.path.exists(dosya_adi):
    wb = excel.Workbooks.Open(dosya_adi)
    ws = wb.Worksheets(1)
    satir_sayisi = ws.UsedRange.Rows.Count + 1
    if satir_sayisi < 2:
        satir_sayisi = 2
else:
    wb = excel.Workbooks.Add()
    ws = wb.Worksheets(1)
    ws.Cells(1, 1).Value = "Zaman"
    ws.Cells(1, 2).Value = "Balık No"
    ws.Cells(1, 3).Value = "X"
    ws.Cells(1, 4).Value = "Y"
    ws.Cells(1, 5).Value = "Genişlik"
    ws.Cells(1, 6).Value = "Yükseklik"
    ws.Cells(1, 7).Value = "Hareket Mesafesi (cm)"
    satir_sayisi = 2
    wb.SaveAs(dosya_adi)

# Basler kamera ayarları
tl_factory = pylon.TlFactory.GetInstance()
devices = tl_factory.EnumerateDevices()

if len(devices) == 0:
    print("Kamera bulunamadı!")
    exit()

camera = pylon.InstantCamera(tl_factory.CreateDevice(devices[0]))
camera.Open()
camera.StartGrabbing()

converter = pylon.ImageFormatConverter()
converter.OutputPixelFormat = pylon.PixelType_RGB8packed
converter.OutputBitAlignment = pylon.OutputBitAlignment_MsbAligned

# Kırmızı HSV aralıkları
lower_red1 = np.array([0, 50, 50])
upper_red1 = np.array([10, 255, 255])
lower_red2 = np.array([170, 50, 50])
upper_red2 = np.array([180, 255, 255])

# Mavi HSV aralıkları
lower_blue = np.array([100, 150, 0])
upper_blue = np.array([140, 255, 255])

son_kayit_zamani = time.time() - 1

# Kamera görüş alanı (cm)
genislik_cm = 32
yukseklik_cm = 25

onceki_merkezler = []  # Önceki frame blob merkezleri
hareketsiz_sure = {}   # Blob no -> sabit kalma süresi
hareketsiz_sinir = 3   # saniye cinsinden sınır

try:
    while camera.IsGrabbing():
        grab_result = camera.RetrieveResult(5000, pylon.TimeoutHandling_ThrowException)

        if grab_result.GrabSucceeded():
            image = converter.Convert(grab_result)
            frame = image.GetArray()
            frame = cv.cvtColor(frame, cv.COLOR_RGB2BGR)

            hsv = cv.cvtColor(frame, cv.COLOR_BGR2HSV)

            # Kırmızı maske
            mask1 = cv.inRange(hsv, lower_red1, upper_red1)
            mask2 = cv.inRange(hsv, lower_red2, upper_red2)
            mask_red = cv.bitwise_or(mask1, mask2)
            mask_red_blur = cv.GaussianBlur(mask_red, (9, 9), 2)

            # Mavi maske
            mask_blue = cv.inRange(hsv, lower_blue, upper_blue)
            mask_blue_blur = cv.GaussianBlur(mask_blue, (9, 9), 2)

            # Konturları bul
            contours_red, _ = cv.findContours(mask_red_blur, cv.RETR_EXTERNAL, cv.CHAIN_APPROX_SIMPLE)
            contours_blue, _ = cv.findContours(mask_blue_blur, cv.RETR_EXTERNAL, cv.CHAIN_APPROX_SIMPLE)

            blob_sayisi = 0
            simdi = time.time()
            mevcut_merkezler_red = []
            mevcut_merkezler_blue = []

            # Kırmızı bloblar
            for contour in contours_red:
                if cv.contourArea(contour) > 1000:
                    blob_sayisi += 1
                    x, y, w, h = cv.boundingRect(contour)
                    center_x = int(x + w / 2)
                    center_y = int(y + h / 2)
                    mevcut_merkezler_red.append((center_x, center_y, x, y, w, h))

            # Mavi bloblar
            for contour in contours_blue:
                if cv.contourArea(contour) > 1000:
                    x, y, w, h = cv.boundingRect(contour)
                    center_x = int(x + w / 2)
                    center_y = int(y + h / 2)
                    mevcut_merkezler_blue.append((center_x, center_y, x, y, w, h))

            hareket_mesafeleri = []
            for i, (cx, cy, x, y, w, h) in enumerate(mevcut_merkezler_red):
                if i < len(onceki_merkezler):
                    onceki_cx, onceki_cy = onceki_merkezler[i][0], onceki_merkezler[i][1]
                else:
                    onceki_cx, onceki_cy = cx, cy

                dx = cx - onceki_cx
                dy = cy - onceki_cy
                piksel_mesafe = math.sqrt(dx ** 2 + dy ** 2)

                frame_h, frame_w = frame.shape[:2]
                cm_per_px_x = genislik_cm / frame_w
                cm_per_px_y = yukseklik_cm / frame_h
                cm_per_px = (cm_per_px_x + cm_per_px_y) / 2

                hareket_cm = piksel_mesafe * cm_per_px
                hareket_mesafeleri.append(hareket_cm)

                # Hareket kontrolü - blob sabit mi?
                if hareket_cm < 0.5:  # 0.5 cm'den küçük hareket sabit kabul
                    if i not in hareketsiz_sure:
                        hareketsiz_sure[i] = 0
                    hareketsiz_sure[i] += (simdi - son_kayit_zamani)
                else:
                    hareketsiz_sure[i] = 0  # Hareket varsa sıfırla

                # Sabit kalan blob için uyarı ve ses
                if hareketsiz_sure.get(i, 0) >= hareketsiz_sinir:
                    uyarimesaji = f" {i+1}. baligi kaybettik!"
                    cv.putText(frame, uyarimesaji, (x, y - 30),
                               cv.FONT_HERSHEY_SIMPLEX, 0.7, (0, 0, 255), 2)
                    print(uyarimesaji)
                    winsound.Beep(1000, 700)

                # Çizimler kırmızı blob
                cv.rectangle(frame, (x, y), (x + w, y + h), (0, 0, 255), 2)
                cv.circle(frame, (cx, cy), 3, (0, 0, 255), -1)
                cv.putText(frame, f"Kirmizi Balik {i+1}", (x, y - 10),
                           cv.FONT_HERSHEY_SIMPLEX, 0.6, (0, 0, 255), 2)
                cv.putText(frame, f"Hareket: {hareket_cm:.2f} cm", (x, y + h + 20),
                           cv.FONT_HERSHEY_SIMPLEX, 0.5, (255, 0, 0), 1)

            # Çizimler mavi blob
            for i, (cx, cy, x, y, w, h) in enumerate(mevcut_merkezler_blue):
                cv.rectangle(frame, (x, y), (x + w, y + h), (255, 0, 0), 2)
                cv.circle(frame, (cx, cy), 3, (255, 0, 0), -1)
                cv.putText(frame, f"balik yemi {i+1}", (x, y - 10),
                           cv.FONT_HERSHEY_SIMPLEX, 0.6, (255, 0, 0), 2)

            # Kırmızı ile mavi bloblar kesişirse "TEHLİKE KAÇ" yaz
            for (rcx, rcy, rx, ry, rw, rh) in mevcut_merkezler_red:
                red_rect = (rx, ry, rx + rw, ry + rh)
                for (bcx, bcy, bx, by, bw, bh) in mevcut_merkezler_blue:
                    blue_rect = (bx, by, bx + bw, by + bh)

                    overlap_x = max(0, min(red_rect[2], blue_rect[2]) - max(red_rect[0], blue_rect[0]))
                    overlap_y = max(0, min(red_rect[3], blue_rect[3]) - max(red_rect[1], blue_rect[1]))
                    if overlap_x > 0 and overlap_y > 0:
                        cv.putText(frame, "TEHLİKE KAC", (bx, by - 30),
                                   cv.FONT_HERSHEY_SIMPLEX, 1, (0, 0, 255), 3)

            if simdi - son_kayit_zamani >= 1:
                zaman = time.strftime("%H:%M:%S")
                for i, (cx, cy, x, y, w, h) in enumerate(mevcut_merkezler_red):
                    hareket = hareket_mesafeleri[i] if i < len(hareket_mesafeleri) else 0
                    try:
                        ws.Cells(satir_sayisi, 1).Value = zaman
                        ws.Cells(satir_sayisi, 2).Value = i + 1  # Blob no
                        ws.Cells(satir_sayisi, 3).Value = x
                        ws.Cells(satir_sayisi, 4).Value = y
                        ws.Cells(satir_sayisi, 5).Value = w
                        ws.Cells(satir_sayisi, 6).Value = h
                        ws.Cells(satir_sayisi, 7).Value = round(hareket, 2)
                        satir_sayisi += 1
                    except Exception as e:
                        print("Excel yazma hatası:", e)

                try:
                    wb.Save()
                except Exception as e:
                    print("Excel dosyası kaydedilemedi:", e)

                son_kayit_zamani = simdi

            onceki_merkezler = mevcut_merkezler_red.copy()

            cv.putText(frame, f"Toplam Kirmizi Balik: {blob_sayisi}", (10, 60),
                       cv.FONT_HERSHEY_TRIPLEX, 0.6, (0, 0, 0), 2)
            cv.namedWindow("kamera",cv.WINDOW_NORMAL)
            cv.imshow("kamera", frame)
            # cv.imshow("maske_kirmizi", mask_red)
            # cv.imshow("maske_mavi", mask_blue)

            key = cv.waitKey(1) & 0xFF
            if key == ord('q'):
                print("Çıkılıyor...")
                break

        grab_result.Release()

finally:
    camera.StopGrabbing()
    camera.Close()
    wb.Save()
    wb.Close()
    excel.Quit()
    cv.destroyAllWindows()
