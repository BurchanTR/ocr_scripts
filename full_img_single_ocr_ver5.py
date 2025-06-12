"""
Akış Şeması
GRUP 1- Hazırlık ve Yardımcı Fonksiyonlar: Yatırım Uygulaması, açıkmı, maximize mı, arka planda mı vs tüm kontrolleri yapan
uyaran veya açıksa öne getiren adımlar.
GRUP 2- Tümleşik Resim Oluşturma: Mouse göstergesini ilk sayfada doğru noktaya kaydırıp veri bölgesinden resim alan ardından
kaydırma yaparak tekrar resim alan, üst üste iki kez aynı resim gelmişse son sayfada olduğunu anlayıp döngüden çıkan bu
işlemler sırasında da toplanan tüm resimleri birleştirip tek bir resim yapan adımlar
GRUP 3- GRUP 2 ile oluşturulan tek resmi senet adı ve fiyat için iki ayrı kolon oluşturacak şekilde bölen ve bu kolonları tek
adımlı OCR ile satırları üzerinden okuyarak senet ismi ve fiyat verilerini daha yüksek ocr başarımı için ayrıştırıp daha
 sonra ocr işlemini uygulayan, ardından listede yer alan tekrarlı senet isimlerini fiyat en sonuncudan gelecek şekilde teke
indirgeyen adımlar.
GRUP 4- Excel'e Aktarma: Oluşan listeyi excel tablosuna aktaran adımlar.
"""
import os
import sys
import re
import pytesseract
import xlwings as xw
import pygetwindow as gw
import time
import pyautogui
import threading
import tkinter as tk
import imagehash
import pandas as pd
from PIL import ImageGrab
import numpy as np
from PIL import Image, ImageTk
import cv2
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor
from concurrent.futures import ProcessPoolExecutor as Executor

# ==== DEĞİŞKEN TANIMLARI VE ATANAN İLK DEĞERLER ====
bbox = bbox_1 = (46, 266, 260, 980)
bbox_2 = (46, 241, 260, 997) 
EXCEL_PATH = "SenetSepet-TAM11.xlsm"
SHEET_NAME = "OCR_list"
IMAGE_PATH = "merged_output.png"
APP_WINDOW_TITLE = "Borsa İşlem Platformu"
SCROLL_PIXELS = -120
DEBUG_MODE = False
DEBUG_STAGES = False
DEBUG_RESULTS = False
highlight_duration = 0.3 # seconds

# ==== KOORDİNAT VE SABİTLER ====
row_height = 42
NUM_ROWS = 17
symbol_x_end = 105
price_x_start = 106
capture_height = 740
start_y = 10
end_y = row_height * 17
scale_factor = 0.9
screen_x = 48       # Yatırım Platformu Uygulamasına göre ayarlı
screen_y = 260
screen_width = 212  # Senet adı ve fiyat kolonlarının toplam genişliği
same_image_flag = False
captured_images = []
merged_img = None

# Grup 1 - Hazırlık ve Yardımcı Fonksiyonlar

def connect_to_open_workbook(target_wb_name):
    # Excel uygulamaları içinde dolaş
    for candidate_app in xw.apps:
        for wb in candidate_app.books:
            if target_wb_name.lower() in wb.name.lower():
                return wb  # Workbook bulundu
    # Eğer buraya kadar geldiyse, workbook açık değil
    raise Exception(f"❌ Workbook '{target_wb_name}' açık değil.")

def bring_investing_app_to_front():
    windows = [w for w in gw.getWindowsWithTitle(APP_WINDOW_TITLE) if w.visible]
    if not windows:
        print("❌ Yatırım Platformu Uygulaması açık değil.")
        return False

    win = windows[0]
    win.activate()
    activate_scroll_area()
    ### time.sleep(0.1)

    # Minimize kontrolü ve restore edilmesi(sol üst koordinatlar -32000 civarındaysa minimize olmuş demektir)
    if win.left <= -32000 or win.top <= -32000:
        win.restore()
        ### time.sleep(0.1)

    # Ekran çözünürlüğünü al
    screen_w, screen_h = pyautogui.size()

    # Tam ekran kontrolü
    is_fullscreen = (win.left == 0 and win.top == 0 and
                     win.width == screen_w and win.height == screen_h)

    if not is_fullscreen:
        win.activate()
        time.sleep(0.1)
        pyautogui.press('f11')
        time.sleep(0.1)

    return True

def activate_scroll_area():
    # Mouse göstergesini ilgili listede konumlandırıyoruz
    pyautogui.moveTo((bbox[2] + 10, bbox[1] + 20))  # 10px sağa, üst taraftan 20px aşağı 
    # bbox iç kısımda dar bir aralık
    #pyautogui.click() # tıklarsak alım satım penceresi açar.
    ### time.sleep(0.1)

def scroll_down():
    activate_scroll_area()
    pyautogui.scroll(SCROLL_PIXELS)

def scroll_to_top_fast():
    activate_scroll_area()
    for _ in range(6):
        pyautogui.scroll(-SCROLL_PIXELS * 4)
        ### time.sleep(0.01)

def show_highlight_box(bbox, page_number, duration=None, margin=2, border_thickness=6, color="black"):
    x1, y1, x2, y2 = bbox
    width = x2 - x1 + margin * 2
    height = y2 - y1 + margin * 2
    x1 -= margin
    y1 -= margin

    def _box():
        root = tk.Tk()
        root.overrideredirect(True)
        # Increase height to accommodate text above the box
        root.geometry(f"{width}x{height + 30}+{x1}+{y1 - 30}")  # Move window up by 30 pixels
        root.attributes("-topmost", True)
        root.attributes("-transparentcolor", "white")  # Make white fully transparent

        canvas = tk.Canvas(root, width=width, height=height + 30, highlightthickness=0, bg="white")
        canvas.pack()

        # Draw black rectangle border (moved down by 30 pixels)
        canvas.create_rectangle(
            border_thickness // 2,
            border_thickness // 2,
            width - border_thickness // 2,
            height - border_thickness // 2,
            outline=color,
            width=border_thickness
        )

        # If duration is specified, destroy after that time
        if duration:
            root.after(int(duration * 1000), root.destroy)
        
        root.mainloop()  # Add this line to keep the window alive
        return root

    # thread'i yarat ve başlat
    thread = threading.Thread(target=_box)
    thread.daemon = True  # Ana program bittiğinde sonlanması için thread daemon True yapıldı
    thread.start()
    return thread  # Takip için thread'i döndür

def capture_screen(custom_bbox=None):
    """
    Ekran görüntüsü alır.
    Args:
        custom_bbox: Özel bbox değerleri. None ise varsayılan değerler kullanılır.
    Returns:
        numpy.ndarray: Ekran görüntüsü
    """
    if custom_bbox is None:
        bbox = (screen_x, screen_y, screen_x + screen_width, screen_y + capture_height)
    else:
        bbox = custom_bbox
    
    try:
        screen = ImageGrab.grab(bbox=bbox)
        img_array = np.array(screen)
        if img_array.size == 0:
            raise ValueError("Çekilen resim boş")
        return img_array
    except Exception as e:
        print(f"Resim çekme hatası: {e}")
        return None

def show_debug_stages(img_dict, title="OCR Stages Debug View", wait_ms=3000):
    """
    DEBUG MODE'da OCR işleme aşamalarının tümünü incelenmek üzerebir pencerede yanyana gösterir.

    Parameters:
        img_dict (dict): {"Aşama adı": görüntü} şeklinde OCR aşamalarını içerir.
        title (str): Görüntü penceresinin başlığı.
        wait_ms (int): ms cinsinden bekleme süresi. 0 sonsuz.
    """
    
    processed_imgs = []

    for name, img in img_dict.items():
        # Gri ya da tek kanal görselleri BGR'ye çevir
        if len(img.shape) == 2:
            img = cv2.cvtColor(img, cv2.COLOR_GRAY2BGR)
        
        # Aşama adını üstte gösteren çubuk ekle
        label_bar = np.full((25, img.shape[1], 3), 255, dtype=np.uint8)
        cv2.putText(label_bar, name, (10, 18), cv2.FONT_HERSHEY_SIMPLEX, 0.5, (0,0,0), 1)

        stacked = np.vstack([label_bar, img])
        processed_imgs.append(stacked)

    # Görselleri yatayda birleştir
    combined = cv2.hconcat(processed_imgs)

    # Göster
    cv2.imshow(title, combined)
    cv2.setWindowProperty(title, cv2.WND_PROP_TOPMOST, 1)
    cv2.moveWindow(title, 100, 100)

def extract_text(img, is_price_area=False):

    if img is None or img.size == 0:
        print("⚠️ Uyarı: OCR'a gönderilen görsel None veya boş.")
        return ""

    if not isinstance(img, (np.ndarray, Image.Image)):
        print(f"[ERROR] Unsupported image object type: {type(img)}")
        return ""
    
    # Eğer cv2 görüntüsü ise:
    if isinstance(img, np.ndarray):
        ocr_input_pil = Image.fromarray(img)
    else:
        ocr_input_pil = img  # Zaten PIL.Image.Image ise

    if is_price_area:
        # Fiyat alanı için özel OCR parametreleri
        config = "--psm 7 --oem 3 -c tessedit_char_whitelist=0123456789,."  # Sadece sayılar ve noktalama
    else:
        config = "--psm 7"  # Normal OCR parametreleri
    try:
        return pytesseract.image_to_string(ocr_input_pil, config=config).strip()
    except Exception as e:
        print(f"❌ OCR hatası: {e}")
        return ""
    
def preprocess(img, is_price_area=False, debug_view = True):
    # Boş değişken tanımları
    gray = clahe =binary = negative = dilated = img.copy()

    if is_price_area:
        try:
            # Fiyat alanı için özel ön işleme
            # 1. Gri tonlama
            gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
            
            # 2. Kontrast artırma
            clahe = cv2.createCLAHE(clipLimit=3.0, tileGridSize=(8,8))
            equalized = clahe.apply(gray) if clahe and gray is not None else gray
            
            # 3. Binary thresholding - griyi tamamen yok et
            _, binary = cv2.threshold(gray, 127, 255, cv2.THRESH_BINARY)
            
            # 4. Negatif görüntü oluştur
            negative = cv2.bitwise_not(binary)
            
            # 5. Morfolojik işlemler
            kernel = np.ones((1, 1), np.uint8)
            dilated = cv2.dilate(negative, kernel, iterations=1)                        
        except Exception as e:
            print(f"Görüntü işleme hatası: {e}")
            return img  # Hata durumunda orijinal görüntüyü döndür
        
        finally:
             if DEBUG_STAGES:
                  show_debug_stages({
                    "Original": img,
                    "Gray": gray,
                    "Clahe Equalized": equalized,
                    "Binary": binary,
                    "Negative": negative,
                    "Dilated": dilated
                   })
        return equalized
            
    else:
        # Sembol alanı için normal ön işleme
        clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8,8))
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        equalized = clahe.apply(gray)
        thresh = cv2.adaptiveThreshold(
            equalized, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
            cv2.THRESH_BINARY, 15, 5
        )
        kernel = np.ones((1, 1), np.uint8)
        dilated = cv2.dilate(thresh, kernel, iterations=1)

        if DEBUG_STAGES:
            show_debug_stages({
                "Original": img,
                "Gray": gray,
                "Clahe Equalized": equalized,
                "Threshold": thresh,
                "Dilated": dilated
                })

        return equalized

def process_single_row(i, full_img):
    y1 = start_y + i * row_height
    y2 = y1 + row_height
    row_img = full_img[y1:y2, :]

    # OCR işlemleri - fiyat alanı için özel işleme
    symbol_img = preprocess(row_img[:, :symbol_x_end], is_price_area=False, debug_view = False)
    price_img  = preprocess(row_img[:, price_x_start:], is_price_area=True, debug_view = False)
    
    # SADECE SON SAYFADA GÖSTER
    if False and same_image_flag:
        # Gri görüntüleri BGR formatına çevir
        symbol_img_tmp = symbol_img
        price_img_tmp = price_img

        if len(symbol_img.shape) == 2:
            symbol_img_tmp = cv2.cvtColor(symbol_img, cv2.COLOR_GRAY2BGR)

        if len(price_img.shape) == 2:
            price_img_tmp = cv2.cvtColor(price_img, cv2.COLOR_GRAY2BGR)

        # Yükseklik uyumsuzsa hizala
        h = row_img.shape[0]
        symbol_img_tmp = cv2.resize(symbol_img_tmp, (symbol_img_tmp.shape[1], h))
        price_img_tmp = cv2.resize(price_img_tmp, (price_img_tmp.shape[1], h))

        # combine işlemi
        try:
            # combined_img_tmp = np.hstack((row_img, symbol_img_tmp, price_img_tmp))

            cv2.imshow("row_img", row_img)
            cv2.waitKey(1)  # pencerenin oluşturulmasını bekle
            cv2.setWindowProperty("row_img", cv2.WND_PROP_TOPMOST, 1)
            cv2.moveWindow("row_img", 100, 100)
            print("row_img görüntüleniyor... Devam etmek için bir tuşa bas.")

            cv2.imshow("symbol_img_tmp", symbol_img_tmp)
            cv2.waitKey(1)  # pencerenin oluşturulmasını bekle
            cv2.setWindowProperty("symbol_img_tmp", cv2.WND_PROP_TOPMOST, 1)
            cv2.moveWindow("symbol_img_tmp", 350, 100)
            print("symbol_img_tmp görüntüleniyor... Devam etmek için bir tuşa bas.")

            cv2.imshow("price_img_tmp", price_img_tmp)
            cv2.waitKey(1)  # pencerenin oluşturulmasını bekle
            cv2.setWindowProperty("price_img_tmp", cv2.WND_PROP_TOPMOST, 1)
            cv2.moveWindow("price_img_tmp", 600, 100)
            print("price_img_tmp görüntüleniyor... Devam etmek için bir tuşa bas.")

            cv2.waitKey(0)
            cv2.destroyAllWindows()
        except Exception as e:
            print(f"❌ Görüntü birleştirme/gösterme hatası: {e}")

    symbol_text = extract_text(symbol_img)
    price_text  = extract_text(price_img, is_price_area=True)

    # Sembolleri Temizliyoruz
    cleaned_symbol = re.sub(r'[^A-Z0-9]', '', symbol_text.upper()).rstrip(':.•·*-')

    # Fiyatları Temizliyoruz
    try:
        temp_price = price_text.replace('.', '')  # binlik ayraç noktaları kaldır
        last_price = temp_price.replace(',', '.')
        cleaned_price = float(last_price)
        cleaned_price = f"{cleaned_price:.2f}"  # Keep 2 decimal places
    except ValueError:
        cleaned_price = ""  # Or set to "N/A"

    return (i, cleaned_symbol, cleaned_price)

#### YENI MANTIK ####
"""
def merge_vertical(images):
    #Verilen PIL imajlarının hepsini dikey olarak tek bir imajda birleştirir.
    total_height = sum(img.height for img in images)
    max_width = max(img.width for img in images)
    merged_image = Image.new('RGB', (max_width, total_height), color=(255, 255, 255))
    y_offset = 0
    for img in images:
        merged_image.paste(img, (0, y_offset))
        y_offset += img.height
    return merged_image
"""
def merge_vertical(images):
    """
    Verilen numpy array'lerinden oluşan görüntüleri dikey olarak birleştirir.
    """
    # Tüm numpy array'lerini RGB olarak PIL imaja çevir
    pil_images = [Image.fromarray(cv2.cvtColor(img, cv2.COLOR_BGR2RGB)) for img in images]

    total_height = sum(img.height for img in pil_images)
    max_width = max(img.width for img in pil_images)

    merged_image = Image.new('RGB', (max_width, total_height), color=(255, 255, 255))

    y_offset = 0
    for img in pil_images:
        merged_image.paste(img, (0, y_offset))
        y_offset += img.height

    return merged_image

# Örnek kullanım:
# captured_images = [img1, img2, img3] gibi bir listen olduğunu varsayalım
# merged_img = merge_vertical(captured_images)
# merged_img.show()  # Gözlemleme için

# Eğer kaydetmek istersen:
# merged_img.save("merged_output.png")

def remove_duplicates(results):
    """
    Removes duplicates by symbol, keeping the last occurrence in the list.
    Also logs filtered-out (duplicate) items.
    Input: List of (i, symbol, price)
    Output: Cleaned list sorted by original order of last appearance.
    """
    if DEBUG_MODE:
        print("\n--- [DEBUG] GİRİŞ LİSTESİ ---")
        for i, symbol, price in results:
            print(f"{i}: {symbol} - {price}")

    symbol_to_index = {}
    for idx, (i, symbol, price) in enumerate(results):
        symbol_key = symbol.strip().upper()
        symbol_to_index[symbol_key] = idx  # sadece en son görülenin indeksini tut

    # Benzersiz olanların indeksleri
    unique_indices = sorted(symbol_to_index.values())

    # Temizlenmiş (benzersiz) sonuçlar
    unique_results = [results[idx] for idx in unique_indices]

    if DEBUG_MODE:
        print("\n--- [DEBUG] DUPLICATE TEMİZLENMİŞ ÇIKIŞ LİSTESİ ---")
        for i, symbol, price in unique_results:
            print(f"{i}: {symbol} - {price}")

    # Elenen (tekrar olan) elemanları bul
    removed = []
    seen = set()
    for idx, (i, symbol, price) in enumerate(results):
        symbol_key = symbol.strip().upper()
        if symbol_key in seen and idx not in unique_indices:
            removed.append((i, symbol, price))
        else:
            seen.add(symbol_key)

    if DEBUG_MODE:
        if removed:
            print("\n--- [DEBUG] ELENEN (DUPLICATE) SATIRLAR ---")
            for i, symbol, price in removed:
                print(f"{i}: {symbol} - {price}")
        else:
            print("\n--- [DEBUG] ELENEN KAYIT YOK ---")

    return unique_results

def wait_until_image_changes(previous_img, max_wait=2.0, check_interval=0.05):
    """
    Ekran görüntüsü değişene kadar bekler. Maksimum bekleme süresi max_wait (saniye).
    previous_img: Bir önceki sayfanın sembol sütunu (numpy array)
    """
    start_time = time.time()
    while time.time() - start_time < max_wait:
        current_img = capture_screen(custom_bbox=bbox)[:, 0:symbol_x_end]  # sadece sembol sütunu
        current_hash = imagehash.average_hash(Image.fromarray(current_img))
        previous_hash = imagehash.average_hash(Image.fromarray(previous_img))

        if current_hash != previous_hash:
            return True  # Değişiklik algılandı
        time.sleep(check_interval)

    print("⚠️ Sayfa değişimi beklenirken zaman aşımı.")
    return False

def kara_kutu(page_number):
    print(f"\n📄 Sayfa {page_number} işleniyor...")
    initial_bbox = (screen_x, screen_y, screen_x + screen_width, screen_y + capture_height)
    current_box_thread = show_highlight_box(initial_bbox, page_number, duration=highlight_duration, color="black")
    time.sleep(highlight_duration)  # İlk kutunun görünmesini bekle

def kara_kutu_sil(current_box_thread):
    if current_box_thread and current_box_thread.is_alive():
        try:
            if tk._default_root:
                tk._default_root.quit()
                tk._default_root.destroy()
                tk._default_root = None
        except Exception as e:
            print(f"Pencere temizleme hatası (önemli değil): {e}")

# ==== MAIN PROCESS ====

def process_rows():
    global bbox
    global same_image_flag
    global merged_img
    hash_full_img = None
    hash_previous_img = None
    diff = 0
    # Scroll ve karşılaştırma döngüsü için değişkenler
    symbol_column_img = None
    previous_symbol_column_img = None
    page_number = 1
    current_box_thread = None  # kare kutunun thread ini takip etmekte gerekli.

    activate_scroll_area()
    scroll_to_top_fast()
    while same_image_flag == False:  #  same_image_flag True olduğunda döngü sonlanacak
        if DEBUG_MODE:
            kara_kutu(page_number)
            kara_kutu_sil(current_box_thread)
        
        # === CAPTURE SCREEN === Başarısız ise 5 kez deniyoruz ===
        retry_count = 0
        max_retries = 5  # Sınırsız da yapılabilir.
        while True:
            full_img = capture_screen(custom_bbox=bbox)
            if full_img is not None:
                break  # Çekilen resim boş değilse döngüden çık
            print("🔁 Resim alınamadı, tekrar deneniyor...")
            retry_count += 1
            if retry_count >= max_retries:
                print("❌ Maksimum deneme sayısına ulaşıldı. İşlem sonlandırılıyor.")
                break  # veya break / raise Exception, akışa göre
            time.sleep(1)

        captured_images.append(full_img)

        # AYNI SAYFA kontrolü için sadece sembollerin olduğu bölgeyi kırpıyoruz
        symbol_column_img = full_img[:, 0:symbol_x_end]  # x=0'den x=105'e kadar olan alan

        # Hash'i kırpılan bölgeden hesaplatıyoruz
        hash_full_img = imagehash.average_hash(Image.fromarray(symbol_column_img))
        
        if hash_previous_img is not None:
            diff = hash_previous_img - hash_full_img
            if diff == 0:  # Fark yok
                same_image_flag = True
            else:
                same_image_flag = False        

        # Son sayfa ise tümleşik imajı oluştur
        if same_image_flag:
            merged_img = merge_vertical(captured_images)
        else:
            previous_symbol_column_img = symbol_column_img.copy()
            hash_previous_img = imagehash.average_hash(Image.fromarray(previous_symbol_column_img))
        page_number += 1
        bbox = bbox_2
        scroll_down()
        time.sleep(0.05)
        #success = wait_until_image_changes(previous_symbol_column_img)
        #if not success:
        #    print("⚠️ Sayfa değişimi algılanamadı. Devam ediliyor ama dikkat gerekebilir.")

    scroll_to_top_fast()
    print(f"\n✅ İşlem tamamlandı! Toplam {page_number-1} sayfa tarandı.")

    # 🔁 Excel Workbook'a dönüyoruz
    excel_windows = [w for w in gw.getWindowsWithTitle("Excel") if w.visible]
    if excel_windows:
        excel_window = excel_windows[0]
        excel_window.activate()
        excel_window.maximize()
        print("✅ Focus returned to Excel.")
    else:
        print("❌ Excel window not found.")

    # 🔁 Varsa kalmış pencereleri temizliyoruz
    if current_box_thread and current_box_thread.is_alive():
        for window in tk._default_root.winfo_children() if tk._default_root else []:
            try:
                window.destroy()
            except:
                pass
        if tk._default_root:
            try:
                tk._default_root.destroy()
            except:
                pass

# ==== TÜMLEŞİK RESMİN OCR İŞLEMLERİ ====

#def load_and_preprocess_image_original(IMAGE_PATH):
def load_and_preprocess_image_original():
    img = merged_img
    gray = cv2.cvtColor(np.array(img), cv2.COLOR_RGB2GRAY)
    thresh = cv2.adaptiveThreshold(
        gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY_INV, 15, 9
    )
    return img, thresh

# def load_and_preprocess_image_actual(IMAGE_PATH, is_price_area=False, debug_view = True):
def load_and_preprocess_image_actual(is_price_area=False, debug_view = True):
    # Boş değişken tanımları
    img = np.array(merged_img)
    gray = clahe =binary = negative = dilated = img.copy()
    
    if is_price_area:
        try:
            # Fiyat alanı için özel ön işleme
            # 1. Gri tonlama
            gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
            
            # 2. Kontrast artırma
            clahe = cv2.createCLAHE(clipLimit=3.0, tileGridSize=(8,8))
            equalized = clahe.apply(gray) if clahe and gray is not None else gray
            
            # 3. Binary thresholding - griyi tamamen yok et
            ###_, binary = cv2.threshold(gray, 127, 255, cv2.THRESH_BINARY)
            
            # 4. Negatif görüntü oluştur
            ###negative = cv2.bitwise_not(binary)
            
            # 5. Morfolojik işlemler
            ###kernel = np.ones((1, 1), np.uint8)
            ###dilated = cv2.dilate(negative, kernel, iterations=1)                        
        except Exception as e:
            print(f"Görüntü işleme hatası: {e}")
            return img  # Hata durumunda orijinal görüntüyü döndür
        
        finally:
             if DEBUG_STAGES and debug_view:
                  """
                  show_debug_stages({
                    "Orig": img, "Gray": gray, "Clahe Equalized": equalized, "Bin": binary, "Neg": negative, "Dilated": dilated
                   })
                  """
                  show_debug_stages({
                    "Orig": img, "Gray": gray, "Clahe Equalized": equalized
                   })
        return img,equalized
            
    else:
        # Sembol alanı için normal ön işleme
        clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8,8))
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        equalized = clahe.apply(gray)
        thresh = cv2.adaptiveThreshold(
            equalized, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
            cv2.THRESH_BINARY, 15, 5
        )
        ###kernel = np.ones((1, 1), np.uint8)
        ###dilated = cv2.dilate(thresh, kernel, iterations=1)

        if DEBUG_STAGES and debug_view :
            """
            show_debug_stages({
                "Orig": img, "Gray": gray, "Clahe Equalized": equalized, "Bin": thresh, "Dilated": dilated
                })
            """
            show_debug_stages({
                "Orig": img, "Gray": gray, "Clahe Equalized": equalized, "Bin": thresh
                })

        return img, equalized
#def process_image(image_path):
def process_image():
    img, thresh = load_and_preprocess_image_original()
    img, equalized = load_and_preprocess_image_actual()

#==========  Daha önce complete_img_single_ocr.ver2.py olan dosya buraya ekleniyor ===========#

def split_columns(image, symbol_ratio=0.6):
    """Görseli sembol ve fiyat kolonu olarak ikiye ayırır."""
    height, width = image.shape[:2]
    split_x = int(width * symbol_ratio)
    symbol_col = image[:, :split_x]
    price_col = image[:, split_x:]
    ###cv2.imwrite("symbol_column.png", symbol_col)
    ###cv2.imwrite("price_column.png", price_col)
    return symbol_col, price_col

def ocr_column(image, psm=6, whitelist=None):
    """Tek bir kolon imajından satır satır OCR yapar."""
    config = f'--psm {psm}'
    if whitelist:
        config += f' -c tessedit_char_whitelist={whitelist}'
    text = pytesseract.image_to_string(image, config=config)
    lines = [line.strip() for line in text.split('\n') if line.strip()]
    return lines

### def process_combined_image_bulk(image_path):

def process_combined_image_bulk():
    """Toplu OCR: Kolonlara ayır, her kolonu OCR yap, eşleştir."""
    print("▶ Toplu OCR başlatılıyor...")

    # Görsel yükle ve CLAHE uygula (daha iyi sonuçlar için)
    # OpenCV ile işleyebilmek için Pillow formatından NumPy dizisine çevirme
    image = np.array(merged_img)  # NumPy formatına çevirme
    image = cv2.cvtColor(image, cv2.COLOR_RGB2BGR)  # RGB'yi BGR formatına çevirme

    lab = cv2.cvtColor(image, cv2.COLOR_BGR2LAB)
    l, a, b = cv2.split(lab)
    clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
    cl = clahe.apply(l)
    limg = cv2.merge((cl, a, b))
    image_clahe = cv2.cvtColor(limg, cv2.COLOR_LAB2BGR)

    # Sütunlara ayır
    symbol_col_img, price_col_img = split_columns(image_clahe)

    # Her sütuna OCR uygula
    
    symbol_img_pre = preprocess_column(symbol_col_img)
    price_img_pre = preprocess_column(price_col_img)
    symbols = ocr_column(symbol_img_pre)
    prices = ocr_column(price_img_pre, whitelist="0123456789.,-")
    
    # Eşleştir
    results = []
    for i in range(max(len(symbols), len(prices))):
        symbol = symbols[i] if i < len(symbols) else None
        price = prices[i] if i < len(prices) else None
        symbol = clean_symbol(symbol) if symbol else None
        price = clean_price(price) if price else None
        if symbol and price:
            results.append((symbol, price))
    return results

def preprocess_column(img):
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8,8))
    return clahe.apply(gray)

def clean_text(text):
    return re.sub(r'\s+', ' ', text.strip())

def clean_price(text):
    text = text.replace(' ', '').replace(',', '.')
    return re.sub(r'[^0-9.\-]', '', text)

def clean_symbol(text):
    return ''.join(c for c in text if c.isalnum()).upper()

"""
def clean_price(text):
    text = text.replace(',', '.')
    return ''.join(c for c in text if c.isdigit() or c == '.')
"""
def connect_to_open_workbook(target_wb_name):
    # Excel uygulamaları içinde dolaş
    for candidate_app in xw.apps:
        for wb in candidate_app.books:
            if target_wb_name.lower() in wb.name.lower():
                return wb  # Workbook bulundu
    # Eğer buraya kadar geldiyse, workbook açık değil
    raise Exception(f"❌ Workbook '{target_wb_name}' is not open in any Excel instance.")

# === Duplikeleri temizleme ===
def remove_duplicates(data):
    seen = set()
    cleaned = []
    for sym, prc in data:
        if sym not in seen and sym != "":
            cleaned.append((sym, prc))
            seen.add(sym)
    return cleaned

# === MATCH AND WRITE BRCH ===

def match_and_write_to_excel_with_xlwings_brch(wb,symbols, prices, excel_path=EXCEL_PATH, sheet_name=SHEET_NAME):
    if len(symbols) != len(prices):
        print(f"⚠️ UYARI: Sembol ({len(symbols)}) ve fiyat ({len(prices)}) sayısı eşleşmiyor!")
        raise ValueError("Sembol ve fiyat sayısı uyuşmuyor. İşlem durduruldu.")

    cleaned_data = []
    for symbol, price in zip(symbols, prices):
        clean_symbol = clean_text(symbol)
        clean_price_val = clean_price(price)
        cleaned_data.append((clean_symbol, clean_price_val))

    df = pd.DataFrame(cleaned_data, columns=['Symbol', 'Price'])

    sheet_names = [s.name.lower() for s in wb.sheets]

    if SHEET_NAME.lower() in sheet_names:
        ws = next(s for s in wb.sheets if s.name.lower() == SHEET_NAME.lower())
        ws.clear_contents()  # veya ws.clear() eğer stiller vs. de silinsin isteniyorsa
    else:
        ws = wb.sheets.add(name=SHEET_NAME)
    ws.range("A1").value = [["Hisse", "Son Fiyat"]]  # Add header

    ws.range("A2").value = cleaned_data  # tüm DataFrame'i tek seferde yaz
    # wb.save(excel_path)
    # wb.close()
    print(f"✔ Excel dosyası yazıldı (xlwings ile): {excel_path}")

# === merged_img yi tek seferde OCR layıp Excel'e yazan ana işlem ===
def process_bulk_image(wb):
    bulk_results = process_combined_image_bulk()
    bulk_results = remove_duplicates(bulk_results)
    symbols, prices = zip(*bulk_results) if bulk_results else ([], [])
    match_and_write_to_excel_with_xlwings_brch(wb, symbols, prices, EXCEL_PATH, SHEET_NAME)

# ==== RUN ====
if __name__ == "__main__":
    start_time = time.time()
    # Prereqs
    target_wb_name = os.path.basename(EXCEL_PATH)
    try:
        wb = connect_to_open_workbook(target_wb_name)
    except Exception as e:
        print(str(e))
        sys.exit(1)
    if not bring_investing_app_to_front():
        sys.exit(1)

    # Yatırım Platformu Uygulaması'ndan tüm sayfaları alıp tümleşik resim haline getiren fonksiyon
    process_rows()
    end_time = time.time()
    print('1.Süre:', end_time - start_time)
    start_time = time.time()
    process_image()
    end_time = time.time()
    print('2.Süre:', end_time - start_time)
    # merged_output.png den 2 kolondan çok satırlı TEK OCR yaparak alınan veriyi excel e yazan kısım
    start_time = time.time()
    process_bulk_image(wb)    
    end_time = time.time()
    print('3.Süre:', end_time - start_time)