#### Bu dosyayı git ile takibe aldım.
#### https://github.com/baba-enkai/ocr_baba_enkai_ver4
#### 26.05.2025
import cv2
import pytesseract
import imagehash
import numpy as np
from PIL import ImageGrab
import xlwings as xw
import win32gui
import win32con
import win32process
import sys
import time
import re
import os
import pyautogui
import tkinter as tk
import threading
import pygetwindow as gw
from datetime import datetime
from PIL import Image, ImageTk
import concurrent.futures
from concurrent.futures import ThreadPoolExecutor

# ==== CONFIG ====
bbox_1 = (46, 266, 260, 980)
bbox_2 = (46, 241, 260, 997) 
bbox = bbox_1 
APP_WINDOW_TITLE = "Borsa İşlem Platformu"  # Replace this with your actual app window title
DEBUG_MODE = True  # Set to False to disable rectangle drawing
DEBUG_STAGES = False
DEBUG_RESULTS = False
EXCEL_PATH = "SenetSepet-TAM11.xlsm"  # Your Excel file
SHEET_NAME = "OCR_list"
SCROLL_PIXELS = -120
OCR_CONFIG = "--psm 7 -c preserve_interword_spaces=1"
highlight_duration = 10 # seconds

# Coordinates and constants
row_height = 42
NUM_ROWS = 17
symbol_x_end = 105
price_x_start = 106
capture_height = 740
start_y = 10
end_y = row_height * 17
scale_factor = 0.9
screen_x = 48       # Adjust this for your app window
screen_y = 260
screen_width = 212  # Adjust this based on the width needed for prices
same_image_flag = False
captured_images = []

#### YENI MANTIK ####
def merge_vertical(images):
    """
    Verilen PIL imajlarının hepsini dikey olarak tek bir imajda birleştirir.
    """
    total_height = sum(img.height for img in images)
    max_width = max(img.width for img in images)
    
    merged_image = Image.new('RGB', (max_width, total_height), color=(255, 255, 255))
    
    y_offset = 0
    for img in images:
        merged_image.paste(img, (0, y_offset))
        y_offset += img.height
    
    return merged_image

# Örnek kullanım:
# captured_images = [img1, img2, img3] gibi bir listen olduğunu varsayalım
# merged_img = merge_vertical(captured_images)
# merged_img.show()  # Gözlemleme için

# Eğer kaydetmek istersen:
# merged_img.save("merged_output.png")

def show_time_window(title, time_str, x=100, y=100):
    def _show():
        root = tk.Tk()
        root.overrideredirect(True)
        root.geometry(f"180x50+{x}+{y}")
        root.attributes("-topmost", True)
        label = tk.Label(root, text=f"{title}:\n{time_str}", font=("Arial", 10), bg="black", fg="lime")
        label.pack(expand=True, fill='both')
        root.after(10000, root.destroy)  # Auto close after 4 seconds
        root.mainloop()
    threading.Thread(target=_show).start()

def mark_start(label="Start Time", x=540, y=100):
    start_time = datetime.now().strftime("%H:%M:%S.%f")[:-3]
    show_time_window(label, start_time, x=x, y=y)
    return datetime.now()

def mark_end(label="End Time", x=760, y=100):
    end_time = datetime.now().strftime("%H:%M:%S.%f")[:-3]
    show_time_window(label, end_time, x=x, y=y)
    return datetime.now()

# Debug için OCR alanını görselleştirmek için gerekli fonksiyon.
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

    # Create and start the thread
    thread = threading.Thread(target=_box)
    thread.daemon = True  # Make thread daemon so it exits when main program exits
    thread.start()
    return thread  # Return the thread for tracking

def activate_scroll_area():
    # Moves and clicks inside the list area to ensure it is scrollable.
    pyautogui.moveTo((bbox[2] + 10, bbox[1] + 20))  # 10px to the right, 20px below top
    # Slight offset inside the bbox
    #pyautogui.click()
    time.sleep(0.1)

def scroll_down():
    activate_scroll_area()
    pyautogui.scroll(SCROLL_PIXELS)

def scroll_to_top_fast():
    activate_scroll_area()
    for _ in range(6):
        pyautogui.scroll(-SCROLL_PIXELS * 4)
        time.sleep(0.01)

def bring_investing_app_to_front():
    windows = [w for w in gw.getWindowsWithTitle(APP_WINDOW_TITLE) if w.visible]
    if not windows:
        print("❌ Could not find the Investing app window.")
        return False

    win = windows[0]
    win.activate()
    activate_scroll_area()
    time.sleep(0.1)

    # Minimize kontrolü (sol üst koordinatlar -32000 civarındaysa minimize olmuş demektir)
    if win.left <= -32000 or win.top <= -32000:
        print("🔄 Uygulama minimize edilmiş, geri getiriliyor...")
        win.restore()
        time.sleep(0.1)

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

# ==== IMAGE PROCESSING ====
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
             if DEBUG_STAGES and debug_view:
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

        if DEBUG_STAGES and debug_view :
            show_debug_stages({
                "Original": img,
                "Gray": gray,
                "Clahe Equalized": equalized,
                "Threshold": thresh,
                "Dilated": dilated
                })

        return equalized
    
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

# ==== SCREEN CAPTURE ====
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
            raise ValueError("Captured image is empty")
        return img_array
    except Exception as e:
        print(f"Error capturing screen: {e}")
        return None

def draw_debug_rectangles_with_text(img, texts, row_count=17):
    img_with_rects = img.copy()
    text_height = 25
    font = cv2.FONT_HERSHEY_SIMPLEX
    font_scale = 0.5
    thickness = 1

    # Draw rectangles on the main image
    for i in range(row_count):
        y1 = start_y + i * row_height
        y2 = y1 + row_height
        cv2.rectangle(img_with_rects, (0, y1), (symbol_x_end, y2), (0, 0, 0), 1)
        cv2.rectangle(img_with_rects, (price_x_start, y1), (screen_width, y2), (0, 0, 0), 1)

    # Create a white background for text display
    text_img = 255 * np.ones((img.shape[0], 220, 3), dtype=np.uint8)

    for i, (symbol, price) in enumerate(texts):
        y = start_y + i * row_height + 20
        text_line = f"{i+1:02d}: {symbol} - {price}"
        cv2.putText(text_img, text_line, (10, y), font, font_scale, (0, 0, 0), thickness)

    # Combine both images side by side
    combined = np.hstack((img_with_rects, text_img))

    cv2.imshow("Debug View: Image + OCR Text", combined)
    cv2.setWindowProperty("Debug View: Image + OCR Text", cv2.WND_PROP_TOPMOST, 1)
    cv2.moveWindow("Debug View: Image + OCR Text", 1300, 100)
    key = cv2.waitKey(1000)  # 1 seconds or key
    cv2.destroyAllWindows()

def show_debug_stages(img_dict, title="OCR Stages Debug View", wait_ms=3000):
    """
    OCR işleme aşamalarını tek bir yatay görselde gösterir.

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
    #cv2.waitKey(wait_ms)
    #cv2.destroyAllWindows()

def connect_to_open_workbook(target_wb_name):
    # Excel uygulamaları içinde dolaş
    for candidate_app in xw.apps:
        for wb in candidate_app.books:
            if target_wb_name.lower() in wb.name.lower():
                return wb  # Workbook bulundu
    # Eğer buraya kadar geldiyse, workbook açık değil
    raise Exception(f"❌ Workbook '{target_wb_name}' is not open in any Excel instance.")

# ==== MAIN PROCESS ====

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

def compare_symbol_lists(list1, list2):
    """
    İki sembol listesini karşılaştırır. Fiyatları dikkate almaz.
    Sadece sembollerin aynı olup olmadığını kontrol eder.
    """
    # Sadece sembolleri al (fiyatları çıkar)
    symbols1 = {item[1] for item in list1}  # item[1] sembol, item[2] fiyat
    symbols2 = {item[1] for item in list2}
    
    return symbols1 == symbols2

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

def adjust_initial_bbox_for_debug_mode(bbox, debug_mode, border_thickness=6, margin=2):
    if not debug_mode:
        #top_margin = 30 + (border_thickness // 2) + margin
        #bottom_margin = (border_thickness // 2) + margin
        #left_margin = right_margin = (border_thickness // 2) + margin
        top_margin = bottom_margin = 5
        left_margin = right_margin = 2
        x1, y1, x2, y2 = bbox
        if DEBUG_MODE:
            print('x1, y1, x2, y2 ==== ', x1, y1, x2, y2)
            print('left_margin, right_margin, top_margin, bottom_margin ==== ', left_margin, right_margin, top_margin, bottom_margin)
        return (
            x1 + left_margin,
            y1 + top_margin,
            x2 - right_margin,
            y2 - bottom_margin
        )
    else:
        return bbox

def process_rows():
    global bbox
    global same_image_flag
    hash_full_img = None
    hash_previous_img = None
    diff = 0
    target_wb_name = os.path.basename(EXCEL_PATH)
    try:
        wb = connect_to_open_workbook(target_wb_name)
    except Exception as e:
        print(str(e))
        sys.exit(1)

    sheet_names = [s.name.lower() for s in wb.sheets]

    if SHEET_NAME.lower() in sheet_names:
        ws = next(s for s in wb.sheets if s.name.lower() == SHEET_NAME.lower())
    else:
        ws = wb.sheets.add(name=SHEET_NAME)

    if False:
        ws.clear_contents()  # veya ws.clear() eğer stiller vs. de silinsin isteniyorsa
        ws.range("A1").value = [["Hisse", "Son Fiyat"]]  # Add header

    if not bring_investing_app_to_front():
        print("⚠️ Please open the investing app in Firefox private window and re-run the script.")
        sys.exit(1)

    activate_scroll_area()
    scroll_to_top_fast()

    # Scroll ve karşılaştırma döngüsü için değişkenler
    global_idx = 0
    previous_results = []
    symbol_column_img = None
    previous_symbol_column_img = None
    page_number = 1
    all_results = []  # Tüm sayfaların sonuçlarını tutacak liste
    current_box_thread = None  # To keep track of the current highlight box thread
    bbox = adjust_initial_bbox_for_debug_mode(bbox, DEBUG_MODE, border_thickness=6, margin=2)
    print ('bbox ==== ', bbox)

    

    while same_image_flag == False:  #  same_image_flag True olduğunda döngü sonlanacak
        if DEBUG_MODE:
            print(f"\n📄 Sayfa {page_number} işleniyor...")

            # === KARA KUTU ===
            initial_bbox = (screen_x, screen_y, screen_x + screen_width, screen_y + capture_height)
            current_box_thread = show_highlight_box(initial_bbox, page_number, duration=highlight_duration, color="black")
            time.sleep(highlight_duration)  # İlk kutunun görünmesini bekle

        # İlk senet isminin pozisyonunu bulurken geçici capture yapılıyor.
        # 111 first_symbol_y = find_first_symbol_position()
        # 112 if DEBUG_MODE:
        # 113    print(f"\n🔍 Hizalama Detayları:")
        # 114    print(f"  • screen_y (başlangıç): {screen_y}")
        # 115    print(f"  • first_symbol_y (bulunan): {first_symbol_y}")
            
        # 116 adjusted_screen_y = first_symbol_y - 5  # Küçük bir margin ekle
        
        # 117 if DEBUG_MODE:
        # 118    print(f"  • adjusted_screen_y (hesaplanan): {adjusted_screen_y}")
        # 119    print(f"  • Offset detayları:")
        # 120    print(f"    - first_symbol_y: {first_symbol_y}")
        # 121    print(f"    - -5 (margin): -5")
        
        # 122 adjusted_bbox = (screen_x, adjusted_screen_y, screen_x + screen_width, adjusted_screen_y + capture_height)
        # 123 print(f"  • Capture bölgesi (x1, y1, x2, y2): {adjusted_bbox}")

        if DEBUG_MODE:
            # KARA KUTU'yu sil
            if current_box_thread and current_box_thread.is_alive():
                try:
                    if tk._default_root:
                        tk._default_root.quit()
                        tk._default_root.destroy()
                        tk._default_root = None
                except Exception as e:
                    print(f"Pencere temizleme hatası (önemli değil): {e}")
        
        # 134 === YEŞİL KUTU ===
        # 135 current_box_thread = show_highlight_box(adjusted_bbox, page_number, duration=highlight_duration, color="green")
        # 136 time.sleep(highlight_duration)  # Hizalama kutunun görünmesini bekle
            
        # 137 === YEŞİL KUTU TEMİZLEME ===
        # 138 try:
        # 139     if current_box_thread and current_box_thread.is_alive():
        # 140         current_box_thread = None  # Thread referansını temizle
        # 141         if tk._default_root:
        # 142             tk._default_root.quit()  # Önce quit çağır
        # 143             tk._default_root.destroy()  # Sonra destroy
        # 144             tk._default_root = None  # Root referansını temizle
        # 145             except Exception as e:
        # 146                 print(f"Temizleme hatası (önemli değil): {e}")
        # 147             time.sleep(0.2)  # Biraz daha uzun bekle
        
        # === GERÇEK CAPTURE SCREEN başarısız ise 5 kez tekrarlıyoruz ===
        retry_count = 0
        max_retries = 5  # Sınırsız da yapılabilir.
        while True:
            full_img = capture_screen(custom_bbox=bbox)
            if full_img is not None:
                break  # Capture başarılı, döngüden çık
            print("🔁 Ekran görüntüsü alınamadı, tekrar deneniyor...")
            retry_count += 1
            if retry_count >= max_retries:
                print("❌ Maksimum deneme sayısına ulaşıldı. İşlem sonlandırılıyor.")
                break  # veya break / raise Exception, senin akışına göre
            time.sleep(1)

        captured_images.append(Image.fromarray(cv2.cvtColor(full_img, cv2.COLOR_BGR2RGB)))

        # Sadece sembollerin olduğu bölgeyi kırpıyoruz
        symbol_column_img = full_img[:, 0:symbol_x_end]  # x=0'den x=105'e kadar olan alan

        # Görüntüyü dosyaya kaydet
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"debug_capture_{timestamp}.png"
        cv2.imwrite(filename, full_img)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"debug_capture_symbol_column_{timestamp}.png"
        cv2.imwrite(filename, symbol_column_img)
        print(f"📸 Ekran görüntüleri kaydedildi: {filename}")

        # Hash bu bölgeden alınsın
        hash_full_img = imagehash.average_hash(Image.fromarray(symbol_column_img))
        # eski ... hash_full_img = imagehash.average_hash(Image.fromarray(full_img))
        if hash_previous_img is None:
            print('İlk sayfa...Karşılaştırma yapılmayacak.')
        else:
            print('hash_previous_img = ', hash_previous_img)
            print('hash_full_img = ', hash_full_img)
            diff = hash_previous_img - hash_full_img
            print(f"Fark: {diff}")
            if diff == 0:  # Fark yok
                same_image_flag = True
                print("🟡Aynı sayfa geldi. Döngü sonlandırılacak.")
            else:
                same_image_flag = False        
                print('Yeni sayfa geldi.Devam ediliyor...')

        # Capture sonrası kontrol
        if full_img is None:
            print("  • Capture başarısız!")
        else:
            print(f"  • Capture başarılı - Görüntü boyutu: {full_img.shape}")

            # SADECE SON SAYFADA GÖSTER
            if False and same_image_flag:
                cv2.imshow("📸 FINAL PAGE - CAPTURED IMAGE", full_img)
                cv2.setWindowProperty("📸 FINAL PAGE - CAPTURED IMAGE", cv2.WND_PROP_TOPMOST, 1)
                cv2.moveWindow("📸 FINAL PAGE - CAPTURED IMAGE", 1300, 100)
                print("👁️ Son sayfa görüntüleniyor... Devam etmek için bir tuşa bas.")
                cv2.waitKey(0)
                cv2.destroyAllWindows()
                _ = preprocess(full_img, debug_view=True)

        results = []
        with ThreadPoolExecutor(max_workers=6) as executor:
            futures = [executor.submit(process_single_row, i, full_img) for i in range(NUM_ROWS)]
            for future in futures:
                try:
                    i, symbol, price = future.result()
                    results.append((global_idx, symbol, price))
                    global_idx += 1
                except Exception as e:
                    print(f"Error processing row: {e}")
                    continue

        if not results:
            print("No results obtained, retrying...")
            time.sleep(1)
            continue

        # Sonuçları sırala
        results.sort(key=lambda x: x[0])
        if DEBUG_RESULTS:
            print("R"*100)
            print("results: ", results)
            print("X"*100)

        # Debug modunda göster - OCR işlemi bittikten sonra
        if False and DEBUG_MODE:
            texts = [(s, p) for _, s, p in results]
            draw_debug_rectangles_with_text(full_img, texts)
            cv2.waitKey(0)                                   # Tuş basılana kadar bekle
            cv2.destroyAllWindows()                          # Sonra tüm pencereleri kapat

        # Önceki liste ile karşılaştır. Bu eski mantık kaldırıldı. Artık imaj karşılaştırma yapıyoruz.

        if DEBUG_MODE and same_image_flag:
            print("P"*100)
            print("Previous Results: ", previous_results)
            print("R"*100)
            print("Results: ", results)
            print("X"*100)

        # Sonuçları ana listeye ekle
        all_results.extend(results)

        # Son sayfa ise tekrarlıları kaldır
        if same_image_flag:
            merged_img = merge_vertical(captured_images)
            merged_img.show()
            merged_img.save("merged_output.png")
            all_results = remove_duplicates(all_results)
            # Excel' e yaz
            start_time = time.time()
            if False:
                start_row = 2
                for i, (_, symbol_text, price_text) in enumerate(all_results):
                    ws.range(f"A{start_row + i}").value = symbol_text
                    ws.range(f"B{start_row + i}").value = price_text
            end_time = time.time()
            print(f"İşlem süresi: {end_time - start_time} saniye")
        
        else:
            # Mevcut sonuçları önceki sonuç olarak kaydet
            previous_results = results.copy()
            previous_symbol_column_img = symbol_column_img.copy()
            hash_previous_img = imagehash.average_hash(Image.fromarray(previous_symbol_column_img))
        page_number += 1
        bbox = bbox_2
        # Scroll yap
        scroll_down()
        time.sleep(1)  # Scroll sonrası bekleme

    scroll_to_top_fast()
    print(f"\n✅ İşlem tamamlandı! Toplam {page_number-1} sayfa tarandı.")
    print(f"📊 Toplam {len(all_results)} adet senet/fiyat çifti bulundu.")

    # 🔁 Return focus to Excel
    excel_windows = [w for w in gw.getWindowsWithTitle("Excel") if w.visible]
    if excel_windows:
        excel_window = excel_windows[0]
        excel_window.activate()
        excel_window.maximize()
        print("✅ Focus returned to Excel.")
    else:
        print("❌ Excel window not found.")

    # Clean up any remaining highlight boxes
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

def find_first_symbol_position():
    """
    İlk senet isminin pozisyonunu tespit eder.
    Returns:
        int: İlk senet isminin y koordinatı
    """

    print("\n" + "="*50)
    print("🔍 FIND_FIRST_SYMBOL_POSITION LOG ENTRIES")
    print("="*50)
    
    # GPT bunu row_img icin soyluyor. Img mi olmalı
    #if row_img is None or row_img.size == 0:
    #    print("⛔ row_img boş, bu satır atlanıyor.")
    #    return None
    print('bbox ==== ', bbox)
    # Başta set edilen box değerlerini kullanıyoruz.
    search_bbox = (bbox[0], bbox[1], bbox[2], bbox[3])
    print(f"\n📐 ARAMA ALANI:")
    print(f"  • Bölge (x1, y1, x2, y2): {search_bbox}")
    print(f"  • Yükseklik: {search_bbox[3] - search_bbox[1]}")

    print('search_bbox ==== ', search_bbox)

    screen = ImageGrab.grab(bbox=search_bbox)
    img = np.array(screen)
    print(f"\n📸 GÖRÜNTÜ BİLGİLERİ:")
    print(f"  • Yakalanan boyut: {img.shape}")

    # Görüntüyü dosyaya kaydet
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"debug_find_image_{timestamp}.png"
    captured_images.append(Image.fromarray(cv2.cvtColor(img, cv2.COLOR_BGR2RGB)))
    cv2.imwrite(filename, img)
    print(f"shape: {img.shape}")  # (height, width) veya (height, width, channels)
    print(f"width: {img.shape[1]}, height: {img.shape[0]}")
    print(f"📸 Find sırasında alınan ekran görüntüleri kaydedildi: {filename}")

    # Görüntüyü işle
    processed_img = preprocess(img, debug_view = False)
    print(f"  • İşlenmiş boyut: {processed_img.shape}")
    
    # Her satırı kontrol et
    found_symbols = []
    print("\n🔎 TARAMA BAŞLIYOR...")
    print(f"  • Başlangıç Y: {search_bbox[1]}")
    print(f"  • Bitiş Y: {search_bbox[3]}")
    
    for y in range(0, processed_img.shape[0], 10):  # 10 piksel adımlarla
        row_img = processed_img[y:y+row_height, :symbol_x_end]
        symbol_text = extract_text(row_img)
        cleaned_symbol = re.sub(r'[^A-Z0-9]', '', symbol_text.upper()).rstrip(':.•·*-')
        
        # Her 100 pikselde bir durum raporu
        if y % 100 == 0:
            print(f"  • Y: {y:4d} | Metin: '{cleaned_symbol}'")
        
        # Eğer geçerli bir sembol bulunduysa
        if len(cleaned_symbol) >= 3:  # En az 3 karakterli semboller
            found_symbols.append((y, cleaned_symbol))
            print(f"\n✅ SEMBOL BULUNDU!")
            print(f"  • Y pozisyonu: {y}")
            print(f"  • Gerçek ekran Y: {search_bbox[1] + y}")  # Ekrandaki gerçek Y pozisyonu
            print(f"  • Sembol: {cleaned_symbol}")
            print(f"  • Ham metin: {symbol_text}")
            print(f"  • Satır yüksekliği: {row_height}")
            print(f"  • İşlenen alan: {y} - {y + row_height}")
            print("="*50)
            return search_bbox[1] + y  # Gerçek ekran Y pozisyonunu döndür
    
    if not found_symbols:
        print("\n❌ SEMBOL BULUNAMADI!")
        print(f"  • Taranan alan: 0-{processed_img.shape[0]} piksel")
        print(f"  • Adım: 10 piksel")
        print(f"  • Min. uzunluk: 3 karakter")
        print("="*50)
    
    return 0  # Eğer bulunamazsa 0 döndür

# ==== RUN ====
if __name__ == "__main__":
    process_rows()
