import cv2
import pytesseract
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
from concurrent.futures import ThreadPoolExecutor

# ==== CONFIG ====
bbox = (48, 260, 260, 1000)
APP_WINDOW_TITLE = "Borsa İşlem Platformu"  # Replace this with your actual app window title
DEBUG_MODE = True  # Set to False to disable rectangle drawing
EXCEL_PATH = "SenetSepet-TAM11.xlsm"  # Your Excel file
SHEET_NAME = "OCR_list"
SCROLL_PIXELS = -120
OCR_CONFIG = "--psm 7"
highlight_duration = 5 # seconds

# Coordinates and constants
row_height = 42
NUM_ROWS = 17
symbol_x_end = 100
price_x_start = 100
capture_height = 740
start_y = 5
end_y = row_height * 17
scale_factor = 0.9

screen_x = 48       # Adjust this for your app window
screen_y = 260
screen_width = 212  # Adjust this based on the width needed for prices

## Not in Use. Keeping for possible future usages.
def show_row_and_text(row_img, symbol_text, price_text, i):
    # Convert BGR to RGB and then to PIL Image
    img_rgb = cv2.cvtColor(row_img, cv2.COLOR_BGR2RGB)
    pil_img = Image.fromarray(img_rgb)

    # === IMAGE WINDOW ===
    def _display():
        root = tk.Tk()
        root.withdraw()  # Hide main window

        img_win = tk.Toplevel(root)
        img_win.title(f"Row {i+1}")
        img_win.geometry(f"+100+{100 + i * 30}")
        img_label = tk.Label(img_win)
        img_label.pack()
        tk_img = ImageTk.PhotoImage(pil_img)
        img_label.config(image=tk_img)
        img_label.image = tk_img  # Prevent garbage collection

        text_win = tk.Toplevel(root)
        text_win.title(f"OCR Text {i+1}")
        text_win.geometry(f"+420+{100 + i * 30}")
        tk.Label(text_win, text=f"Symbol: {symbol_text}", font=("Courier", 12)).pack()
        tk.Label(text_win, text=f"Price : {price_text}", font=("Courier", 12)).pack()

        def close_windows(event=None):
            img_win.destroy()
            text_win.destroy()
            root.destroy()

        img_win.bind("<Escape>", close_windows)
        text_win.bind("<Escape>", close_windows)

        img_win.after(3000, close_windows)  # Auto close

        root.mainloop()

    threading.Thread(target=_display).start()

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

# Function to show visual bbox with light blue-gray, very transparent
def show_highlight_box(bbox, duration=4, margin=2, border_thickness=6):
    x1, y1, x2, y2 = bbox
    width = x2 - x1 + margin * 2
    height = y2 - y1 + margin * 2
    x1 -= margin
    y1 -= margin

    def _box():
        blink_interval = 250  # milliseconds
        root = tk.Tk()
        root.overrideredirect(True)
        root.geometry(f"{width}x{height}+{x1}+{y1}")
        root.attributes("-topmost", True)
        root.attributes("-transparentcolor", "white")  # Make white fully transparent

        canvas = tk.Canvas(root, width=width, height=height, highlightthickness=0, bg="white")
        canvas.pack()

        # Draw only the black rectangle border (no fill)
        canvas.create_rectangle(
            border_thickness // 2,
            border_thickness // 2,
            width - border_thickness // 2,
            height - border_thickness // 2,
            outline="black",
            width=border_thickness
        )

        root.after(int(duration * 1000), root.destroy)
        root.mainloop()

    threading.Thread(target=_box).start()

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

    if is_fullscreen:
        print("✅ Uygulama zaten tam ekran.")
    else:
        print("⚠️ Uygulama tam ekran değil. F11 gönderiliyor...")
        win.activate()
        time.sleep(0.1)
        pyautogui.press('f11')
        time.sleep(0.1)
        print("✅ F11 gönderildi.")

    return True

# ==== IMAGE PROCESSING ====
def preprocess(img):
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    thresh = cv2.adaptiveThreshold(
        gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY, 15, 5
    )
    kernel = np.ones((1, 1), np.uint8)
    dilated = cv2.dilate(thresh, kernel, iterations=1)
    return dilated
    # Resize if scale_factor is not 1 (default no resize)
    if scale_factor != 1.0:
        width = int(dilated.shape[1] * scale_factor)
        height = int(dilated.shape[0] * scale_factor)
        resized = cv2.resize(dilated, (width, height), interpolation=cv2.INTER_AREA)
        return resized
    return dilated

def extract_text(img):
    config = "--psm 7"  # Assume single line
    return pytesseract.image_to_string(img, config=config).strip()

# ==== SCREEN CAPTURE ====
def capture_screen():
    bbox = (screen_x, screen_y, screen_x + screen_width, screen_y + capture_height)
    screen = ImageGrab.grab(bbox=bbox)
    return np.array(screen)

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
    cv2.moveWindow("Debug View: Image + OCR Text", 300, 200)
    #cv2.waitKey(0)
    key = cv2.waitKey(8000)  # 8 seconds or key
    cv2.destroyAllWindows()

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

    symbol_img = preprocess(row_img[:, :symbol_x_end])
    price_img  = preprocess(row_img[:, price_x_start:])

    symbol_text = extract_text(symbol_img)
    price_text  = extract_text(price_img)

    # === CLEANING STARTS HERE ===
    cleaned_symbol = re.sub(r'[^A-Z0-9]', '', symbol_text.upper()).rstrip(':.•·*-')

    try:
        temp_price = price_text.replace('.', '')  # binlik ayraç noktaları kaldır
        last_price = temp_price.replace(',', '.')
        cleaned_price = float(last_price)
        cleaned_price = f"{cleaned_price:.2f}"  # Keep 2 decimal places
    except ValueError:
        cleaned_price = ""  # Or set to "N/A"
    # === CLEANING ENDS HERE ===

    return (i, cleaned_symbol, cleaned_price)

def process_rows():
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

    ws.clear_contents()  # veya ws.clear() eğer stiller vs. de silinsin isteniyorsa
    ws.range("A1").value = [["Hisse", "Son Fiyat"]]  # Add header

    if not bring_investing_app_to_front():
        print("⚠️ Please open the investing app in Firefox private window and re-run the script.")
        sys.exit(1)
    # Ensure focus to Right Region at the start

    activate_scroll_area()
    scroll_to_top_fast()

    if DEBUG_MODE:
        show_highlight_box(bbox, highlight_duration)

    full_img = capture_screen()

    start = mark_start()

    results = []
    with ThreadPoolExecutor(max_workers=6) as executor:
        futures = [executor.submit(process_single_row, i, full_img) for i in range(NUM_ROWS)]
        for future in futures:
            results.append(future.result())

    # Sort by row index to ensure correct Excel order
    results.sort(key=lambda x: x[0])

    # Excel output
    for i, symbol_text, price_text in results:
        ws.range(f"A{i+2}").value = symbol_text
        ws.range(f"B{i+2}").value = price_text

    end = mark_end()

    if DEBUG_MODE:
        results.sort(key=lambda x: x[0])  # Sort by index        
        texts = [(s, p) for _, s, p in results]
        draw_debug_rectangles_with_text(full_img, texts)

    # 🔁 Return focus to Excel
    excel_windows = [w for w in gw.getWindowsWithTitle("Excel") if w.visible]

    if excel_windows:
        excel_window = excel_windows[0]
        excel_window.activate()
        # Optional: maximize if you want full screen
        # excel_window.maximize()
        print("✅ Focus returned to Excel.")
    else:
        print("❌ Excel window not found.")

# ==== RUN ====
if __name__ == "__main__":
    process_rows()
