import cv2
import pytesseract
import numpy as np
from PIL import ImageGrab
import xlwings as xw
import win32gui
import win32con
import win32process
import time

# ==== CONFIG ====
APP_WINDOW_TITLE = "Borsa İşlem Platformu"  # Replace this with your actual app window title
DEBUG_MODE = True  # Set to False to disable rectangle drawing
EXCEL_PATH = "SenetSepet-TAM11.xlsm"  # Your Excel file
SHEET_NAME = "OCR_list"

# Coordinates and constants
row_height = 42
symbol_x_end = 100
price_x_start = 100
capture_height = 740
start_y = 5
end_y = row_height * 17

screen_x = 10       # Adjust this for your app window
screen_y = 10
screen_width = 500  # Adjust this based on the width needed for prices

# ==== WINDOW CONTROL ====
def bring_window_to_front(window_title):
    def enum_callback(hwnd, _):
        if window_title.lower() in win32gui.GetWindowText(hwnd).lower():
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            win32gui.SetForegroundWindow(hwnd)
    win32gui.EnumWindows(enum_callback, None)
    time.sleep(0.4)  # Give time for window to come to front

# ==== IMAGE PROCESSING ====
def preprocess(img):
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    thresh = cv2.adaptiveThreshold(
        gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY, 11, 2
    )
    kernel = np.ones((1, 1), np.uint8)
    dilated = cv2.dilate(thresh, kernel, iterations=1)
    return dilated

def extract_text(img):
    config = "--psm 7"  # Assume single line
    return pytesseract.image_to_string(img, config=config).strip()

# ==== SCREEN CAPTURE ====
def capture_screen():
    bbox = (screen_x, screen_y, screen_x + screen_width, screen_y + capture_height)
    screen = ImageGrab.grab(bbox=bbox)
    return np.array(screen)

# ==== DEBUG DRAWING ====
def draw_debug_rectangles(img, row_count=17):
    for i in range(row_count):
        y1 = start_y + i * row_height
        y2 = y1 + row_height
        cv2.rectangle(img, (0, y1), (symbol_x_end, y2), (0, 255, 0), 1)
        cv2.rectangle(img, (price_x_start, y1), (screen_width, y2), (255, 0, 0), 1)
    cv2.imshow("Debug - Row Rectangles", img)
    cv2.waitKey(0)
    cv2.destroyAllWindows()

# ==== MAIN PROCESS ====
def process_rows():
    bring_window_to_front(APP_WINDOW_TITLE)
    img = capture_screen()

    if DEBUG_MODE:
        print("t1")
        draw_debug_rectangles(img.copy())
        print("t2")
    exit()
    wb = xw.Book(EXCEL_PATH)
    ws = wb.sheets[SHEET_NAME]

    for i in range(17):  # 0 to 16
        y1 = start_y + i * row_height
        y2 = y1 + row_height
        row_img = img[y1:y2, :]

        symbol_img = preprocess(row_img[:, :symbol_x_end])
        price_img = preprocess(row_img[:, price_x_start:])

        symbol_text = extract_text(symbol_img)
        price_text = extract_text(price_img)

        # Excel output
        ws.range(f"A{i+1}").value = symbol_text
        ws.range(f"B{i+1}").value = price_text

        if DEBUG_MODE:
            print(f"Row {i+1}: {symbol_text} - {price_text}")

# ==== RUN ====
if __name__ == "__main__":
    process_rows()
