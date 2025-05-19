#### This will be first public ocr supported inventing app update solution...
import pytesseract
import pyautogui
import threading
import cv2
import numpy as np
from PIL import ImageGrab
import time

# Sabitler
start_y = 5
row_height = 42
symbol_x_end = 100
price_x_start = 100
rows_per_screen = 17

# OCR sonuçları geçici listeye yazılacak
ocr_results = []

def preprocess(img):
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    thresh = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                   cv2.THRESH_BINARY, 11, 2)
    return thresh

def ocr_thread(row_img, row_index, results):
    symbol_img = preprocess(row_img[:, :symbol_x_end])
    price_img  = preprocess(row_img[:, price_x_start:])
    
    symbol_text = pytesseract.image_to_string(symbol_img, config='--psm 7').strip()
    price_text  = pytesseract.image_to_string(price_img, config='--psm 7').strip()

    results[row_index] = (symbol_text, price_text)

def process_screen(img):
    threads = []
    results = [None] * rows_per_screen

    for i in range(rows_per_screen):
        y1 = start_y + i * row_height
        y2 = y1 + row_height
        row_img = img[y1:y2, :]

        thread = threading.Thread(target=ocr_thread, args=(row_img, i, results))
        threads.append(thread)
        thread.start()

    for thread in threads:
        thread.join()

    return results

if __name__ == "__main__":
    time.sleep(10)  # Ekranı hazırla, 3 saniye bekle
    screen = np.array(ImageGrab.grab())
    results = process_screen(screen)

    for i, (symbol, price) in enumerate(results):
        print(f"{i+1:02d}: {symbol} --> {price}")

