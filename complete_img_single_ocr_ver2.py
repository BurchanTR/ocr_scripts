import os
import sys

import cv2
import re
import time
import pytesseract
import numpy as np
from PIL import Image
import xlwings as xw
import pandas as pd
from concurrent.futures import ProcessPoolExecutor as Executor

# ==== SABİTLER ====
EXCEL_PATH = "SenetSepet-TAM11.xlsm"
SHEET_NAME = "OCR_list"

# -------------------------------------------------
# ALTERNATİF YAPI: Toplu OCR ile işleme - BAŞLANGIÇ
# -------------------------------------------------

def split_columns(image, symbol_ratio=0.6):
    """Görseli sembol ve fiyat kolonu olarak ikiye ayırır."""
    height, width = image.shape[:2]
    split_x = int(width * symbol_ratio)
    symbol_col = image[:, :split_x]
    price_col = image[:, split_x:]
    cv2.imwrite("symbol_column.png", symbol_col)
    cv2.imwrite("price_column.png", price_col)
    return symbol_col, price_col

def ocr_column(image, psm=6, whitelist=None):
    """Tek bir kolon imajından satır satır OCR yapar."""
    config = f'--psm {psm}'
    if whitelist:
        config += f' -c tessedit_char_whitelist={whitelist}'
    text = pytesseract.image_to_string(image, config=config)
    lines = [line.strip() for line in text.split('\n') if line.strip()]
    return lines

def process_combined_image_bulk(image_path):
    """Toplu OCR: Kolonlara ayır, her kolonu OCR yap, eşleştir."""
    print("▶ Toplu OCR başlatılıyor...")

    # Görsel yükle ve CLAHE uygula (daha iyi sonuçlar için)
    image = cv2.imread(image_path)
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
    
    #print(f"Symbols: {(symbols)}")
    #print(f"Prices: {(prices)}")
    #print(f"Symbols satır sayısı: {len(symbols)}")
    #print(f"Prices satır sayısı: {len(prices)}")

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

import re

def clean_text(text):
    return re.sub(r'\s+', ' ', text.strip())

def clean_price(text):
    text = text.replace(' ', '').replace(',', '.')
    return re.sub(r'[^0-9.\-]', '', text)

# -------------------------------------------------
# ALTERNATİF YAPI: Toplu OCR ile işleme - BİTİŞ
# -------------------------------------------------

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

def match_and_write_to_excel_with_xlwings_brch(symbols, prices, excel_path=EXCEL_PATH, sheet_name=SHEET_NAME):
    if len(symbols) != len(prices):
        print(f"⚠️ UYARI: Sembol ({len(symbols)}) ve fiyat ({len(prices)}) sayısı eşleşmiyor!")
        raise ValueError("Sembol ve fiyat sayısı uyuşmuyor. İşlem durduruldu.")

    cleaned_data = []
    for symbol, price in zip(symbols, prices):
        clean_symbol = clean_text(symbol)
        clean_price_val = clean_price(price)
        cleaned_data.append((clean_symbol, clean_price_val))

    df = pd.DataFrame(cleaned_data, columns=['Symbol', 'Price'])

    target_wb_name = os.path.basename(EXCEL_PATH)
    try:
        wb = connect_to_open_workbook(target_wb_name)
    except Exception as e:
        print(str(e))
        sys.exit(1)

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

# === Ana işlem ===
def process_image(image_path):
    start_time = time.time()
    bulk_results = process_combined_image_bulk("merged_output.png")
    bulk_results = remove_duplicates(bulk_results)
    #elapsed = time.time() - start_time
    #print(f"✔ Toplu OCR tamamlandı. {len(bulk_results)} satır. Süre: {elapsed:.2f} saniye.")
    #start_time = time.time()
    symbols, prices = zip(*bulk_results) if bulk_results else ([], [])
    match_and_write_to_excel_with_xlwings_brch(symbols, prices, EXCEL_PATH, SHEET_NAME)
    elapsed = time.time() - start_time
    print(f"✔ Merged_output.png işleme Süresi: {elapsed:.2f} saniye.")

# === Kullanım ===
if __name__ == "__main__":
    process_image("merged_output.png")
