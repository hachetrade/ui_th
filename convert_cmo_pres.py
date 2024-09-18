import sys
import os
import pdfplumber
import pandas as pd
import re
from tkinter import filedialog, messagebox, PhotoImage, simpledialog
import tkinter as tk
from PIL import Image, ImageTk


pdf_path = "pres_cmo.pdf"
out_path = "prescmo"


def convert_cmo_pres(pdf_path, output_path, uds_albaran):
    headers = ['#', 'Cantidad', 'Referencia', 'Precio unitario', 'Proveedor']
    all_lines_match = []
    pattern_01 = r'^(\d{3}) (\d+)'
    pattern_03 = r'^\d{3} \d+ (.+?) S'
    pattern_04 = r'(\d{1,3}(?:,\d{2}))'

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            lines = text.split('\n')
            print(lines)
            for line in lines:
                print(line)
                matches = re.match(pattern_01, line)
                if matches:
                    ref = matches.group(1)
                    Proveedor = 'CORTES METALURGICOS OVIEDO, S.L.'
                    cantidad = int(matches.group(2)) // uds_albaran
                    match_articulo = re.search(pattern_03, line)
                    match_precio = re.search(pattern_04, line)
                    if match_articulo and match_precio:
                        articulo = match_articulo.group(1)
                        precio = float((match_precio.group(1).replace(',', '.')))
                        all_lines_match.append([ref, cantidad, articulo, precio, Proveedor])

    df = pd.DataFrame(all_lines_match, columns=headers)
    print(df)
    df.to_excel(output_path, index=None)

    convert_cmo_pres(pdf_path, output_path, 1)
