import pdfplumber
import pandas as pd
import openpyxl
import re
from tkinter import filedialog, messagebox
import tkinter as tk

# Function to convert BOM PDF to Excel
def convert_bom_to_excel(pdf_path, output_path):
    headers = ['level', 'pos', 'article', 'cantidad']
    all_lines_match = []
    error_lines = []
    pattern = r'^(\d) (\d{4}) (.*?)(?: (\d,\d{3}))?$'

    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                lines = text.split('\n')
                for line in lines:
                    match = re.match(pattern, line)
                    if match:
                        level = match.group(1)
                        pos = match.group(2)
                        if pos == '0000':
                            article = match.group(3)
                            qty = 1
                            all_lines_match.append((level, pos, article, qty))
                        else:
                            pattern_01 = r'^(\S+ .*?) (\d+,\d{3})'
                            match_01 = re.match(pattern_01, match.group(3))
                            if match_01:
                                article = match_01.group(1)
                                qty = int(match_01.group(2).split(',')[0])
                                all_lines_match.append((level, pos, article, qty))
                            else:
                                error_lines.append(line)

        df_all = pd.DataFrame(all_lines_match, columns=headers)
        df_all.to_excel(output_path, index=False)
        return True, error_lines  # Return status and errors

    except Exception as e:
        return False, str(e)  # Return failure status and error message

# Function to convert CMO PDF to Excel
def convert_cmo_to_excel(pdf_path, output_path, uds_albaran):
    headers = ['#', 'Cantidad', 'Referencia', 'Precio unitario']
    all_lines_match = []
    pattern_01 = r'^(\d{3}) (\d+)'
    pattern_03 = r'^\d{3} \d+ (.+?) S'
    pattern_04 = r'(\d{1,3}(?:,\d{2}))'

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            lines = text.split('\n')
            for line in lines:
                matches = re.match(pattern_01, line)
                if matches:
                    ref = matches.group(1)
                    cantidad = int(matches.group(2)) // uds_albaran
                    match_articulo = re.search(pattern_03, line)
                    match_precio = re.search(pattern_04, line)
                    if match_articulo and match_precio:
                        articulo = match_articulo.group(1)
                        precio = float((match_precio.group(1).replace(',', '.')))
                        all_lines_match.append([ref, cantidad, articulo, precio])

    df = pd.DataFrame(all_lines_match, columns=headers)
    df.to_excel(output_path, index=False)

# Function to convert EBAKILAN PDF to Excel
def convert_ebakilan_to_excel(pdf_path, output_path, uds_albaran):
    headers = ['#', 'Cantidad', 'Referencia', 'Precio unitario']
    all_lines_match = []
    pattern_10 = r'^PIEZA (.+?)'
    pattern = r'355700 - (\d+) (\S+) (\d+)'
    pattern_04 = r'(\d{1,3}(?:,\d{2}))'

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            lines = text.split('\n')
            for line in lines:
                matches = re.match(pattern_10, line)
                if matches:
                    extract_match = re.search(pattern, line)
                    if extract_match:
                        ref = extract_match.group(1)
                        cantidad = int(extract_match.group(3)) // uds_albaran
                        articulo = extract_match.group(2)
                        precio_match = re.search(pattern_04, line)
                        if precio_match:
                            precio = float((precio_match.group(1).replace(',', '.')))
                            all_lines_match.append([ref, articulo, cantidad, precio])

    df = pd.DataFrame(all_lines_match, columns=headers)
    df.to_excel(output_path, index=False)

# Tkinter UI Function for BOM
def upload_and_process_bom():
    pdf_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if pdf_path:
        output_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if output_path:
            success, errors = convert_bom_to_excel(pdf_path, output_path)
            if success:
                messagebox.showinfo("Success", f"BOM converted successfully and saved to {output_path}")
                if errors:
                    error_log_path = output_path.replace('.xlsx', '_errors.txt')
                    with open(error_log_path, 'w') as f:
                        f.write('\n'.join(errors))
                    messagebox.showwarning("Warnings", f"Some lines couldn't be processed. Check {error_log_path}")
            else:
                messagebox.showerror("Error", f"An error occurred: {errors}")

# Tkinter UI Function for CMO
def upload_and_process_cmo():
    pdf_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if pdf_path:
        output_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if output_path:
            uds_albaran = int(input('De cuantas uds es el albarán?'))
            convert_cmo_to_excel(pdf_path, output_path, uds_albaran)
            messagebox.showinfo("Success", f"CMO components list processed and saved to {output_path}")

# Tkinter UI Function for EBAKILAN
def upload_and_process_ebakilan():
    pdf_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if pdf_path:
        output_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if output_path:
            uds_albaran = int(input('De cuantas uds es el albarán?'))
            convert_ebakilan_to_excel(pdf_path, output_path, uds_albaran)
            messagebox.showinfo("Success", f"EBAKILAN components list processed and saved to {output_path}")

# Tkinter UI setup
def create_app():
    root = tk.Tk()
    root.title("PDF to Excel Converter")

    tk.Label(root, text="Upload PDF for BOM Materials (PUTZMEISTER)").pack()
    tk.Button(root, text="Upload BOM", command=upload_and_process_bom).pack()

    tk.Label(root, text="Upload PDF for CMO (vendor) List of Components").pack()
    tk.Button(root, text="Upload CMO", command=upload_and_process_cmo).pack()

    tk.Label(root, text="Upload PDF for EBAKILAN (vendor) List of Components").pack()
    tk.Button(root, text="Upload EBAKILAN", command=upload_and_process_ebakilan).pack()

    root.mainloop()

if __name__ == "__main__":
    create_app()
