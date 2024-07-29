import os
import pdfplumber
import pandas as pd
import re
from tkinter import filedialog, messagebox, PhotoImage, simpledialog
import tkinter as tk
from PIL import Image, ImageTk

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

# Tkinter UI functions for file management
def upload_file(label_widget, is_components_list=False):
    file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if file_path:
        file_name = os.path.basename(file_path)
        label_widget.config(text=file_name)
        label_widget.file_path = file_path  # Store the file path
        label_widget.icon_label = tk.Label(root, image=pdf_icon)
        label_widget.icon_label.pack(side="left", padx=5)
        if is_components_list:
            ask_for_units(label_widget)

def upload_file_ex(label_widget, is_components_list=False):
    file_path = filedialog.askopenfilename(filetypes=[("EXCEL files", "*.xlsx")])
    if file_path:
        file_name = os.path.basename(file_path)
        label_widget.config(text=file_name)
        label_widget.file_path = file_path  # Store the file path
        label_widget.icon_label = tk.Label(root, image=pdf_icon)
        label_widget.icon_label.pack(side="left", padx=5)

def clear_file(label_widget):
    label_widget.config(text="")
    label_widget.file_path = None  # Clear the stored file path
    if hasattr(label_widget, 'icon_label'):
        label_widget.icon_label.destroy()  # Remove the icon if it exists

def ask_for_units(label_widget):
    units = simpledialog.askinteger("Input", "Enter the number of units per albarán:", minvalue=1, parent=root)
    if units is not None:
        label_widget.units = units  # Store the units for further processing

def upload_and_process_bom():
    pdf_path = bom_pdf_label.file_path
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
                bom_excel_label.config(text=os.path.basename(output_path))
            else:
                messagebox.showerror("Error", f"An error occurred: {errors}")

def upload_and_process_cmo():
    pdf_path = cmo_pdf_label.file_path
    units = getattr(cmo_pdf_label, 'units', None)
    if pdf_path and units:
        output_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if output_path:
            convert_cmo_to_excel(pdf_path, output_path, units)
            messagebox.showinfo("Success", f"CMO components list processed and saved to {output_path}")
            cmo_excel_label.config(text=os.path.basename(output_path))

def upload_and_process_ebakilan():
    pdf_path = ebakilan_pdf_label.file_path
    units = getattr(ebakilan_pdf_label, 'units', None)
    if pdf_path and units:
        output_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if output_path:
            convert_ebakilan_to_excel(pdf_path, output_path, units)
            messagebox.showinfo("Success", f"EBAKILAN components list processed and saved to {output_path}")
            ebakilan_excel_label.config(text=os.path.basename(output_path))

# Tkinter UI setup
def create_app():
    global bom_pdf_label, cmo_pdf_label, ebakilan_pdf_label
    global bom_excel_label, cmo_excel_label, ebakilan_excel_label
    global pdf_icon
    global root
    root = tk.Tk()
    root.title("PDF to Excel Converter")
    root.geometry("600x700")  # Set window size

    # Load images
    customer_logo = PhotoImage(file='logos/thlogo.png')
    pdf_image = Image.open('logos/pdf_icon.png')
    pdf_image = pdf_image.resize((10, 10), Image.Resampling.LANCZOS)  # Resize to 10 pixels
    pdf_icon = ImageTk.PhotoImage(pdf_image)

    # Display customer logo
    tk.Label(root, image=customer_logo).pack()

    def show_bom_ui():
        hide_all_frames()
        bom_frame.pack(fill="both", expand=True)
        bom_pdf_label.pack()
        tk.Button(bom_frame, text="Cargar CMO BOM", command=lambda: upload_file(bom_pdf_label)).pack(pady=5)
        tk.Button(bom_frame, text="Clear", command=lambda: clear_file(bom_pdf_label)).pack(pady=5)
        tk.Button(bom_frame, text="Convertir BOM to Excel", command=upload_and_process_bom).pack(pady=5)
        bom_excel_label.pack()

    def show_cmo_ui():
        hide_all_frames()
        cmo_frame.pack(fill="both", expand=True)
        cmo_pdf_label.pack()
        tk.Button(cmo_frame, text="Cargar albarán CMO", command=lambda: upload_file(cmo_pdf_label, is_components_list=True)).pack(pady=5)
        tk.Button(cmo_frame, text="Clear", command=lambda: clear_file(cmo_pdf_label)).pack(pady=5)
        tk.Button(cmo_frame, text="Convertir albarán CMO to Excel", command=upload_and_process_cmo).pack(pady=5)
        cmo_excel_label.pack()

    def show_ebakilan_ui():
        hide_all_frames()
        ebakilan_frame.pack(fill="both", expand=True)
        ebakilan_pdf_label.pack()
        tk.Button(ebakilan_frame, text="Cargar albarán EBAKILAN", command=lambda: upload_file(ebakilan_pdf_label, is_components_list=True)).pack(pady=5)
        tk.Button(ebakilan_frame, text="Clear", command=lambda: clear_file(ebakilan_pdf_label)).pack(pady=5)
        tk.Button(ebakilan_frame, text="Convertir albarán EBAKILAN to Excel", command=upload_and_process_ebakilan).pack(pady=5)
        ebakilan_excel_label.pack()

    def hide_all_frames():
        bom_frame.pack_forget()
        cmo_frame.pack_forget()
        ebakilan_frame.pack_forget()

    # Main UI
    tk.Label(root, text="").pack(pady=10)
    tk.Label(root, text="OPCIONES:").pack(pady=10)
    tk.Button(root, text="CMO BOM->excel", command=show_bom_ui).pack(pady=5)
    tk.Button(root, text="CMO albarán->excel", command=show_cmo_ui).pack(pady=5)
    tk.Button(root, text="EBAKILAN albarán->excel", command=show_ebakilan_ui).pack(pady=5)
    tk.Label(root, text="").pack(pady=10)
    tk.Button(root, text="Albarán excel->artículos Odoo", command=show_ebakilan_ui, bg="#f3f2f1",
              fg="blue").pack(pady=5)
    tk.Button(root, text="Albarán excel->Ldm Odoo", command=show_ebakilan_ui, bg="#f0f0f0", fg="blue").pack(
        pady=5)
    # Frames for different processes
    bom_frame = tk.Frame(root)
    cmo_frame = tk.Frame(root)
    ebakilan_frame = tk.Frame(root)

    # Labels for displaying selected files and results
    bom_pdf_label = tk.Label(bom_frame, text="", compound='left', fg="blue")
    cmo_pdf_label = tk.Label(cmo_frame, text="", compound='left', fg="blue")
    ebakilan_pdf_label = tk.Label(ebakilan_frame, text="", compound='left', fg="blue")
    bom_excel_label = tk.Label(bom_frame, text="", fg="green")
    cmo_excel_label = tk.Label(cmo_frame, text="", fg="green")
    ebakilan_excel_label = tk.Label(ebakilan_frame, text="", fg="green")

    root.mainloop()

if __name__ == "__main__":
    create_app()
