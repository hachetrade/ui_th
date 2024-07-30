import sys
import os
import pdfplumber
import pandas as pd
import re
from tkinter import filedialog, messagebox, PhotoImage, simpledialog
import tkinter as tk
from PIL import Image, ImageTk

# Sample data for checkboxes
CHECKBOX_OPTIONS = {
    'tipo articulo': ['Almacenable', 'Consumible', 'Servicio'],
    'Venta': ['Si', 'No'],
    'Rutas': ['Obtener Bajo Pedido (MTO)', 'Comprar', 'Fabricar'],
    'Categoria': ['All', 'All / Materiales', 'All / Materiales / Accesorios', 'All / Materiales / Aceros', 'All / Materiales / Oxicorte', 'All / Producto Terminado / Repuestos']
}

common_headers = ['product_tag_ids', 'type', 'sale_ok',
               'route_ids', 'categ_id']

def resource_path(relative_path):
    """ Get the absolute path to a resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except AttributeError:
        # If the application is not bundled with PyInstaller, this will be the base path
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)







# Function to convert BOM PDF to Excel
def populate_articulos_odoo(input_file, label_widget):
    headers = ['id', 'name', 'product_tag_ids', 'standard_price', 'type', 'sale_ok', 'seller_ids', 'seller_ids/price',
               'route_ids', 'categ_id', 'seller_ids/min_qty']
    lines= []

    df = pd.read_excel(input_file)
    for index, row  in df.iterrows():
        pop_line = {}
        pop_line['id']=""
        pop_line['name']= row['Referencia']
        pop_line['product_tag_ids']= getattr(label_widget, 'etiquetas', None)
        pop_line['standard_price']= row['Precio unitario']
        pop_line['type']= getattr(label_widget, 'tipo articulo', None)
        pop_line['sale_ok']= 1 if getattr(label_widget, 'venta', None) =='Si' else 0
        pop_line['seller_ids']= row['Proveedor']
        pop_line['seller_ids/price']= row['Precio unitario']
        pop_line['route_ids']= getattr(label_widget, 'Rutas', None)
        pop_line['categ_id']= getattr(label_widget, 'Categoria', None)
        pop_line['seller_ids/min_qty']= 1
        lines.append(pop_line)

    df_final = pd.DataFrame(lines, columns=headers)
    output_file = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile = input_file.split('.')[0]+"_import_odoo", filetypes=[("Excel files", "*.xlsx")])
    if output_file:
        try:
            df_final.to_excel(output_file, index=False)
            messagebox.showinfo("Hecho", f"documento {output_file} generado con éxito")
        except Exception as e:
            messagebox.showerror("ERROR", f"Error generando archivo {output_file}")



def process_albaran():
    or_file_path = import_articles_label.file_path
    if or_file_path:
        populate_articulos_odoo(or_file_path, import_articles_label)

def common_variables(label_widget):
    variable_values = {}

    def on_submit():
        # Store selected options in label_widget attributes
        try:
            etiquetas = etiqueta_entry.get()
            setattr(label_widget, 'etiquetas', etiquetas)
            # Store selected options in label_widget attributes
            for var_name, var_value in variable_values.items():
                if isinstance(var_value, dict):  # Ensure var_value is a dictionary
                    selected = [option for option, var in var_value.items() if var.get()]
                    # Join the selected options into a single string
                    formatted_selection = ','.join(selected)
                    setattr(label_widget, var_name, formatted_selection)
                else:
                    print(f"Unexpected data type for var_value: {type(var_value)}")  # Debugging line
        except Exception as e:
            print(f"Error in on_submit: {e}")
        finally:
            popup.destroy()

    popup = tk.Toplevel()
    popup.title("Establece las variables comunes al lote a importar")
    popup.geometry("400x600")

    tk.Label(popup, text='Etiquetas').grid(row=0, column=0, padx=10, pady=5, sticky='w')
    etiqueta_entry = tk.Entry(popup, width=40)
    etiqueta_entry.grid(row=0, column=1, padx=10, pady=5, sticky='w')

    row = 1
    # variable_values['etiquetas'] = simpledialog.askstring("Etiquetas",'etiquetas', parent=label_widget)
    for var_name, options in CHECKBOX_OPTIONS.items():
        tk.Label(popup, text=var_name).grid(row=row, column=0, padx=10, pady=5, sticky='w')
        variable_values[var_name] = {}
        for option in options:
            var = tk.BooleanVar()
            chk = tk.Checkbutton(popup, text=option, variable=var)
            chk.grid(row=row, column=1, padx=10, pady=5, sticky='w')
            variable_values[var_name][option] = var
            row += 1

    submit_button = tk.Button(popup, text="Submit", command=on_submit)
    submit_button.grid(row=row, columnspan=2, pady=10)

    popup.transient(root)
    popup.grab_set()
    root.wait_window(popup)
    process_albaran()


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
    headers = ['#', 'Cantidad', 'Referencia', 'Precio unitario', 'Proveedor']
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
                    Proveedor = 'CORTES METALURGICOS OVIEDO, S.L.'
                    cantidad = int(matches.group(2)) // uds_albaran
                    match_articulo = re.search(pattern_03, line)
                    match_precio = re.search(pattern_04, line)
                    if match_articulo and match_precio:
                        articulo = match_articulo.group(1)
                        precio = float((match_precio.group(1).replace(',', '.')))
                        all_lines_match.append([ref, cantidad, articulo, precio, Proveedor])

    df = pd.DataFrame(all_lines_match, columns=headers)
    df.to_excel(output_path, index=None)

# Function to convert EBAKILAN PDF to Excel
def convert_ebakilan_to_excel(pdf_path, output_path, uds_albaran):
    headers = ['#', 'Referencia', 'Cantidad', 'Precio unitario', 'Proveedor']
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
                    Proveedor = 'EBAKILAN TOLOSA S.L.'
                    extract_match = re.search(pattern, line)
                    if extract_match:
                        ref = extract_match.group(1)
                        cantidad = int(extract_match.group(3)) // uds_albaran
                        articulo = extract_match.group(2)
                        precio_match = re.search(pattern_04, line)
                        if precio_match:
                            precio = float((precio_match.group(1).replace(',', '.')))
                            all_lines_match.append([ref, articulo, cantidad, precio, Proveedor])

    df = pd.DataFrame(all_lines_match, columns=headers)
    df.to_excel(output_path, index=False)

# Tkinter UI functions for file management
def upload_file(label_widget, choose, is_components_list=False):
    file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if file_path:
        file_name = os.path.basename(file_path)
        label_widget.config(text=file_name)
        label_widget.file_path = file_path  # Store the file path
        # label_widget.icon_label = tk.Label(root, image=pdf_icon)
        # label_widget.icon_label.pack(side="left", padx=5)

    match choose:
        case 1:
            upload_and_process_bom()
        case 2:
            ask_for_units(label_widget)
            upload_and_process_cmo(label_widget)
        case 3:
            ask_for_units(label_widget)
            upload_and_process_ebakilan(label_widget)
        case _:
            messagebox.showerror("Error", f"has elegido {choose}")


#upload albaran excel
def upload_file_ex(label_widget):
    file_path = filedialog.askopenfilename(filetypes=[("EXCEL files", "*.xlsx")])
    if file_path:
        file_name = os.path.basename(file_path)
        file_name_without_extension = file_name.split('.')[0]
        label_widget.config(text=file_name)
        label_widget.just_name = file_name_without_extension
        label_widget.file_path = file_path  # Store the file path
        # label_widget.icon_label = tk.Label(root, image=pdf_icon)
        # label_widget.icon_label.pack(side="left", padx=5)
        common_variables(label_widget)


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

def upload_and_process_cmo(label_widget):
    print(label_widget.file_path, label_widget.units)
    pdf_path = label_widget.file_path
    units = getattr(cmo_pdf_label, 'units', None)
    if pdf_path and units:
        output_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if output_path:
            convert_cmo_to_excel(pdf_path, output_path, units)
            messagebox.showinfo("Success", f"CMO components list processed and saved to {output_path}")
            cmo_excel_label.config(text=os.path.basename(output_path))

def upload_and_process_ebakilan(label_widget):
    pdf_path = label_widget.file_path
    units = getattr(ebakilan_pdf_label, 'units', None)
    if pdf_path and units:
        output_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if output_path:
            convert_ebakilan_to_excel(pdf_path, output_path, units)
            messagebox.showinfo("Success", f"EBAKILAN components list processed and saved to {output_path}")
            ebakilan_excel_label.config(text=os.path.basename(output_path))

# Tkinter UI setup
def create_app():
    global common_headers, choose
    global bom_pdf_label, cmo_pdf_label, ebakilan_pdf_label, import_articles_label
    global bom_excel_label, cmo_excel_label, ebakilan_excel_label
    # global pdf_icon
    global root
    root = tk.Tk()
    root.title("TH convertidor")
    root.geometry("500x400")  # Set window size

    # Load images
    logo_path = resource_path("logos/thlogo.png")
    customer_logo = PhotoImage(file=logo_path)
    logo_label = tk.Label(root, image=customer_logo)
    logo_label.pack()

    def hide_all_frames():
        bom_frame.pack_forget()
        cmo_frame.pack_forget()
        ebakilan_frame.pack_forget()
        import_articles_frame.pack_forget()

        # Main UI
    tk.Label(root, text="Elige opción:").pack(pady=10)
    tk.Button(root, text="CMO BOM pdf", command=lambda: upload_file(label_widget=bom_pdf_label, choose=1)).pack(pady=5)
    tk.Button(root, text="CMO albarán pdf", command=lambda: upload_file(label_widget=cmo_pdf_label, choose=2)).pack(pady=5)
    tk.Button(root, text="EBAKILAN albarán pdf", command=lambda: upload_file(label_widget=ebakilan_pdf_label, choose=3)).pack(pady=5)
    tk.Label(root, text="").pack(pady=10)
    tk.Button(root, text="Generar archivo de importación", command=lambda: upload_file_ex(import_articles_label), bg="#f3f2f1", fg="blue").pack(pady=5)

    # Frames for different processes
    bom_frame = tk.Frame(root)
    cmo_frame = tk.Frame(root)
    ebakilan_frame = tk.Frame(root)
    import_articles_frame = tk.Frame(root)


    # Labels for displaying selected files and results
    bom_pdf_label = tk.Label(bom_frame, text="", compound='left', fg="blue")
    cmo_pdf_label = tk.Label(cmo_frame, text="", compound='left', fg="blue")
    ebakilan_pdf_label = tk.Label(ebakilan_frame, text="", compound='left', fg="blue")
    bom_excel_label = tk.Label(bom_frame, text="", fg="green")
    cmo_excel_label = tk.Label(cmo_frame, text="", fg="green")
    import_articles_label = tk.Label(import_articles_frame, text="")
    ebakilan_excel_label = tk.Label(ebakilan_frame, text="", fg="green")

    root.mainloop()

if __name__ == "__main__":
    create_app()
