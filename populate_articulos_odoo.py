import os
import pdfplumber
import pandas as pd
import re
from tkinter import filedialog, messagebox, PhotoImage, simpledialog
import tkinter as tk
from PIL import Image, ImageTk

common_headers = ['product_tag_ids', 'type', 'sale_ok',
               'route_ids', 'categ_id']


def populate_articulos_odoo(input_file, label_widget):
    headers = ['id', 'name', 'product_tag_ids', 'standard_price', 'type', 'sale_ok', 'seller_ids', 'seller_ids/price',
               'route_ids', 'categ_id']
    lines= []

    df = pd.read_excel(input_file)
    for index, row  in df.iterrows():
        pop_line = {}
        pop_line['id']=""
        pop_line['name']= row['Referencia']
        pop_line['product_tag_ids']= getattr(label_widget, 'product_tag_ids', None)
        pop_line['standard_price']= row['Precio unitario']
        pop_line['type']= getattr(label_widget, 'type', None)
        pop_line['sale_ok']= getattr(label_widget, 'sale_ok', None)
        pop_line['seller_ids']= row['Proveedor']
        pop_line['seller_ids/price']= row['Precio unitario']
        pop_line['route_ids']= getattr(label_widget, 'route_ids', None)
        pop_line['categ_id']= getattr(label_widget, 'categ_id', None)
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
    variables = {}
    for item in common_headers:
        variables[item] = simpledialog.askstring("Variables comunes", item, parent=label_widget)
        setattr(label_widget, item, variables[item])



def upload_file(label_widget):
    file_path = filedialog.askopenfilename(filetypes=[("EXCEL files", "*.xlsx")])
    if file_path:
        file_name = os.path.basename(file_path)
        file_name_without_extension = file_name.split('.')[0]
        label_widget.config(text=file_name)
        label_widget.just_name = file_name_without_extension
        label_widget.file_path = file_path  # Store the file path
        label_widget.icon_label = tk.Label(root, image=pdf_icon)
        label_widget.icon_label.pack(side="left", padx=5)
        common_variables(label_widget)

def clear_file(label_widget):
    label_widget.config(text="")
    label_widget.file_path = None  # Clear the stored file path
    if hasattr(label_widget, 'icon_label'):
        label_widget.icon_label.destroy()  # Remove the icon if it exists

def create_app():
    global common_headers
    global import_articles_label
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



    def show_albaran_ui():
        hide_all_frames()
        import_articles_frame.pack(fill="both", expand=True)
        import_articles_label.pack()
        tk.Button(import_articles_frame, text="Cargar Excel",
                  command=lambda: upload_file(import_articles_label)).pack(pady=5)
        tk.Button(import_articles_frame, text="Clear", command=lambda: clear_file(import_articles_label)).pack(pady=5)
        tk.Button(import_articles_frame, text="Crear excel para importar articulos a Odoo",
                  command= process_albaran).pack(pady=5)
        import_articles_label.pack()

    def hide_all_frames():
        import_articles_frame.pack_forget()
        ldm_frame.pack_forget()


    # Main UI
    tk.Label(root, text="").pack(pady=10)
    tk.Label(root, text="OPCIONES:").pack(pady=10)
    tk.Label(root, text="").pack(pady=10)
    tk.Button(root, text="Albarán excel->artículos Odoo", command=show_albaran_ui, bg="#f3f2f1",
              fg="blue").pack(pady=5)
    tk.Button(root, text="Albarán excel->Ldm Odoo", command=show_albaran_ui, bg="#f0f0f0", fg="blue").pack(
        pady=5)
    # Frames for different processes
    import_articles_frame = tk.Frame(root)
    ldm_frame = tk.Frame(root)

    # Labels for displaying selected files and results
    import_articles_label = tk.Label(import_articles_frame, text="")
    ldm_label = tk.Label(ldm_frame, text="")


    root.mainloop()

if __name__ == "__main__":
     create_app()
