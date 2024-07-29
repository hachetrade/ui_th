import tkinter as tk
from tkinter import ttk

# Sample data for checkboxes
CHECKBOX_OPTIONS = {
    'tipo articulo': ['Almacenable', 'Comsumible', 'Servicio'],
    'Venta': ['Si', 'No'],
    'Rutas': ['Comprar', 'Fabricar', 'Obtener Bajo Pedido (MTO)'],
    'Categoria': ['All', 'Materiales', 'Accesories', 'Aceros', 'Oxicorte', 'Repuestos']
}


def common_variables(label_widget):
    variable_values = {}

    def on_submit():
        # Store selected options in label_widget attributes
        for var_name, var_value in variable_values.items():
            selected = [option for option, var in var_value.items() if var.get()]
            setattr(label_widget, var_name, selected)
        popup.destroy()

    popup = tk.Toplevel()
    popup.title("Establece las variables comunes al lote a importar")
    popup.geometry("400x600")

    row = 0
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


# Example usage within the main app
root = tk.Tk()
root.title("Odoo Import Helper")

# Assuming import_articles_label is defined somewhere in your app
import_articles_label = tk.Label(root, text="Select File")
import_articles_label.pack()

# Button to set common variables
tk.Button(root, text="Set Common Variables", command=lambda: common_variables(import_articles_label)).pack(pady=10)

root.mainloop()
