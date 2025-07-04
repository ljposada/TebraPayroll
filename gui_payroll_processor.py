import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
from payroll_processor import process_payroll

def select_input():
    path = filedialog.askopenfilename(
        filetypes=[('Excel Files', '*.xlsx')],
        title='Selecciona el archivo de reporte')
    if path:
        input_var.set(path)


def select_output():
    path = filedialog.askdirectory(title='Selecciona la carpeta de salida')
    if path:
        output_var.set(path)


def run_processing():
    infile = input_var.get()
    outdir = output_var.get()
    if not infile or not os.path.exists(infile):
        messagebox.showerror('Error', 'Selecciona un archivo de entrada válido.')
        return
    if not outdir or not os.path.isdir(outdir):
        messagebox.showerror('Error', 'Selecciona una carpeta de salida válida.')
        return
    basename = os.path.basename(infile)
    name, _ = os.path.splitext(basename)
    outfile = os.path.join(outdir, f'{name}_consolidado.xlsx')
    try:
        messagebox.showinfo('Procesando', f'Procesando {basename}...')
        process_payroll(infile, outfile)
        messagebox.showinfo('Éxito', f'Archivo generado en:\n{outfile}')
    except Exception as e:
        messagebox.showerror('Error durante el procesamiento', str(e))


def main():
    global input_var, output_var
    root = tk.Tk()
    root.title('Payroll Processor GUI')
    root.resizable(False, False)

    input_var = tk.StringVar()
    output_var = tk.StringVar()

    frame = tk.Frame(root, padx=10, pady=10)
    frame.pack()

    tk.Label(frame, text='Payroll Processor', font=('Arial', 16, 'bold')).grid(row=0, column=0, columnspan=3, pady=(0,10))
    tk.Label(frame, text='Archivo de reporte:').grid(row=1, column=0, sticky='e')
    tk.Entry(frame, textvariable=input_var, width=40).grid(row=1, column=1)
    tk.Button(frame, text='Examinar', command=select_input).grid(row=1, column=2, padx=(5,0))

    tk.Label(frame, text='Carpeta de salida:').grid(row=2, column=0, sticky='e', pady=(5,0))
    tk.Entry(frame, textvariable=output_var, width=40).grid(row=2, column=1, pady=(5,0))
    tk.Button(frame, text='Examinar', command=select_output).grid(row=2, column=2, padx=(5,0), pady=(5,0))

    tk.Button(frame, text='Procesar', command=run_processing).grid(row=3, column=1, pady=(10,0))

    root.mainloop()

if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        messagebox.showerror('Error crítico', str(e))
        sys.exit(1)
