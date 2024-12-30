import tkinter as tk
from tkinter import filedialog
import pandas as pd
import numpy as np

file_paths = []

def read_data(path):
  df = pd.read_csv(path)
  start = 0

  for i, id in enumerate(df[df.columns[0]]):
    if type(id) == str:
      if "tanggal" in id.lower():
        start = i
        break
        # print(outlet)

  df0 = pd.read_csv(path, skiprows=i+1)
  df0 = df0[~df0[df0.columns[0]].str.contains('Saldo|Mutasi', case=False, na=False)]
  return df0

def cetak():
    hasil = []
    df_hasil = pd.DataFrame([[]])

    for file in file_paths:
        hasil = read_data(file)
        df_hasil = pd.concat([df_hasil, hasil], ignore_index=True)
    print(df_hasil)
    
    save_filename = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                            initialdir="C:/",
                                            title="Save",
                                            filetypes=(('Microsoft Excel', "*.xlsx"), ("All Files", "*.*")))

    if save_filename:
      df_hasil.to_excel(save_filename, index=False)

def open_files():
    global file_paths
    # Membuka dialog file untuk memilih beberapa file
    file_paths = filedialog.askopenfilenames(
        title="Pilih File Excel",
        filetypes=[("Excel Files", "*.csv")]
    )
    filename = []

    for file in file_paths:
        filename.append(file.split("/")[-1])
    print(filename)
    # Menampilkan nama file di kotak teks
    if file_paths:
        text_box.delete("1.0", tk.END)  # Menghapus teks sebelumnya
        for file_path in filename:
            text_box.insert(tk.END, file_path + "\n")  # Menambahkan nama file

# Membuat jendela utama Tkinter
root = tk.Tk()
root.title("MUTASI LAPORAN SALES")

width = 600
height = 340

screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

x = (screen_width / 2) - (width / 2)
y = (screen_height / 2) - (height / 2)

root.geometry(f"{width}x{height}+{int(x)}+{int(y)}")
root.resizable(False, False)

# Tombol "Open Files"
open_files_button = tk.Button(root, width=10, relief="ridge", borderwidth=1, text="Open Files", command=open_files)
open_files_button.place(x=415, y=300)

cetak_button = tk.Button(root, width=10, relief="ridge", borderwidth=1, text="Cetak", command=cetak)
cetak_button.place(x=500, y=300)

# Kotak teks untuk menampilkan nama file
text_box = tk.Text(root, width=71, height=17)
text_box.place(x=10, y=10)

# Menjalankan aplikasi Tkinter
root.mainloop()