import tkinter as tk
from tkinter import filedialog
import pandas as pd
import numpy as np
import notif
from tkinter import *

files_bca = []
files_bri = []
files_permata = []

def cetak():
    print(files_bca)
    print(files_bri)
    print(files_permata)

    df_akhir = pd.DataFrame([[]])

    df_bca = pd.DataFrame([[]])
    if len(files_bca) > 0:
        for path in files_bca:
            df = proses_bca(path)
            df_bca = pd.concat([df_bca, df], ignore_index=True)
        df_akhir = pd.concat([df_akhir, df_bca], ignore_index=True)
        df_akhir = df_akhir.dropna(how="all")

    df_bri = pd.DataFrame([[]])
    if len(files_bri) > 0:
        for path in files_bri:
            df = proses_bri(path)
            df_bri = pd.concat([df_bri, df], ignore_index=True)
        df_akhir = pd.concat([df_akhir, df_bri], ignore_index=True)
        df_akhir = df_akhir.dropna(how="all")
    
    df_permata = pd.DataFrame([[]])
    if len(files_permata) > 0:
        for path in files_permata:
            df = proses_permata(path)
            df_permata = pd.concat([df_permata, df], ignore_index=True)
        df_akhir = pd.concat([df_akhir, df_permata], ignore_index=True)
        df_akhir = df_akhir.dropna(how="all")

    save_filename = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                              initialdir="C:/",
                                              title="Save",
                                              filetypes=(('Microsoft Excel', "*.xlsx"), ("csv", "*.csv")))
    if save_filename:
        df_akhir.to_excel(save_filename, index=False)
        notif.success_save()
  
def proses_bca(path):
    if path.split(".")[-1] == "csv":
        df = pd.read_csv(path)
    elif path.split(".")[-1] == "xlsx":
        df = pd.read_excel(path)

    start_rows = 0
    for i, data in enumerate(df[df.columns[0]]):
        if type(data) == str:
            if "tanggal" in data.lower():
                start_rows = i

    if path.split(".")[-1] == "csv":
        df = pd.read_csv(path, skiprows= start_rows+1)
    elif path.split(".")[-1] == "xlsx":
        df = pd.read_excel(path, skiprows= start_rows+1)

    def nominal_after_tax(data):
        hasil = data.replace(" CR", "")
        hasil = hasil.replace(",", "")
        hasil = hasil[:-3]
        hasil = float(hasil)
        return hasil

    df['After Tax'] = df[df.columns[3]].apply(nominal_after_tax)

    def kategori_mutasi(data):
        if ("jemputan" in data.lower()) or ("setoran tunai" in data.lower()):
            return "CS"
        elif ("kartu kredit" in data.lower()):
            return "KR"
        elif ("kr otomatis mid" in data.lower()) and ("tgh" in data.lower()):
            return "DB"
        elif ("kr otomatis tanggal"in data.lower()) and ("qr" in data.lower()):
            return "QR"
        else:
            return ""

    df["Tipe"] = df[df.columns[1]].apply(kategori_mutasi)

    def admin_n_tax(data):
        if "ddr" in data.lower():
            hasil = data.lower().split("ddr:")[-1]
            hasil = hasil.replace(" ", "")
            try:
                hasil = float(hasil)
                final_hasil = int(hasil)
            except:
                final_hasil = 0
            return final_hasil
        elif "adm" in data.lower():
            hasil = data.lower().split("adm:")[-1]
            hasil = hasil.replace(" ", "")
            try:
                hasil = float(hasil)
                final_hasil = int(hasil)
            except:
                final_hasil = 0
            return final_hasil
        else:
            return 0

    df["Tax"] = df[df.columns[1]].apply(admin_n_tax)

    df["Amount"] = df["After Tax"] + df["Tax"]
    df["Amount"] = df["Amount"].astype(int)
    df["Bank"] = ["BCA" for i in df[df.columns[0]]]

    final_df = df[[df.columns[0],df.columns[1], df.columns[5],df.columns[7],"Bank"]]
    return final_df

def proses_permata(path):
    if path.split(".")[-1] == "csv":
        df = pd.read_csv(path)
    elif path.split(".")[-1] == "xlsx":
        df = pd.read_excel(path)

    df = df.drop([df.columns[0],df.columns[1],df.columns[3],df.columns[4],df.columns[6],df.columns[7],df.columns[9],df.columns[10]],axis="columns")

    def tipe_transaksi(data):
        if "qr" in data.lower():
            return "QR"
        elif "payment merchant" in data.lower():
            return "KR"
        else:
            return ""
        
    def amount(data):
        hasil = data.replace(",", "")
        return float(hasil)

    df["Tipe Transaksi"] = df[df.columns[2]].apply(tipe_transaksi)
    df['Amount'] = df['Amount'].apply(amount)
    df = df[[df.columns[0],df.columns[2],df.columns[3],df.columns[1]]]
    df.columns = ["Tanggal Transaksi", "Keterangan", "Tipe", "Amount"]
    df["Bank"] = ["Permata" for i in df[df.columns[0]]]
    df = df[df["Tipe"]!=""]
    df = df.reset_index(drop=True)
    return df

def proses_bri(path):
    if path.split(".")[-1] == "csv":
        df = pd.read_csv(path)
    elif path.split(".")[-1] == "xlsx":
        df = pd.read_excel(path)
    
    start_rows = 0
    for i, data in enumerate(df[df.columns[0]]):
        if type(data) == str:
            if "mid" in data.lower():
                start_rows = i
            break

    if path.split(".")[-1] == "csv":
        df = pd.read_csv(path, skiprows= start_rows+1)
    elif path.split(".")[-1] == "xlsx":
        df = pd.read_excel(path, skiprows= start_rows+1)
        
    df['Keterangan'] = df[df.columns[6]] + df[df.columns[5]]
    df_hasil = df[[df.columns[3],"Keterangan",df.columns[12],df.columns[8]]]

    def tipe(data):
        if "qris" in data.lower():
            return "QR"
        elif "debit" in data.lower():
            return "DB"
        elif "credit" in data.lower():
            return "KR"
        else:
            return ""

    df_hasil["Tipe"] = df_hasil[df_hasil.columns[2]].apply(tipe)

    def amount(data):
        data = data.replace(",","")
        return int(data)

    df_hasil = df_hasil[[df_hasil.columns[0],df_hasil.columns[1],df_hasil.columns[4],df_hasil.columns[3]]]
    df_hasil.columns = ["Tanggal Transaksi", "Keterangan", "Tipe", "Amount"]
    df_hasil["Amount"] = df_hasil["Amount"].apply(amount)
    df_hasil["Bank"] = ["BRI" for i in df_hasil[df_hasil.columns[0]]]
    df_hasil = df_hasil[df_hasil['Tipe']!=""]
    df_hasil = df_hasil.reset_index(drop=True)
    return df_hasil

def main():

    def bca_openfiles():
        global files_bca
        # Membuka dialog file untuk memilih beberapa file
        upload_files = filedialog.askopenfilenames(
            title="Pilih File Excel",
            filetypes=[("CSV Files", "*.csv"),("Excel Files", "*.xls*")]
        )
        filename = []

        for file in upload_files:
            filename.append(file.split("/")[-1])
        
        # Menampilkan nama file di kotak teks
        if upload_files:
            files_bca = upload_files
            bca_box.delete("1.0", tk.END)  # Menghapus teks sebelumnya
            for file_path in filename:
               bca_box.insert(tk.END, file_path + "\n")
        else:
            bca_box.delete("1.0", tk.END)
            files_bca = []

    def bri_openfiles():
        global files_bri
        # Membuka dialog file untuk memilih beberapa file
        upload_files = filedialog.askopenfilenames(
            title="Pilih File Excel",
            filetypes=[("CSV Files", "*.csv"),("Excel Files", "*.xls*")]
        )
        filename = []

        for file in upload_files:
            filename.append(file.split("/")[-1])
        
        # Menampilkan nama file di kotak teks
        if upload_files:
            files_bri = upload_files
            bri_box.delete("1.0", tk.END)  # Menghapus teks sebelumnya
            for file_path in filename:
                bri_box.insert(tk.END, file_path + "\n")
        else:
            bri_box.delete("1.0", tk.END)
            files_bri = []

    def permata_openfiles():
        global files_permata
        # Membuka dialog file untuk memilih beberapa file
        upload_files = filedialog.askopenfilenames(
            title="Pilih File Excel",
            filetypes=[("CSV Files", "*.csv"),("Excel Files", "*.xls*")]
        )
        filename = []

        for file in upload_files:
            filename.append(file.split("/")[-1])
        
        # Menampilkan nama file di kotak teks
        if upload_files:
            files_permata = upload_files
            permata_box.delete("1.0", tk.END)  # Menghapus teks sebelumnya
            for file_path in filename:
                permata_box.insert(tk.END, file_path + "\n")
        else:
            permata_box.delete("1.0", tk.END)
            files_permata = []

    root = tk.Tk()
    root.title("Input Mutasi Bank v3.1")

    width = 385
    height = 740

    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()

    x = (screen_width / 2) - (width / 2)
    y = (screen_height / 2) - (height / 2)

    root.geometry(f"{width}x{height}+{int(x)}+{int(y)}")
    root.resizable(False, False)

    bca_paned = PanedWindow(bd=2, relief="groove")
    bca_paned.place(x=10, y=10, width=368, height=300)
    bca_label = Label(root, text="Tarikan BCA")
    bca_label.place(x=20, y=1)
    bca_box = tk.Text(bca_paned, width=42, height=14, relief="groove", bd=2)
    bca_box.place(x=10, y=10)
    bca_openfiles_btn = tk.Button(bca_paned, width=10, relief="ridge", borderwidth=1, text="Open Files", height=1, command=bca_openfiles)
    bca_openfiles_btn.place(x=255, y=255, width=90)

    bri_paned = PanedWindow(bd=2, relief="groove")
    bri_paned.place(x=10, y=320, width=370, height=180)
    bri_label = Label(root, text="Tarikan BRI")
    bri_label.place(x=20, y=311)
    bri_box = tk.Text(bri_paned, width=42, height=7, relief="groove")
    bri_box.place(x=10, y=10)
    bca_openfiles_btn = tk.Button(bri_paned, width=10, relief="ridge", borderwidth=1, text="Open Files", height=1, command=bri_openfiles)
    bca_openfiles_btn.place(x=255, y=140, width=90)

    permata_paned = PanedWindow(bd=2, relief="groove")
    permata_paned.place(x=10, y=510, width=370, height=180)
    permata_label = Label(root, text="Tarikan Permata")
    permata_label.place(x=20, y=501)
    permata_box = tk.Text(permata_paned, width=42, height=7, relief="groove")
    permata_box.place(x=10, y=10)
    permata_openfiles_btn = tk.Button(permata_paned, width=10, relief="ridge", borderwidth=1, text="Open Files", height=1, command=permata_openfiles)
    permata_openfiles_btn.place(x=255, y=140, width=90)

    cetak_button = tk.Button(root, width=10, relief="ridge", borderwidth=1, text="Simpan", height=1, command=cetak)
    cetak_button.place(x=150, y=700, width=90)

    root.mainloop()

if __name__ == "__main__":
    main()