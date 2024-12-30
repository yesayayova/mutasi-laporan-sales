from tkinter import *
from tkinter import filedialog
from tkinter import messagebox

def success_save():
    messagebox.showinfo("Notifikasi", "Data berhasil disimpan")

def fail_save():
    messagebox.showinfo("Notifikasi", "Data gagal disimpan")

def error(tarikan):
    messagebox.showinfo("Notifikasi", "Error pada data "+tarikan)