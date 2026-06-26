import tkinter as tk
from tkinter import messagebox
import csv
from datetime import datetime

# Fungsi untuk menyimpan data
def save_data():
    nama = entry_nama.get().strip()
    prodi = entry_prodi.get().strip()
    nomor = entry_nomor.get().strip()

    # Validasi input
    if not nama and not prodi and not nomor:
        messagebox.showerror("Error", "Tidak ada data yang diisi!")
        return

    # Simpan data ke CSV
    with open("data_mahasiswa.csv", "a", newline="") as f:
        writer = csv.writer(f)
        writer.writerow([nama, prodi, nomor, datetime.now().strftime("%d-%m-%Y %H:%M:%S")])

    messagebox.showinfo("Sukses", "Data berhasil disimpan!")
    entry_nama.delete(0, tk.END)
    entry_prodi.delete(0, tk.END)
    entry_nomor.delete(0, tk.END)

# Buat jendela utama
root = tk.Tk()
root.title("Aplikasi Mahasiswa Baru USTI")
root.geometry("400x250")

# Label dan Entry
tk.Label(root, text="Nama Mahasiswa:").pack(pady=5)
entry_nama = tk.Entry(root, width=40)
entry_nama.pack()

tk.Label(root, text="Prodi-Kelas:").pack(pady=5)
entry_prodi = tk.Entry(root, width=40)
entry_prodi.pack()

tk.Label(root, text="Nomor:").pack(pady=5)
entry_nomor = tk.Entry(root, width=40)
entry_nomor.pack()

# Tombol Save
btn_save = tk.Button(root, text="SAVE", command=save_data)
btn_save.pack(pady=20)

# Jalankan aplikasi
root.mainloop()