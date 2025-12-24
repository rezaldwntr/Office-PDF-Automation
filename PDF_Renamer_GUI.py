import os
import shutil
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

# --- Fungsi Logika ---
def mulai_proses():
    folder_asal = filedialog.askdirectory(title="1. Pilih Folder PDF Asli")
    if not folder_asal: return

    file_excel = filedialog.askopenfilename(title="2. Pilih File Excel", filetypes=[("Excel files", "*.xlsx *.xls")])
    if not file_excel: return

    label_status.config(text="Sedang memproses... Mohon tunggu.", fg="blue")
    root.update()

    try:
        parent_dir = os.path.dirname(folder_asal)
        folder_tujuan_nama = "HASIL_RENAME_" + os.path.basename(folder_asal)
        folder_tujuan = os.path.join(parent_dir, folder_tujuan_nama)

        if not os.path.exists(folder_tujuan):
            os.makedirs(folder_tujuan)

        files = [f for f in os.listdir(folder_asal) if f.lower().endswith('.pdf')]
        files.sort()

        if not files:
             messagebox.showwarning("Kosong", "Tidak ada file PDF di folder tersebut.")
             label_status.config(text="Siap.", fg="black")
             return

        df = pd.read_excel(file_excel, header=None)
        nama_baru_list = df[0].astype(str).tolist()

        limit = min(len(files), len(nama_baru_list))
        count = 0
        errors = 0

        for i in range(limit):
            old_name = files[i]
            new_name = nama_baru_list[i].strip()
            for char in '<>:"/\\|?*': new_name = new_name.replace(char, '')
            if not new_name.lower().endswith('.pdf'): new_name += ".pdf"

            try:
                shutil.copy2(os.path.join(folder_asal, old_name), os.path.join(folder_tujuan, new_name))
                count += 1
            except:
                errors += 1

        messagebox.showinfo("Selesai", f"Berhasil: {count}\nGagal: {errors}\n\nLokasi:\n{folder_tujuan}")
        label_status.config(text="Selesai. Siap lagi.", fg="green")

    except Exception as e:
        messagebox.showerror("Error", str(e))
        label_status.config(text="Error.", fg="red")

def show_info():
    info_text = (
        "INFO PENTING UNTUK PENGGUNA BARU:\n\n"
        "1. Jika muncul layar biru 'Windows protected your PC' saat membuka aplikasi:\n"
        "   - Itu BUKAN virus.\n"
        "   - Itu karena aplikasi ini buatan sendiri (rezaldwntr) dan belum terdaftar di Microsoft.\n"
        "   - Solusi: Klik 'More info' lalu tombol 'Run anyway'.\n\n"
        "2. Pastikan urutan file PDF di folder sudah sesuai abjad agar cocok dengan urutan Excel.\n\n"
        "Created by: rezaldwntr"
    )
    messagebox.showinfo("Tentang Aplikasi & Bantuan", info_text)

# --- Tampilan GUI ---
root = tk.Tk()
root.title("Mass PDF Renamer Tool")
root.geometry("500x320") # Sedikit diperbesar
root.resizable(False, False)

# Judul
lbl_judul = tk.Label(root, text="Aplikasi Rename PDF Massal", font=("Arial", 14, "bold"), pady=10)
lbl_judul.pack()

# Area Utama
frame = tk.Frame(root, padx=20)
frame.pack(fill=tk.BOTH, expand=True)

desc = tk.Label(frame, text="Ubah nama banyak file PDF sekaligus sesuai data Excel.\nAman & Tidak menghapus file asli.", justify=tk.CENTER)
desc.pack(pady=5)

# Tombol Eksekusi
btn_start = tk.Button(frame, text="MULAI PROSES\n(Pilih Folder PDF & Excel)", font=("Arial", 11, "bold"), bg="#28a745", fg="white", height=2, command=mulai_proses)
btn_start.pack(fill=tk.X, pady=15)

# Tombol Info
btn_info = tk.Button(frame, text="Baca: Info Jika Muncul Error / Layar Biru", bg="#ffc107", command=show_info)
btn_info.pack(fill=tk.X, pady=5)

# Status Bar & Copyright
frame_bawah = tk.Frame(root, bd=1, relief=tk.SUNKEN)
frame_bawah.pack(side=tk.BOTTOM, fill=tk.X)

label_status = tk.Label(frame_bawah, text="Siap.", anchor=tk.W, padx=5)
label_status.pack(side=tk.LEFT)

label_copy = tk.Label(frame_bawah, text="© rezaldwntr", fg="grey", padx=5)
label_copy.pack(side=tk.RIGHT)

root.mainloop()
