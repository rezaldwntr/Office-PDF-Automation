import os
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from pypdf import PdfWriter

class PDFMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Aplikasi Penggabung Perjanjian & TTD")
        self.root.geometry("600x550") # Sedikit dipertinggi agar muat

        # --- Variabel ---
        self.input_folder = tk.StringVar()
        self.output_folder = tk.StringVar()
        self.suffix_isi = tk.StringVar(value="_Isi")
        self.suffix_ttd = tk.StringVar(value="_TTD")

        # --- UI Layout ---
        
        # Header
        lbl_judul = tk.Label(root, text="PDF MERGER TOOL", font=("Arial", 16, "bold"))
        lbl_judul.pack(pady=(15, 10))

        # 1. Pilih Folder Input
        lbl_input = tk.Label(root, text="1. Folder Sumber (File PDF Terpisah):", font=("Arial", 10, "bold"))
        lbl_input.pack(pady=(5, 0), anchor="w", padx=20)
        
        frame_input = tk.Frame(root)
        frame_input.pack(fill="x", padx=20)
        entry_input = tk.Entry(frame_input, textvariable=self.input_folder)
        entry_input.pack(side="left", fill="x", expand=True, padx=(0, 5))
        btn_input = tk.Button(frame_input, text="Pilih Folder", command=self.browse_input)
        btn_input.pack(side="right")

        # 2. Pilih Folder Output
        lbl_output = tk.Label(root, text="2. Folder Hasil (Tempat Simpan):", font=("Arial", 10, "bold"))
        lbl_output.pack(pady=(10, 0), anchor="w", padx=20)
        
        frame_output = tk.Frame(root)
        frame_output.pack(fill="x", padx=20)
        entry_output = tk.Entry(frame_output, textvariable=self.output_folder)
        entry_output.pack(side="left", fill="x", expand=True, padx=(0, 5))
        btn_output = tk.Button(frame_output, text="Pilih Folder", command=self.browse_output)
        btn_output.pack(side="right")

        # 3. Konfigurasi Nama File
        lbl_config = tk.Label(root, text="3. Konfigurasi Penamaan File (Akhiran):", font=("Arial", 10, "bold"))
        lbl_config.pack(pady=(15, 5), anchor="w", padx=20)

        frame_config = tk.Frame(root)
        frame_config.pack(fill="x", padx=20)
        
        tk.Label(frame_config, text="Akhiran File Isi:").pack(side="left")
        tk.Entry(frame_config, textvariable=self.suffix_isi, width=10).pack(side="left", padx=5)
        
        tk.Label(frame_config, text="Akhiran File TTD:").pack(side="left", padx=(20, 0))
        tk.Entry(frame_config, textvariable=self.suffix_ttd, width=10).pack(side="left", padx=5)

        # 4. Tombol Eksekusi
        btn_process = tk.Button(root, text="MULAI PENGGABUNGAN", bg="#007bff", fg="white", font=("Arial", 11, "bold"), height=2, command=self.process_files)
        btn_process.pack(pady=20, fill="x", padx=60)

        # 5. Area Log
        lbl_log = tk.Label(root, text="Log Proses:", font=("Arial", 9))
        lbl_log.pack(anchor="w", padx=20)
        self.log_area = scrolledtext.ScrolledText(root, height=8, state='disabled')
        self.log_area.pack(fill="both", expand=True, padx=20, pady=(0, 5))

        # --- CREDIT FOOTER ---
        lbl_credit = tk.Label(root, text="Created by rezaldwntr", font=("Segoe UI", 8, "italic"), fg="gray")
        lbl_credit.pack(side="bottom", pady=10)

    def browse_input(self):
        folder = filedialog.askdirectory()
        if folder:
            self.input_folder.set(folder)

    def browse_output(self):
        folder = filedialog.askdirectory()
        if folder:
            self.output_folder.set(folder)

    def log(self, message):
        self.log_area.config(state='normal')
        self.log_area.insert(tk.END, message + "\n")
        self.log_area.see(tk.END)
        self.log_area.config(state='disabled')

    def process_files(self):
        in_dir = self.input_folder.get()
        out_dir = self.output_folder.get()
        s_isi = self.suffix_isi.get()
        s_ttd = self.suffix_ttd.get()

        if not in_dir or not out_dir:
            messagebox.showwarning("Peringatan", "Harap pilih folder sumber dan folder hasil dulu!")
            return

        self.log_area.config(state='normal')
        self.log_area.delete(1.0, tk.END)
        self.log_area.config(state='disabled')
        
        self.log(f"--- Memulai Proses by rezaldwntr ---")
        self.log(f"Mencari pasangan: *{s_isi}.pdf dan *{s_ttd}.pdf")

        try:
            files = [f for f in os.listdir(in_dir) if f.lower().endswith('.pdf')]
            data_pasangan = {}

            for f in files:
                nama_bersih = ""
                tipe = ""
                if s_isi in f:
                    nama_bersih = f.split(s_isi)[0]
                    tipe = "isi"
                elif s_ttd in f:
                    nama_bersih = f.split(s_ttd)[0]
                    tipe = "ttd"
                else:
                    continue
                
                if nama_bersih not in data_pasangan:
                    data_pasangan[nama_bersih] = {}
                data_pasangan[nama_bersih][tipe] = f

            sukses = 0
            
            for nama, file_dict in data_pasangan.items():
                if 'isi' in file_dict and 'ttd' in file_dict:
                    path_isi = os.path.join(in_dir, file_dict['isi'])
                    path_ttd = os.path.join(in_dir, file_dict['ttd'])
                    path_out = os.path.join(out_dir, f"{nama}_Lengkap.pdf")

                    try:
                        merger = PdfWriter()
                        merger.append(path_isi)
                        merger.append(path_ttd)
                        merger.write(path_out)
                        merger.close()
                        self.log(f"✅ SUKSES: {nama}")
                        sukses += 1
                    except Exception as e:
                        self.log(f"❌ ERROR {nama}: {e}")
                else:
                    kurang = "TTD" if 'isi' in file_dict else "Isi"
                    self.log(f"⚠️ SKIP: {nama} (Tidak ada file {kurang})")

            messagebox.showinfo("Selesai", f"Proses selesai!\nBerhasil menggabungkan {sukses} dokumen.")
            self.log(f"--- Selesai. Total Sukses: {sukses} ---")

        except Exception as e:
            messagebox.showerror("Error Fatal", str(e))

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFMergerApp(root)
    root.mainloop()