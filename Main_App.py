import os
import shutil
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from pypdf import PdfWriter

# ==============================================================================
# BASE FRAME (Kerangka Dasar untuk Setiap Halaman)
# ==============================================================================
class BasePage(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller

    def show_menu(self):
        """Kembali ke Menu Utama"""
        self.controller.show_frame("MainMenu")

# ==============================================================================
# MENU UTAMA (DASHBOARD)
# ==============================================================================
class MainMenu(BasePage):
    def __init__(self, parent, controller):
        BasePage.__init__(self, parent, controller)
        
        # Judul
        tk.Label(self, text="PDF TOOLS MASTER", font=("Arial", 20, "bold"), fg="#333").pack(pady=(40, 10))
        tk.Label(self, text="Pilih alat yang ingin digunakan:", font=("Arial", 10), fg="#666").pack(pady=(0, 30))

        # Tombol-tombol Menu
        frame_btn = tk.Frame(self)
        frame_btn.pack(fill="both", expand=True, padx=50)

        self.btn_create(frame_btn, "1. Simple Renamer (Urut Excel)", "#17a2b8", 
                        lambda: controller.show_frame("SimpleRenamerPage"))
        
        self.btn_create(frame_btn, "2. Advanced Renamer (Tabel/Edit)", "#28a745", 
                        lambda: controller.show_frame("AdvancedRenamerPage"))
        
        self.btn_create(frame_btn, "3. Gabung PDF (Isi + TTD)", "#ffc107", 
                        lambda: controller.show_frame("MergerPage"))

        # Footer
        tk.Label(self, text="Created by rezaldwntr", font=("Segoe UI", 9, "italic"), fg="gray").pack(side="bottom", pady=20)
        tk.Button(self, text="?", width=3, command=self.show_info).place(x=350, y=10) # Sesuaikan X dengan lebar window

    def btn_create(self, parent, text, color, command):
        fg_color = "white" if color != "#ffc107" else "black"
        btn = tk.Button(parent, text=text, bg=color, fg=fg_color, font=("Arial", 11, "bold"), height=3, command=command)
        btn.pack(fill="x", pady=8)

    def show_info(self):
        msg = ("Aplikasi All-in-One PDF Tools\n\n"
               "1. Simple Renamer: Rename cepat berdasarkan urutan file di folder vs baris Excel.\n"
               "2. Advanced Renamer: Rename dengan tampilan tabel, cari, replace, dan edit manual.\n"
               "3. Gabung PDF: Menggabungkan dokumen dan TTD secara otomatis berdasarkan nama.")
        messagebox.showinfo("Info Aplikasi", msg)

# ==============================================================================
# HALAMAN 1: SIMPLE RENAMER
# ==============================================================================
class SimpleRenamerPage(BasePage):
    def __init__(self, parent, controller):
        BasePage.__init__(self, parent, controller)
        
        # Header + Tombol Kembali
        header = tk.Frame(self)
        header.pack(fill="x", pady=10, padx=10)
        tk.Button(header, text="< Kembali", command=self.show_menu).pack(side="left")
        tk.Label(header, text="Metode 1: Simple Rename", font=("Arial", 14, "bold")).pack(side="left", padx=20)

        tk.Label(self, text="Cocokkan urutan file PDF dengan baris Excel.", fg="gray").pack(pady=5)

        # Konten
        container = tk.Frame(self)
        container.pack(expand=True)

        btn_start = tk.Button(container, text="MULAI PROSES\n(Pilih Folder PDF & Excel)", 
                              font=("Arial", 11, "bold"), bg="#28a745", fg="white", height=3, width=30, 
                              command=self.mulai_proses)
        btn_start.pack(pady=20)

        self.label_status = tk.Label(self, text="Siap.", relief=tk.SUNKEN, anchor=tk.W)
        self.label_status.pack(side=tk.BOTTOM, fill=tk.X)

    def mulai_proses(self):
        folder_asal = filedialog.askdirectory(title="1. Pilih Folder PDF Asli")
        if not folder_asal: return
        file_excel = filedialog.askopenfilename(title="2. Pilih File Excel", filetypes=[("Excel files", "*.xlsx *.xls")])
        if not file_excel: return

        self.label_status.config(text="Memproses...", fg="blue")
        self.update_idletasks()
        
        try:
            parent_dir = os.path.dirname(folder_asal)
            folder_tujuan = os.path.join(parent_dir, "HASIL_RENAME_" + os.path.basename(folder_asal))
            if not os.path.exists(folder_tujuan): os.makedirs(folder_tujuan)

            files = [f for f in os.listdir(folder_asal) if f.lower().endswith('.pdf')]
            files.sort()
            df = pd.read_excel(file_excel, header=None)
            nama_baru_list = df[0].astype(str).tolist()

            limit = min(len(files), len(nama_baru_list))
            count = 0
            for i in range(limit):
                old_name = files[i]
                new_name = nama_baru_list[i].strip()
                for char in '<>:"/\\|?*': new_name = new_name.replace(char, '')
                if not new_name.lower().endswith('.pdf'): new_name += ".pdf"
                try:
                    shutil.copy2(os.path.join(folder_asal, old_name), os.path.join(folder_tujuan, new_name))
                    count += 1
                except: pass

            messagebox.showinfo("Selesai", f"Berhasil: {count} file.\nLokasi: {folder_tujuan}")
            self.label_status.config(text="Selesai.", fg="green")
            
        except Exception as e:
            messagebox.showerror("Error", str(e))

# ==============================================================================
# HALAMAN 2: ADVANCED RENAMER
# ==============================================================================
class AdvancedRenamerPage(BasePage):
    def __init__(self, parent, controller):
        BasePage.__init__(self, parent, controller)
        self.folder_path = ""
        self.file_data = [] 

        # Header
        header = tk.Frame(self)
        header.pack(fill="x", pady=10, padx=10)
        tk.Button(header, text="< Kembali", command=self.show_menu).pack(side="left")
        tk.Label(header, text="Metode 2: Advanced Renamer", font=("Arial", 14, "bold")).pack(side="left", padx=20)

        # Toolbar Folder
        frame_top = tk.Frame(self, pady=5, padx=10)
        frame_top.pack(fill="x")
        tk.Label(frame_top, text="Folder:").pack(side="left")
        self.entry_path = tk.Entry(frame_top, width=40)
        self.entry_path.pack(side="left", padx=5)
        tk.Button(frame_top, text="Pilih", command=self.browse_folder).pack(side="left")
        tk.Button(frame_top, text="Refresh", command=self.load_files, bg="#f0ad4e", fg="white").pack(side="right")

        # Tools
        frame_tool = tk.LabelFrame(self, text="Tools", padx=5, pady=5)
        frame_tool.pack(fill="x", padx=10, pady=5)
        
        tk.Label(frame_tool, text="🔍 Cari:").pack(side="left")
        self.entry_search = tk.Entry(frame_tool, width=12)
        self.entry_search.pack(side="left", padx=5)
        self.entry_search.bind("<KeyRelease>", self.filter_table)

        tk.Label(frame_tool, text="| Ganti:").pack(side="left", padx=5)
        self.entry_find = tk.Entry(frame_tool, width=10)
        self.entry_find.pack(side="left")
        tk.Label(frame_tool, text="->").pack(side="left")
        self.entry_replace = tk.Entry(frame_tool, width=10)
        self.entry_replace.pack(side="left")
        tk.Button(frame_tool, text="Terapkan", command=self.bulk_replace_selected, bg="#5bc0de", font=("Arial", 9)).pack(side="left", padx=5)

        # Tabel
        frame_table = tk.Frame(self)
        frame_table.pack(fill="both", expand=True, padx=10)
        columns = ("no", "original", "new", "status")
        self.tree = ttk.Treeview(frame_table, columns=columns, show="headings", selectmode="extended")
        self.tree.heading("no", text="No"); self.tree.column("no", width=40)
        self.tree.heading("original", text="Nama Asli"); self.tree.column("original", width=250)
        self.tree.heading("new", text="Nama Baru (Edit Klik 2x)"); self.tree.column("new", width=250)
        self.tree.heading("status", text="Status")
        
        scrollbar = ttk.Scrollbar(frame_table, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.tree.pack(side="left", fill="both", expand=True)
        self.tree.bind("<Double-1>", self.on_double_click)

        # Footer
        tk.Button(self, text="EKSEKUSI RENAME", font=("Arial", 11, "bold"), bg="#28a745", fg="white", 
                  command=self.execute_rename).pack(pady=10)

    def browse_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.entry_path.delete(0, tk.END); self.entry_path.insert(0, folder)
            self.folder_path = folder; self.load_files()

    def load_files(self):
        if not self.folder_path: return
        self.file_data = []
        try:
            files = [f for f in os.listdir(self.folder_path) if f.lower().endswith('.pdf')]
            files.sort()
            for f in files: self.file_data.append({'original': f, 'new': f, 'status': 'Pending'})
            self.update_table(self.file_data)
        except Exception as e: messagebox.showerror("Error", str(e))

    def update_table(self, data):
        for item in self.tree.get_children(): self.tree.delete(item)
        for index, item in enumerate(data):
            real_index = self.file_data.index(item)
            tag = "changed" if item['original'] != item['new'] else "normal"
            self.tree.insert("", "end", iid=real_index, values=(index+1, item['original'], item['new'], item['status']), tags=(tag,))
        self.tree.tag_configure("changed", foreground="blue", font=("Arial", 9, "bold"))

    def filter_table(self, event=None):
        query = self.entry_search.get().lower()
        filtered = [x for x in self.file_data if query in x['original'].lower() or query in x['new'].lower()]
        self.update_table(filtered)

    def on_double_click(self, event):
        item = self.tree.identify("item", event.x, event.y)
        col = self.tree.identify_column(event.x)
        if col == "#3" and item:
            x, y, w, h = self.tree.bbox(item, col)
            entry = tk.Entry(self.tree); entry.place(x=x, y=y, width=w, height=h)
            entry.insert(0, self.tree.item(item, "values")[2]); entry.focus()
            def save(e):
                self.file_data[int(item)]['new'] = entry.get()
                self.file_data[int(item)]['status'] = 'Edited'
                entry.destroy(); self.filter_table()
            entry.bind("<Return>", save); entry.bind("<FocusOut>", lambda e: entry.destroy())

    def bulk_replace_selected(self):
        find, replace = self.entry_find.get(), self.entry_replace.get()
        for iid in self.tree.selection():
            self.file_data[int(iid)]['new'] = self.file_data[int(iid)]['new'].replace(find, replace)
            self.file_data[int(iid)]['status'] = 'Bulk Edit'
        self.filter_table()

    def execute_rename(self):
        if messagebox.askyesno("Konfirmasi", "Yakin ubah nama file?"):
            count = 0
            for item in self.file_data:
                if item['original'] != item['new']:
                    try:
                        os.rename(os.path.join(self.folder_path, item['original']), os.path.join(self.folder_path, item['new']))
                        item['original'] = item['new']; item['status'] = "OK"; count += 1
                    except: item['status'] = "Error"
            self.filter_table(); messagebox.showinfo("Info", f"{count} file berhasil di-rename.")

# ==============================================================================
# HALAMAN 3: MERGER
# ==============================================================================
class MergerPage(BasePage):
    def __init__(self, parent, controller):
        BasePage.__init__(self, parent, controller)
        self.input_folder = tk.StringVar()
        self.output_folder = tk.StringVar()
        self.suffix_isi = tk.StringVar(value="_Isi")
        self.suffix_ttd = tk.StringVar(value="_TTD")

        header = tk.Frame(self)
        header.pack(fill="x", pady=10, padx=10)
        tk.Button(header, text="< Kembali", command=self.show_menu).pack(side="left")
        tk.Label(header, text="Metode 3: Gabung PDF", font=("Arial", 14, "bold")).pack(side="left", padx=20)

        # Form
        container = tk.Frame(self, padx=20)
        container.pack(fill="x")

        tk.Label(container, text="Folder Sumber:").pack(anchor="w")
        f1 = tk.Frame(container); f1.pack(fill="x", pady=(0, 10))
        tk.Entry(f1, textvariable=self.input_folder).pack(side="left", fill="x", expand=True)
        tk.Button(f1, text="Pilih", command=lambda: self.input_folder.set(filedialog.askdirectory())).pack(side="right")

        tk.Label(container, text="Folder Hasil:").pack(anchor="w")
        f2 = tk.Frame(container); f2.pack(fill="x", pady=(0, 10))
        tk.Entry(f2, textvariable=self.output_folder).pack(side="left", fill="x", expand=True)
        tk.Button(f2, text="Pilih", command=lambda: self.output_folder.set(filedialog.askdirectory())).pack(side="right")

        f3 = tk.Frame(container); f3.pack(fill="x", pady=5)
        tk.Label(f3, text="Akhiran Isi:").pack(side="left")
        tk.Entry(f3, textvariable=self.suffix_isi, width=10).pack(side="left", padx=5)
        tk.Label(f3, text="Akhiran TTD:").pack(side="left", padx=20)
        tk.Entry(f3, textvariable=self.suffix_ttd, width=10).pack(side="left", padx=5)

        tk.Button(self, text="MULAI GABUNG PDF", bg="#007bff", fg="white", font=("Arial", 11, "bold"), 
                  command=self.process_files).pack(pady=20, fill="x", padx=60)

        self.log_area = scrolledtext.ScrolledText(self, height=10); self.log_area.pack(fill="both", expand=True, padx=20, pady=10)

    def log(self, txt): self.log_area.insert(tk.END, txt + "\n"); self.log_area.see(tk.END)

    def process_files(self):
        in_dir, out_dir = self.input_folder.get(), self.output_folder.get()
        s_isi, s_ttd = self.suffix_isi.get(), self.suffix_ttd.get()
        if not in_dir or not out_dir: return messagebox.showwarning("Warning", "Pilih folder dulu!")
        
        files = [f for f in os.listdir(in_dir) if f.lower().endswith('.pdf')]
        pairs = {}
        for f in files:
            name, kind = "", ""
            if s_isi in f: name, kind = f.split(s_isi)[0], "isi"
            elif s_ttd in f: name, kind = f.split(s_ttd)[0], "ttd"
            else: continue
            if name not in pairs: pairs[name] = {}
            pairs[name][kind] = f

        sukses = 0; self.log("--- Mulai ---")
        for name, p in pairs.items():
            if 'isi' in p and 'ttd' in p:
                try:
                    merger = PdfWriter()
                    merger.append(os.path.join(in_dir, p['isi'])); merger.append(os.path.join(in_dir, p['ttd']))
                    merger.write(os.path.join(out_dir, f"{name}_Lengkap.pdf")); merger.close()
                    self.log(f"✅ Sukses: {name}"); sukses += 1
                except Exception as e: self.log(f"❌ Error {name}: {e}")
            else: self.log(f"⚠️ Skip {name} (Tidak lengkap)")
        messagebox.showinfo("Info", f"Selesai: {sukses} dokumen.")

# ==============================================================================
# MAIN APP CONTROLLER (PENGATUR HALAMAN)
# ==============================================================================
class PDFToolsApp(tk.Tk):
    def __init__(self):
        tk.Tk.__init__(self)
        self.title("All-in-One PDF Tools - rezaldwntr")
        self.geometry("800x600") # Ukuran fix agar muat semua halaman
        self.resizable(False, False)

        # Container Utama (Tempat menaruh halaman)
        self.container = tk.Frame(self)
        self.container.pack(side="top", fill="both", expand=True)
        self.container.grid_rowconfigure(0, weight=1)
        self.container.grid_columnconfigure(0, weight=1)

        self.frames = {}
        # Mendaftarkan semua halaman
        for F in (MainMenu, SimpleRenamerPage, AdvancedRenamerPage, MergerPage):
            page_name = F.__name__
            frame = F(parent=self.container, controller=self)
            self.frames[page_name] = frame
            # Tumpuk semua halaman di tempat yang sama (grid 0,0)
            frame.grid(row=0, column=0, sticky="nsew")

        self.show_frame("MainMenu")

    def show_frame(self, page_name):
        """Menampilkan halaman yang diminta dan menyembunyikan yang lain"""
        frame = self.frames[page_name]
        frame.tkraise()

if __name__ == "__main__":
    app = PDFToolsApp()
    app.mainloop()
