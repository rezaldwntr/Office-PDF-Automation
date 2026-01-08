import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading

class ExcelRenamerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Style PDF Renamer")
        self.root.geometry("900x600")
        
        # Data Penyimpanan
        self.folder_path = ""
        self.file_data = [] # List of dict: {'original': 'name.pdf', 'new': 'name.pdf', 'status': ''}
        self.filtered_indices = [] # Indices of displayed items

        # --- UI HEADER ---
        frame_top = tk.Frame(root, pady=10, padx=10)
        frame_top.pack(fill="x")

        tk.Label(frame_top, text="Folder Target:", font=("Arial", 10, "bold")).pack(side="left")
        self.entry_path = tk.Entry(frame_top, width=40)
        self.entry_path.pack(side="left", padx=5)
        tk.Button(frame_top, text="Pilih Folder", command=self.browse_folder, bg="#ddd").pack(side="left")
        
        tk.Button(frame_top, text="REFRESH / RESET", command=self.load_files, bg="#f0ad4e", fg="white").pack(side="right")

        # --- UI TOOLBAR (FILTER & BULK EDIT) ---
        frame_tool = tk.LabelFrame(root, text="Tools (Seperti Excel)", padx=10, pady=5)
        frame_tool.pack(fill="x", padx=10)

        # Filter
        tk.Label(frame_tool, text="🔍 Filter/Cari:").pack(side="left")
        self.entry_search = tk.Entry(frame_tool, width=20)
        self.entry_search.pack(side="left", padx=5)
        self.entry_search.bind("<KeyRelease>", self.filter_table)

        # Bulk Replace
        tk.Label(frame_tool, text="|  Ganti Teks (Bulk Replace):").pack(side="left", padx=(20, 5))
        self.entry_find = tk.Entry(frame_tool, width=15)
        self.entry_find.pack(side="left")
        tk.Label(frame_tool, text="jadi").pack(side="left", padx=2)
        self.entry_replace = tk.Entry(frame_tool, width=15)
        self.entry_replace.pack(side="left")
        tk.Button(frame_tool, text="Terapkan ke Baris Terpilih", command=self.bulk_replace_selected, bg="#5bc0de", fg="white").pack(side="left", padx=5)

        # --- UI TABEL (TREEVIEW) ---
        frame_table = tk.Frame(root)
        frame_table.pack(fill="both", expand=True, padx=10, pady=10)

        columns = ("no", "original", "new", "status")
        self.tree = ttk.Treeview(frame_table, columns=columns, show="headings", selectmode="extended")
        
        self.tree.heading("no", text="No")
        self.tree.heading("original", text="Nama File Asli (Read Only)")
        self.tree.heading("new", text="Nama Baru (Double Click utk Edit)")
        self.tree.heading("status", text="Status")

        self.tree.column("no", width=50, stretch=False)
        self.tree.column("original", width=300)
        self.tree.column("new", width=300)
        self.tree.column("status", width=100)

        # Scrollbar
        scrollbar = ttk.Scrollbar(frame_table, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.tree.pack(side="left", fill="both", expand=True)

        # Bind Double Click untuk Edit
        self.tree.bind("<Double-1>", self.on_double_click)

        # --- UI FOOTER ---
        frame_footer = tk.Frame(root, pady=10)
        frame_footer.pack(fill="x")
        
        tk.Button(frame_footer, text="EKSEKUSI PERUBAHAN NAMA (RENAME ALL)", font=("Arial", 12, "bold"), bg="#28a745", fg="white", command=self.execute_rename).pack(pady=5)
        
        tk.Label(root, text="Created by rezaldwntr", font=("Segoe UI", 8, "italic"), fg="gray").pack(side="bottom", pady=5)

    def browse_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.entry_path.delete(0, tk.END)
            self.entry_path.insert(0, folder)
            self.folder_path = folder
            self.load_files()

    def load_files(self):
        if not self.folder_path: 
            return
        
        # Clear Data
        self.file_data = []
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        try:
            files = [f for f in os.listdir(self.folder_path) if f.lower().endswith('.pdf')]
            files.sort()
            
            for f in files:
                self.file_data.append({
                    'original': f,
                    'new': f,
                    'status': 'Pending'
                })
            
            self.update_table(self.file_data)
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def update_table(self, data_list):
        # Bersihkan tabel visual
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        self.filtered_indices = []
        
        for index, item in enumerate(data_list):
            # Trik: Simpan index asli di iid agar mudah tracking saat filter
            # Cari index asli di self.file_data
            real_index = self.file_data.index(item)
            self.filtered_indices.append(real_index)
            
            tag = "changed" if item['original'] != item['new'] else "normal"
            
            self.tree.insert("", "end", iid=real_index, values=(
                index + 1,
                item['original'],
                item['new'],
                item['status']
            ), tags=(tag,))
        
        self.tree.tag_configure("changed", foreground="blue", font=("Arial", 9, "bold"))
        self.tree.tag_configure("normal", foreground="black")

    def filter_table(self, event=None):
        query = self.entry_search.get().lower()
        if not query:
            self.update_table(self.file_data)
            return

        filtered = [item for item in self.file_data if query in item['original'].lower() or query in item['new'].lower()]
        self.update_table(filtered)

    def on_double_click(self, event):
        # Fitur Edit Cell Manual
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell": return
        
        column = self.tree.identify_column(event.x)
        if column != "#3": # Hanya kolom 'New Name' (kolom ke-3) yang bisa diedit
            return
            
        selected_iid = self.tree.focus()
        if not selected_iid: return

        # Lokasi Cell
        x, y, w, h = self.tree.bbox(selected_iid, column)
        
        # Ambil nilai saat ini
        current_value = self.tree.item(selected_iid, "values")[2]
        
        # Buat Entry Widget di atas cell
        entry_edit = tk.Entry(self.tree)
        entry_edit.place(x=x, y=y, width=w, height=h)
        entry_edit.insert(0, current_value)
        entry_edit.focus()

        def save_edit(event=None):
            new_text = entry_edit.get()
            # Update data master
            idx = int(selected_iid)
            self.file_data[idx]['new'] = new_text
            self.file_data[idx]['status'] = 'Edited'
            
            entry_edit.destroy()
            # Refresh tampilan saat ini (tetap mempertahankan filter)
            self.filter_table()

        entry_edit.bind("<Return>", save_edit)
        entry_edit.bind("<FocusOut>", lambda e: entry_edit.destroy())

    def bulk_replace_selected(self):
        find_txt = self.entry_find.get()
        replace_txt = self.entry_replace.get()
        
        selected_items = self.tree.selection()
        
        if not selected_items:
            messagebox.showinfo("Info", "Pilih baris dulu (bisa CTRL+Klik atau SHIFT+Klik) untuk diganti.")
            return

        for iid in selected_items:
            idx = int(iid)
            current_new_name = self.file_data[idx]['new']
            
            if find_txt in current_new_name:
                new_name = current_new_name.replace(find_txt, replace_txt)
                self.file_data[idx]['new'] = new_name
                self.file_data[idx]['status'] = 'Bulk Edit'
        
        self.filter_table()

    def execute_rename(self):
        if not self.folder_path: return
        
        confirm = messagebox.askyesno("Konfirmasi", "Apakah Anda yakin ingin mengubah nama file yang telah diedit?")
        if not confirm: return

        error_count = 0
        success_count = 0
        
        for item in self.file_data:
            # Hanya proses yang namanya berubah
            if item['original'] != item['new']:
                old_path = os.path.join(self.folder_path, item['original'])
                new_path = os.path.join(self.folder_path, item['new'])
                
                try:
                    # Cek ekstensi, jaga-jaga user lupa nulis .pdf
                    if not new_path.lower().endswith(".pdf"):
                        new_path += ".pdf"

                    os.rename(old_path, new_path)
                    item['original'] = os.path.basename(new_path) # Update original jadi nama baru
                    item['new'] = os.path.basename(new_path)
                    item['status'] = "OK"
                    success_count += 1
                except Exception as e:
                    item['status'] = "Error"
                    print(e)
                    error_count += 1
        
        self.filter_table()
        msg = f"Selesai! {success_count} berhasil."
        if error_count > 0:
            msg += f"\n{error_count} gagal (mungkin nama file sudah ada/duplikat)."
        messagebox.showinfo("Laporan", msg)

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelRenamerApp(root)
    root.mainloop()