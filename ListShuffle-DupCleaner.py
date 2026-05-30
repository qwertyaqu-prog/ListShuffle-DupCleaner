import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from tkinter import font as tkfont
from openpyxl import Workbook
import math
import random

class ListShufflerTab(ttk.Frame):
    def __init__(self, master):
        super().__init__(master)
        self.app = master.master  # Reference to the App instance
        self.create_widgets()
        
    def update_item_count(self, event=None):
        """Update counter jumlah item secara real-time"""
        try:
            raw_text = self.text_input.get("1.0", tk.END).strip()
            items = [line.strip() for line in raw_text.split('\n') if line.strip()]
            self.count_label.config(text=f"Total Item: {len(items)}")
            
            try:
                num = int(self.entry_columns.get())
                if num > 0:
                    if self.items_per_column_mode.get():
                        items_per_column = num
                        num_columns = math.ceil(len(items) / items_per_column) if len(items) > 0 else 0
                        self.column_info_label.config(text=f"{num_columns} kolom" if num_columns > 0 else "")
                    else:
                        num_columns = num
                        items_per_column = math.ceil(len(items) / num_columns) if len(items) > 0 else 0
                        self.column_info_label.config(text=f"{items_per_column}/kolom" if items_per_column > 0 else "")
                else:
                    self.column_info_label.config(text="")
            except ValueError:
                self.column_info_label.config(text="")
        except Exception as e:
            print(f"Error in update_item_count: {str(e)}")

    def process_list(self):
        try:
            raw_text = self.text_input.get("1.0", tk.END).strip()
            if not raw_text:
                messagebox.showerror("Error", "Silakan paste list Anda terlebih dahulu!")
                return
            
            items = [line.strip() for line in raw_text.split('\n') if line.strip()]
            if not items:
                messagebox.showerror("Error", "Tidak ada item yang valid untuk diproses!")
                return
            
            if self.shuffle_var.get():
                try:
                    random.shuffle(items)
                    self.status_label.config(text="List telah diacak")
                    self.app.style.configure("StatusGreen.TLabel", foreground="#4CAF50")
                    self.status_label.configure(style="StatusGreen.TLabel")
                except Exception as e:
                    messagebox.showerror("Error", f"Gagal mengacak list:\n{str(e)}")
                    return
            
            try:
                num = int(self.entry_columns.get())
                if num < 1:
                    messagebox.showerror("Error", "Nilai harus lebih dari 0!")
                    return
            except ValueError:
                messagebox.showerror("Error", "Harap masukkan angka yang valid!")
                return
            
            if self.items_per_column_mode.get():
                items_per_column = num
                num_columns = math.ceil(len(items) / items_per_column)
            else:
                num_columns = num
                items_per_column = math.ceil(len(items) / num_columns)
            
            wb = Workbook()
            ws = wb.active
            assert ws is not None  # Guarantee ws is not None for type checker
            ws.title = "Data Terbagi"
            
            for col in range(num_columns):
                start_idx = col * items_per_column
                end_idx = start_idx + items_per_column
                for row, item in enumerate(items[start_idx:end_idx], 1):
                    ws.cell(row=row, column=col+1, value=item)
            
            output_file = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")],
                initialfile="List_Teracak.xlsx" if self.shuffle_var.get() else "List_Original.xlsx"
            )
            
            if output_file:
                try:
                    wb.save(output_file)
                    self.show_success_message(output_file, len(items), num_columns, items_per_column)
                except Exception as e:
                    messagebox.showerror("Error", f"Gagal menyimpan file:\n{str(e)}")
                
        except Exception as e:
            messagebox.showerror("Error", f"Terjadi kesalahan sistem:\n{str(e)}")

    def show_success_message(self, filename, item_count, column_count, items_per_column):
        success_msg = f"""
        File Excel berhasil dibuat!
        
        Lokasi: {filename}
        Total Item: {item_count}
        Kolom dibuat: {column_count}
        Item per kolom: {items_per_column}
        Status: {'TERACAK' if self.shuffle_var.get() else 'ORIGINAL'}
        Mode: {'ITEM PER KOLOM' if self.items_per_column_mode.get() else 'JUMLAH KOLOM'}
        """
        messagebox.showinfo("Sukses", success_msg.strip())

    def clear_input(self):
        self.text_input.delete("1.0", tk.END)
        self.status_label.config(text="")
        self.count_label.config(text="Total Item: 0")
        self.column_info_label.config(text="")

    def toggle_mode(self):
        if self.items_per_column_mode.get():
            self.mode_label.config(text="Item per Kolom:")
            self.column_mode_label.config(text="Mode: Item per Kolom")
            self.app.style.configure("ModeBlue.TLabel", foreground="#2196F3")
            self.column_mode_label.configure(style="ModeBlue.TLabel")
        else:
            self.mode_label.config(text="Jumlah Kolom:")
            self.column_mode_label.config(text="Mode: Jumlah Kolom")
            self.app.style.configure("ModeGreen.TLabel", foreground="#4CAF50")
            self.column_mode_label.configure(style="ModeGreen.TLabel")
        self.update_item_count()

    def create_widgets(self):
        # Text Input Area
        input_frame = ttk.Frame(self)
        input_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 5))

        input_header = ttk.Frame(input_frame)
        input_header.pack(fill=tk.X, pady=(0, 5))
        ttk.Label(input_header, text="Paste list Anda (1 item per baris):").pack(side=tk.LEFT)
        self.count_label = ttk.Label(input_header, text="Total Item: 0")
        self.app.style.configure("Primary.TLabel", foreground="#2196F3")
        self.count_label.configure(style="Primary.TLabel")
        self.count_label.pack(side=tk.RIGHT)

        self.text_input = scrolledtext.ScrolledText(
            input_frame, 
            height=8,
            wrap=tk.WORD,
            font=('Segoe UI', 9),
            padx=5,
            pady=5,
            bd=1,
            relief=tk.SOLID,
            highlightthickness=1,
            highlightbackground="#e0e0e0"
        )
        self.text_input.pack(expand=True, fill=tk.BOTH)
        self.text_input.bind("<KeyRelease>", self.update_item_count)

        # Control Panel
        control_panel = ttk.Frame(input_frame)
        control_panel.pack(fill=tk.X, pady=(5, 5))

        # Mode toggle
        mode_row = ttk.Frame(control_panel)
        mode_row.pack(fill=tk.X, pady=2)
        self.items_per_column_mode = tk.BooleanVar(value=False)
        mode_check = ttk.Checkbutton(
            mode_row,
            text="Mode Item/Kolom",
            variable=self.items_per_column_mode,
            command=self.toggle_mode
        )
        mode_check.pack(side=tk.LEFT)

        # Column input
        input_row = ttk.Frame(control_panel)
        input_row.pack(fill=tk.X, pady=2)
        self.mode_label = ttk.Label(input_row, text="Jumlah Kolom:")
        self.mode_label.pack(side=tk.LEFT)

        self.entry_columns = ttk.Entry(input_row, width=6)
        self.entry_columns.pack(side=tk.LEFT, padx=5)
        self.entry_columns.insert(0, "3")

        self.column_info_label = ttk.Label(input_row, text="")
        self.app.style.configure("Info.TLabel", foreground="#757575", font=('Segoe UI', 8))
        self.column_info_label.configure(style="Info.TLabel")
        self.column_info_label.pack(side=tk.LEFT)

        # Shuffle option
        shuffle_row = ttk.Frame(control_panel)
        shuffle_row.pack(fill=tk.X, pady=2)
        self.shuffle_var = tk.BooleanVar(value=False)
        shuffle_check = ttk.Checkbutton(
            shuffle_row,
            text="Acak List",
            variable=self.shuffle_var,
            command=lambda: [
                self.status_label.config(text="Mode pengacakan AKTIF" if self.shuffle_var.get() else ""),
                self.app.style.configure("StatusGreen.TLabel", foreground="#4CAF50"),
                self.status_label.configure(style="StatusGreen.TLabel" if self.shuffle_var.get() else "TLabel")
            ]
        )
        shuffle_check.pack(side=tk.LEFT)

        # Button row
        button_row = ttk.Frame(control_panel)
        button_row.pack(fill=tk.X, pady=(5, 5))
        ttk.Button(button_row, text="Clear", command=self.clear_input, width=10).pack(side=tk.LEFT, padx=2)
        ttk.Button(button_row, text="Copy", command=lambda: [
                  self.app.clipboard_clear(), 
                  self.app.clipboard_append(self.text_input.get("1.0", tk.END)), 
                  self.status_label.config(text="Teks disalin!"),
                  self.app.style.configure("StatusBlue.TLabel", foreground="#2196F3"),
                  self.status_label.configure(style="StatusBlue.TLabel")
                  ], width=10).pack(side=tk.LEFT, padx=2)

        # Mode indicator
        self.column_mode_label = ttk.Label(
            control_panel,
            text="Mode: Jumlah Kolom"
        )
        self.app.style.configure("ModeGreen.TLabel", foreground="#4CAF50", font=('Segoe UI', 8, 'italic'))
        self.column_mode_label.configure(style="ModeGreen.TLabel")
        self.column_mode_label.pack()

        # Process Button
        process_btn = ttk.Button(
            self,
            text="GENERATE EXCEL",
            command=self.process_list,
            style="Bold.TButton",
            width=20
        )
        process_btn.pack(pady=(5, 5))

        # Status Bar
        self.status_label = ttk.Label(
            self,
            text="",
            font=('Segoe UI', 8),
            anchor=tk.CENTER
        )
        self.status_label.pack(fill=tk.X)

        # Bind entry_columns to update item count when changed
        self.entry_columns.bind("<KeyRelease>", self.update_item_count)

class DuplicateCleanerTab(ttk.Frame):
    def __init__(self, master):
        super().__init__(master)
        self.app = master.master  # Reference to the App instance
        self.create_widgets()
        
    def update_results(self):
        input_text = self.text_input.get("1.0", tk.END).strip()
        nomor_list = [n.strip() for n in input_text.split('\n') if n.strip()]
        
        if not nomor_list:
            # Clear all outputs if input is empty
            self.text_output.delete("1.0", tk.END)
            self.text_duplikat.delete("1.0", tk.END)
            self.label_hasil.config(text="Hasil: 0 item")
            self.label_info.config(text="✔ Data sudah unik")
            self.app.style.configure("Success.TLabel", foreground="#388e3c")
            self.label_info.configure(style="Success.TLabel")
            return
        
        # Process duplicates
        nomor_unik = []
        nomor_duplikat = []
        
        for nomor in nomor_list:
            if nomor not in nomor_unik:
                nomor_unik.append(nomor)
            else:
                nomor_duplikat.append(nomor)
        
        total_output = len(nomor_unik)
        
        # Update output
        self.text_output.delete("1.0", tk.END)
        self.text_output.insert("1.0", "\n".join(sorted(nomor_unik)))
        self.label_hasil.config(text=f"Hasil: {total_output} item (Tanpa duplikat)")
        
        self.text_duplikat.delete("1.0", tk.END)
        if nomor_duplikat:
            self.text_duplikat.insert("1.0", "\n".join(sorted(nomor_duplikat)))
            self.label_info.config(text=f"♻ {len(nomor_duplikat)} duplikat dihapus")
            self.app.style.configure("Warning.TLabel", foreground="#d32f2f")
            self.label_info.configure(style="Warning.TLabel")
        else:
            self.text_duplikat.insert("1.0", "✔ Tidak ada duplikat")
            self.label_info.config(text="✔ Data sudah unik")
            self.app.style.configure("Success.TLabel", foreground="#388e3c")
            self.label_info.configure(style="Success.TLabel")

    def update_column1_counter(self, event=None):
        column1_text = self.text_input.get("1.0", tk.END).strip()
        column1_items = [n.strip() for n in column1_text.split('\n') if n.strip()]
        self.label_total_input.config(text=f"Total: {len(column1_items)} item")

    def update_column2_counter(self, event=None):
        column2_text = self.text_column2.get("1.0", tk.END).strip()
        column2_items = [n.strip() for n in column2_text.split('\n') if n.strip()]
        self.label_total_column2.config(text=f"Total: {len(column2_items)} item")

    def remove_from_column1(self):
        column1_text = self.text_input.get("1.0", tk.END).strip()
        column2_text = self.text_column2.get("1.0", tk.END).strip()
        
        column1_items = [n.strip() for n in column1_text.split('\n') if n.strip()]
        column2_items = [n.strip() for n in column2_text.split('\n') if n.strip()]
        
        if not column1_items:
            messagebox.showwarning("Peringatan", "Kolom 1 kosong!")
            return
        
        # Remove items from column1 that exist in column2
        result_items = [item for item in column1_items if item not in column2_items]
        removed_items = [item for item in column1_items if item in column2_items]
        
        # Update clean results
        self.text_output.delete("1.0", tk.END)
        self.text_output.insert("1.0", "\n".join(sorted(result_items)))
        self.label_hasil.config(text=f"Hasil: {len(result_items)} item (Setelah hapus dari Kolom 2)")
        
        # Update duplicates section
        self.text_duplikat.delete("1.0", tk.END)
        if removed_items:
            self.text_duplikat.insert("1.0", "\n".join(sorted(removed_items)))
            self.label_info.config(text=f"♻ {len(removed_items)} item dihapus karena ada di Kolom 2")
            self.app.style.configure("Warning.TLabel", foreground="#d32f2f")
            self.label_info.configure(style="Warning.TLabel")
        else:
            self.text_duplikat.insert("1.0", "✔ Tidak ada item yang dihapus")
            self.label_info.config(text="✔ Tidak ada item yang dihapus")
            self.app.style.configure("Success.TLabel", foreground="#388e3c")
            self.label_info.configure(style="Success.TLabel")

    def clear_all(self):
        self.text_input.delete("1.0", tk.END)
        self.text_column2.delete("1.0", tk.END)
        self.text_output.delete("1.0", tk.END)
        self.text_duplikat.delete("1.0", tk.END)
        self.label_total_input.config(text="Total: 0 item")
        self.label_total_column2.config(text="Total: 0 item")
        self.label_hasil.config(text="Hasil: 0 item")
        self.label_info.config(text="✔ Tidak ada item yang dihapus")
        self.app.style.configure("Success.TLabel", foreground="#388e3c")
        self.label_info.configure(style="Success.TLabel")

    def create_widgets(self):
        # Main Container - Vertical Layout
        main_frame = ttk.Frame(self)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Input Section - Column 1
        input_header = ttk.Frame(main_frame)
        input_header.pack(fill=tk.X)
        self.app.style.configure("Accent.TLabel", foreground="#1976d2", font=('Segoe UI', 9, 'bold'))
        ttk.Label(input_header, text="Data Utama", style="Accent.TLabel").pack(side=tk.LEFT)
        self.label_total_input = ttk.Label(input_header, text="Total: 0 item")
        self.app.style.configure("Info.TLabel", foreground="#616161", font=('Segoe UI', 8))
        self.label_total_input.configure(style="Info.TLabel")
        self.label_total_input.pack(side=tk.RIGHT)

        self.text_input = scrolledtext.ScrolledText(main_frame, height=6, font=('Segoe UI', 9), 
                                         bd=1, relief=tk.SOLID, padx=5, pady=5,
                                         highlightbackground="#e0e0e0", highlightthickness=1)
        self.text_input.pack(fill=tk.BOTH, pady=(0, 5))

        # Input Section - Column 2
        column2_header = ttk.Frame(main_frame)
        column2_header.pack(fill=tk.X)
        self.app.style.configure("WarningRed.TLabel", foreground="#d32f2f", font=('Segoe UI', 9, 'bold'))
        ttk.Label(column2_header, text="Data Pembanding", style="WarningRed.TLabel").pack(side=tk.LEFT)
        self.label_total_column2 = ttk.Label(column2_header, text="Total: 0 item", style="Info.TLabel")
        self.label_total_column2.pack(side=tk.RIGHT)

        self.text_column2 = scrolledtext.ScrolledText(main_frame, height=4, font=('Segoe UI', 9), 
                                           bd=1, relief=tk.SOLID, padx=5, pady=5,
                                           highlightbackground="#e0e0e0", highlightthickness=1)
        self.text_column2.pack(fill=tk.BOTH, pady=(0, 5))

        # Button Frame
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=5)

        # Remove Duplicates Button
        remove_dup_button = ttk.Button(button_frame, text="Hapus Duplikat", 
                                     command=self.update_results)
        self.app.style.configure("Accent.TButton", background="#1976d2", foreground="white")
        remove_dup_button.configure(style="Accent.TButton")
        remove_dup_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 2))

        # Remove from Column 2 Button
        remove_col2_button = ttk.Button(button_frame, text="Hapus Item Kolom 2", 
                                      command=self.remove_from_column1)
        self.app.style.configure("Warning.TButton", background="#d32f2f", foreground="white")
        remove_col2_button.configure(style="Warning.TButton")
        remove_col2_button.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=(2, 0))

        # Clear Button
        clear_button = ttk.Button(main_frame, text="Clear All", 
                                command=self.clear_all)
        self.app.style.configure("Info.TButton", background="#616161", foreground="white")
        clear_button.configure(style="Info.TButton")
        clear_button.pack(fill=tk.X, pady=2)

        # Output Section
        # Clean Results
        clean_header = ttk.Frame(main_frame)
        clean_header.pack(fill=tk.X)
        self.app.style.configure("Success.TLabel", foreground="#388e3c", font=('Segoe UI', 9, 'bold'))
        ttk.Label(clean_header, text="Hasil Bersih", style="Success.TLabel").pack(side=tk.LEFT)
        self.label_hasil = ttk.Label(clean_header, text="Hasil: 0 item", style="Info.TLabel")
        self.label_hasil.pack(side=tk.RIGHT)

        self.text_output = scrolledtext.ScrolledText(main_frame, height=5, font=('Segoe UI', 9), 
                                          bg="#e8f5e9", bd=1, relief=tk.SOLID, padx=5, pady=5,
                                          highlightbackground="#e0e0e0", highlightthickness=1)
        self.text_output.pack(fill=tk.BOTH, pady=(0, 5))

        # Duplicates
        dup_header = ttk.Frame(main_frame)
        dup_header.pack(fill=tk.X)
        ttk.Label(dup_header, text="Item yang Dihapus", style="WarningRed.TLabel").pack(side=tk.LEFT)
        self.label_info = ttk.Label(dup_header, text="✔ Tidak ada item yang dihapus", style="Success.TLabel")
        self.label_info.pack(side=tk.RIGHT)

        self.text_duplikat = scrolledtext.ScrolledText(main_frame, height=4, font=('Segoe UI', 9), 
                                             bg="#ffebee", bd=1, relief=tk.SOLID, padx=5, pady=5,
                                             highlightbackground="#e0e0e0", highlightthickness=1)
        self.text_duplikat.pack(fill=tk.BOTH)

        # Bind events for realtime counters
        self.text_input.bind("<KeyRelease>", self.update_column1_counter)
        self.text_input.bind("<ButtonRelease>", self.update_column1_counter)
        self.text_column2.bind("<KeyRelease>", self.update_column2_counter)
        self.text_column2.bind("<ButtonRelease>", self.update_column2_counter)

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("List Tools")
        self.geometry("360x550")
        self.resizable(False, False)
        self.style = ttk.Style()
        self.configure_style()
        self.create_widgets()
        
    def configure_style(self):
        self.style.theme_use('clam')
        
        # Color scheme
        bg_color = "#f5f5f5"
        primary_color = "#2196F3"
        success_color = "#4CAF50"
        warning_color = "#FF9800"
        danger_color = "#d32f2f"
        text_color = "#333333"
        info_color = "#616161"

        self.style.configure(".", background=bg_color, foreground=text_color)
        self.style.configure("TFrame", background=bg_color)
        self.style.configure("TLabel", background=bg_color, font=('Segoe UI', 9))
        self.style.configure("TButton", font=('Segoe UI', 9), padding=5)
        self.style.configure("Bold.TButton", font=('Segoe UI', 9, 'bold'))
        
        # Label styles
        self.style.configure("Accent.TLabel", foreground=primary_color, font=('Segoe UI', 9, 'bold'))
        self.style.configure("Warning.TLabel", foreground=warning_color)
        self.style.configure("WarningRed.TLabel", foreground=danger_color, font=('Segoe UI', 9, 'bold'))
        self.style.configure("Success.TLabel", foreground=success_color)
        self.style.configure("Info.TLabel", foreground=info_color, font=('Segoe UI', 8))
        self.style.configure("Primary.TLabel", foreground=primary_color)
        
        # Button styles
        self.style.configure("Accent.TButton", background=primary_color, foreground="white")
        self.style.configure("Warning.TButton", background=danger_color, foreground="white")
        self.style.configure("Success.TButton", background=success_color, foreground="white")
        self.style.configure("Info.TButton", background=info_color, foreground="white")
        
        # Special styles for status messages
        self.style.configure("StatusGreen.TLabel", foreground=success_color)
        self.style.configure("StatusBlue.TLabel", foreground=primary_color)
        self.style.configure("ModeGreen.TLabel", foreground=success_color, font=('Segoe UI', 8, 'italic'))
        self.style.configure("ModeBlue.TLabel", foreground=primary_color, font=('Segoe UI', 8, 'italic'))

    def create_widgets(self):
        # Notebook (Tabbed interface)
        notebook = ttk.Notebook(self)
        notebook.pack(fill=tk.BOTH, expand=True)

        # Create tabs
        tab1 = ListShufflerTab(notebook)
        tab2 = DuplicateCleanerTab(notebook)

        notebook.add(tab1, text="Pembagi List")
        notebook.add(tab2, text="Duplikat List")

if __name__ == "__main__":
    app = App()
    app.mainloop()