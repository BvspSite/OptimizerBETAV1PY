import os
import sys
from tkinter import Tk, filedialog, messagebox, Button, Label, Menu
from comtypes import client  # Untuk Word to PDF
from pdf2docx import Converter  # Untuk PDF to Word

class XboyOptimizer:
    def __init__(self, root):
        self.root = root
        self.root.title("Xboy Optimizer + Converter")
        self.root.geometry("600x400")
        
        # Tambahkan menu bar
        menubar = Menu(root)
        
        # Menu Optimizer
        optimizer_menu = Menu(menubar, tearoff=0)
        optimizer_menu.add_command(label="Bersihkan Cache", command=self.clean_cache)
        optimizer_menu.add_command(label="Optimasi Startup", command=self.optimize_startup)
        menubar.add_cascade(label="Optimizer", menu=optimizer_menu)
        
        # Menu Converter
        converter_menu = Menu(menubar, tearoff=0)
        converter_menu.add_command(label="Word ke PDF", command=self.word_to_pdf)
        converter_menu.add_command(label="PDF ke Word", command=self.pdf_to_word)
        menubar.add_cascade(label="Konversi Dokumen", menu=converter_menu)
        
        root.config(menu=menubar)

    def clean_cache(self):
        messagebox.showinfo("Info", "Fitur bersihkan cache di sini")

    def optimize_startup(self):
        messagebox.showinfo("Info", "Fitur optimasi startup di sini")

    def word_to_pdf(self):
        """Konversi Word ke PDF"""
        try:
            file_path = filedialog.askopenfilename(
                title="Pilih File Word",
                filetypes=[("Word Documents", "*.docx")]
            )
            if not file_path:
                return

            word = client.CreateObject("Word.Application")
            doc = word.Documents.Open(file_path)
            pdf_path = os.path.splitext(file_path)[0] + ".pdf"
            doc.SaveAs(pdf_path, FileFormat=17)
            doc.Close()
            word.Quit()
            messagebox.showinfo("Sukses", f"File disimpan sebagai:\n{pdf_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Gagal konversi:\n{str(e)}")

    def pdf_to_word(self):
        """Konversi PDF ke Word"""
        try:
            file_path = filedialog.askopenfilename(
                title="Pilih File PDF",
                filetypes=[("PDF Files", "*.pdf")]
            )
            if not file_path:
                return

            output_path = os.path.splitext(file_path)[0] + ".docx"
            cv = Converter(file_path)
            cv.convert(output_path)
            cv.close()
            messagebox.showinfo("Sukses", f"File disimpan sebagai:\n{output_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Gagal konversi:\n{str(e)}")

if __name__ == "__main__":
    root = Tk()
    app = XboyOptimizer(root)
    root.mainloop()