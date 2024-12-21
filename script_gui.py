# script_gui.py
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import script_functions
from threading import Thread

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Wordde Resimleri Adam Edici")

        # Directory selection
        self.dir_label = tk.Label(root, text="Dizini sec:")
        self.dir_label.pack(pady=5)
        
        self.dir_entry = tk.Entry(root, width=50)
        self.dir_entry.pack(pady=5)

        self.dir_button = tk.Button(root, text="Goruntule", command=self.select_directory, bg="light gray")
        self.dir_button.pack(pady=5)

        # Status display
        self.status_text = scrolledtext.ScrolledText(root, width=50, height=10, state='disabled')
        self.status_text.pack(pady=10)

        # Log Renkleri için Tag'leri Tanımla
        self.status_text.tag_configure("black", foreground="black")
        self.status_text.tag_configure("red", foreground="red")
        self.status_text.tag_configure("green", foreground="green")

        # Generate Word button
        self.generate_button = tk.Button(root, text="Wordu Olustur", command=self.generate_word_file, bg="light gray")
        self.generate_button.pack(pady=10)

        # Directory variable
        self.selected_dir = ""

    def select_directory(self):
        directory = filedialog.askdirectory()
        if directory:
            self.selected_dir = directory
            self.dir_entry.delete(0, tk.END)
            self.dir_entry.insert(0, directory)

    def log_status(self, message, color="black"):
        self.status_text.configure(state='normal')
        self.status_text.insert(tk.END, message + "\n", color)
        self.status_text.configure(state='disabled')
        self.status_text.see(tk.END)

    def generate_word_file(self):
        if not self.selected_dir:
            messagebox.showerror("Error", "Lutfen dizin seciniz.")
            return
        
        output_file = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word Documents", "*.docx")], initialfile="word_resimler")
        if not output_file:
            return
        
        def task():
            self.log_status("Word olusturuluyor...")
            try:
                script_functions.create_word(self.selected_dir, output_file, log_callback=self.log_status)
            except Exception as e:
                self.log_status(f"Hata: {e}")
        
        Thread(target=task).start()

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
