import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkinter.messagebox import showinfo, showerror, askyesno
import PyPDF2

class PDFApplication:
    def __init__(self):
        self.all = []
        self.path = ""
        self.path1 = ""

        self.window = tk.Tk()
        self.window.title('PDF')
        self.window.geometry('350x380+440+180')
        self.window.resizable(height=False, width=False)

        label_style = ttk.Style()
        label_style.configure('TLabel', foreground='#000000', font=('OCR A Extended', 11))
        entry_style = ttk.Style()
        entry_style.configure('TEntry', font=('Dotum', 15))
        button_style = ttk.Style()
        button_style.configure('TButton', foreground='#000000', font=('DotumChe', 10))

        self.tab_control = ttk.Notebook(self.window)
        self.first_tab = ttk.Frame(self.tab_control)
        self.second_tab = ttk.Frame(self.tab_control)
        self.third_tab = ttk.Frame(self.tab_control)

        self.tab_control.add(self.first_tab, text='PDF merger')
        self.tab_control.add(self.second_tab, text='PDF encrypter')
        self.tab_control.add(self.third_tab, text='PDF decrypter')
        self.tab_control.pack(expand=1, fill="both")

        self.first_canvas = tk.Canvas(self.first_tab, width=350, height=380)
        self.first_canvas.pack()

        self.second_canvas = tk.Canvas(self.second_tab, width=350, height=380)
        self.second_canvas.pack()

        self.third_canvas = tk.Canvas(self.third_tab, width=350, height=380)
        self.third_canvas.pack()

        self.listbox = tk.Listbox(self.window, height=5, width=25, bg="white", font='Courier 11')
        self.first_canvas.create_window(170, 70, window=self.listbox)

        self.add_files_button = tk.Button(self.window, text="Add Files", width=10, font='arial 12 bold',
                                          borderwidth=2, bg='light blue', command=self.add_files)
        self.first_canvas.create_window(170, 150, window=self.add_files_button)

        self.merge_pdf_button = tk.Button(self.window, text="Merge PDF", width=15, font='arial 12 bold',
                                          borderwidth=2, bg='light blue', command=self.merge_pdfs)
        self.first_canvas.create_window(170, 220, window=self.merge_pdf_button)

        self.display_label = tk.Label(self.window, bg="white", text="", width=29)
        self.second_canvas.create_window(170, 30, window=self.display_label)

        self.select_file_button = tk.Button(self.window, text="Select File", width=10, font='arial 10 bold',
                                            borderwidth=2, bg='light blue', command=self.select_encrypt_file)
        self.second_canvas.create_window(170, 60, window=self.select_file_button)

        self.entry_label = tk.Label(self.window,
                                    text="Enter the password you would like to use \n for encrypting this file : ",
                                    width=32, height=4, font=("Times New Roman", 12))
        self.second_canvas.create_window(170, 150, window=self.entry_label)

        self.entry_box = tk.Entry(self.window, bg="light green", width=25, font="arial 12 bold")
        self.second_canvas.create_window(170, 190, window=self.entry_box)

        self.encrypt_file = tk.Button(self.window, text="Encrypt File", width=10, font='arial 10 bold',
                                      borderwidth=2, bg='light blue', command=self.go_for_encrypt)
        self.second_canvas.create_window(170, 225, window=self.encrypt_file)

        self.display_label1 = tk.Label(self.window, bg="white", text="", width=29)
        self.third_canvas.create_window(170, 30, window=self.display_label1)

        self.select_file_button1 = tk.Button(self.window, text="Select File", width=10, font='arial 10 bold',
                                             borderwidth=2, bg='light blue', command=self.select_decrypt_file)
        self.third_canvas.create_window(170, 60, window=self.select_file_button1)

        self.entry_label1 = tk.Label(self.window, text="Enter PDF password : ", width=32, height=4,
                                     font=("Times New Roman", 12))
        self.third_canvas.create_window(170, 150, window=self.entry_label1)

        self.entry_box1 = tk.Entry(self.window, bg="light green", width=25, font="arial 12 bold")
        self.third_canvas.create_window(170, 190, window=self.entry_box1)

        self.decrypt_file = tk.Button(self.window, text="Decrypt", width=10, font='arial 10 bold',
                                      borderwidth=2, bg='light blue', command=self.go_for_decrypt)
        self.third_canvas.create_window(170, 225, window=self.decrypt_file)

        self.window.protocol('WM_DELETE_WINDOW', self.close_window)
        self.window.mainloop()

    def set_empty(self):
        self.listbox.delete(0, tk.END)
        self.display_label.config(text="")
        self.entry_box.delete(0, 'end')
        self.display_label1.config(text="")
        self.entry_box1.delete(0, 'end')
        self.path = ''
        self.all = []
        self.path1 = ''

    def add_files(self):
        try:
            filenames = filedialog.askopenfilenames(initialdir="/", title="Select PDF files",
                                                    filetypes=(("PDF files", "*.pdf"), ("all files", "*.*")))
            file_list = filenames
            pdf_list = [file for file in file_list if file.lower().endswith('.pdf')]
            for path in pdf_list:
                file_name = path.split('/')[-1]
                if file_name.lower().endswith('.pdf'):
                    self.listbox.insert(tk.END, file_name)
            self.all.extend(pdf_list)
        except:
            messagebox.showerror('Permission denied', 'Please Choose Another Folder!')

    def merge_pdfs(self):
        if self.all:
            merged_pdf = PyPDF2.PdfMerger()
            for file in self.all:
                merged_pdf.append(file)
            save_file = filedialog.asksaveasfilename(defaultextension=".pdf",
                                            filetypes=(("PDF files", "*.pdf"), ("all files", "*.*")))
            if save_file.endswith('.pdf'):
                try:
                    merged_pdf.write(save_file)
                    messagebox.showinfo("Task", 'Merged PDF saved successfully!')
                except Exception as e:
                    messagebox.showerror('Task', 'Invalid file name. PDF file was not saved.')
            merged_pdf.close()
            
            self.set_empty()

    def select_encrypt_file(self):
        try:
            en_filename = filedialog.askopenfilenames(initialdir="/", title="Select PDF files",
                                        filetypes=(("PDF files", "*.pdf"), ("all files", "*.*")))
            file_list1 = en_filename
            pdf_list1 = [file for file in file_list1 if file.lower().endswith('.pdf')]
            for self.path in pdf_list1:
                file_name1 = self.path.split('/')[-1]
                if file_name1.lower().endswith('.pdf'):
                    self.display_label.config(text=file_name1)
        except:
            messagebox.showerror('Permission denied', 'Please Choose Another Folder!')

    def go_for_encrypt(self):
        if self.path:
            a = PyPDF2.PdfReader(self.path)
            writer = PyPDF2.PdfWriter()
            page = a.pages
            for i in page:
                writer.add_page(i)
            if self.entry_box.get():
                writer.encrypt(self.entry_box.get())
                save_file1 = filedialog.asksaveasfilename(defaultextension=".pdf",
                                                        filetypes=(("PDF files", "*.pdf"), ("all files", "*.*")))
                with open(save_file1, 'wb') as f:
                    writer.write(f)
                messagebox.showinfo("Status", "File saved Successfully")
                self.set_empty()
            else:
                messagebox.showerror("Passwd Error", 'Entry box should be empty!')
        else:
            messagebox.showwarning("Warning", "Select any File")

    def select_decrypt_file(self):
        try:
            de_filename = filedialog.askopenfilenames(initialdir="/", title="Select PDF files",
                                                    filetypes=(("PDF files", "*.pdf"), ("all files", "*.*")))
            file_list2 = de_filename
            pdf_list2 = [file for file in file_list2 if file.lower().endswith('.pdf')]
            for self.path1 in pdf_list2:
                file_name2 = self.path1.split('/')[-1]
                if file_name2.lower().endswith('.pdf'):
                    self.display_label1.config(text=file_name2)
        except:
            messagebox.showerror('Permission denied', 'Please Choose Another Folder!')

    def go_for_decrypt(self):
        if self.path1:
            a1 = PyPDF2.PdfReader(self.path1)
            writer1 = PyPDF2.PdfWriter()
            page1 = a1.pages
            if self.entry_box1.get():
                if a1.is_encrypted:
                    try:
                        a1.decrypt(self.entry_box1.get())
                        for i in page1:
                            writer1.add_page(i)
                        save_file2 = filedialog.asksaveasfilename(defaultextension=".pdf",
                                                                filetypes=(("PDF files", "*.pdf"), ("all files", "*.*")))
                        try:
                            with open(save_file2, 'wb') as f:
                                writer1.write(f)
                            messagebox.showinfo("Status", "File saved Successfully")
                            self.set_empty()
                        except:
                            pass
                    except:
                        messagebox.showwarning("Warning", "Passward is Wrong!")
                        self.entry_box1.delete(0, 'end')
                else:
                    messagebox.showwarning("Warning", "This File is not an Encrypted File")
                    self.set_empty()
                    
            else:
                messagebox.showerror("Passwd Error", 'Entry box should not be empty!')
        else:
            messagebox.showwarning("Warning", "Select any File")

    def close_window(self):
        if askyesno(title='Close QR Code Generator-Detector', message='Are you sure you want to close the application?'):
            self.window.destroy()

if __name__ == '__main__':
    PDFApplication()
