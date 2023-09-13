# pyinstaller  PdfToText.py -w --onefile
import tkinter as tk
from tkinter import filedialog
import os
import PyPDF2
import xlsxwriter


class PdfToText:
    def __init__(self):
        self.fr = tk.Tk()
        self.fr.title("PDF to TXT Converter")
        self.fr.geometry("900x300")
        self.path_var = tk.StringVar()
        self.create_widgets()

    def error():
        print("error")

    def create_widgets(self):
        file_input = tk.Label(self.fr, text="PDF to TXT Converter :   " + str(self.path_var), font=("ariel", 15),
                              relief='sunken', width=72,
                              height=2,
                              anchor='w', padx=20, borderwidth=1, bg='white')
        file_input.place(x=30, y=50)

        gen_btn = tk.Button(self.fr, text="Generate", command=self.generate_txt, font=("ariel", 15),
                            width=18, height=2, anchor='w', padx=20, relief='raised', borderwidth=1, bg='green',
                            fg='white')
        gen_btn.place(x=30, y=200)

        browse_btn = tk.Button(self.fr, text="Choose Folder", command=self.choose_folder, font=("ariel", 15),
                               width=18, height=2, anchor='w', padx=20, relief='raised', borderwidth=1)
        browse_btn.place(x=30, y=120)

    def choose_folder(self):
        folder_path = filedialog.askdirectory()
        self.path_var.set(folder_path)

    def generate_txt(self):
        folder_path = self.path_var.get()
        if not os.path.exists(folder_path + '\\Text Files'):
            os.makedirs(os.path.join(folder_path, 'Text Files'))
        new_path = os.path.join(folder_path, 'Text Files')
        if folder_path:
            try:
                self.open_file(folder_path, new_path)
                success_msg = tk.Label(self.fr, text="TXT File Generated Successfully", fg="green", font=("ariel", 15))
                success_msg.place(x=300, y=220)
            except Exception as e:
                failure_msg = tk.Label(self.fr, text="Error: TXT File Generation Failed", fg="red", font=("ariel", 15))
                print(e)
                failure_msg.place(x=300, y=220)
        else:
            failure_msg = tk.Label(self.fr, text="Error: Please Choose a Folder", fg="red", font=("ariel", 15))
            failure_msg.place(x=300, y=220)

    def open_file(self,folder_path, new_path):
        folder_path = os.path.join(folder_path, 'PDF Files')
        file_list = os.listdir(folder_path)
        for filename in file_list:
            filepath = os.path.join(folder_path, filename)
            output_path = os.path.join(new_path, filename.replace(".pdf", ".txt"))
            if os.path.isfile(filepath):
                try:
                    open_pdf = PyPDF2.PdfReader(filepath)
                    pages = len(open_pdf.pages)
                    text = ""
                    for page_num in range(pages):
                        page_obj = open_pdf.pages[page_num]
                        text += page_obj.extract_text()
                    with open(output_path, mode="a", encoding="utf-8") as file1:
                        file1.write(text)
                except Exception as e:
                    print('Error processing file:', filename)
                    print(e)
                    error()
            else:
                print('Invalid file:', filename)
                error()

    def run(self):
        self.fr.mainloop()


if __name__ == "__main__":
    app = PdfToText()
    app.run()
