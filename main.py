from tkinter import *
from tkinter import filedialog as fd
from tkinter import messagebox as mb
from tkinter import ttk

import os
from translate_excel import *
global file, new_file, progress_count


def main():
    def insert_text():
        file_name = fd.askopenfilename(
            filetypes=(("Excel files", "*.xlsx"),
                       ("All files", "*.*")),
            initialdir=os.path.dirname(os.path.abspath('main.py')))
        global file, new_file
        file = os.path.abspath(file_name)
        new_file = os.path.join(os.path.dirname(file_name),
                                f'{os.path.splitext(os.path.abspath(file_name))[0]}_en_ru.xlsx')
        if not os.path.isfile(file_name):
            mb.showinfo("Open Excel file",
                        "Excel-file not selected!")
        text.delete(1.0, END)
        text.insert(1.0, str(os.path.abspath(file_name)))

    def translate_to_en_ru():
        def open_1():
            os.startfile(file)

        def open_2():
            os.startfile(new_file)

        to_translate(file, new_file)
        # Results on new window:
        new_window = Toplevel(root)
        new_window.geometry('400x200')
        new_window.resizable(False, False)
        new_window.grid_columnconfigure(0, weight=1)
        new_window.grid_columnconfigure(1, weight=1)
        new_window.grid_columnconfigure(2, weight=1)
        Label(new_window, text='New File:').grid(row=0, column=1)
        text2 = Text(new_window, width=40,  height=4)
        text2.grid(row=1, column=0, columnspan=3)
        text2.insert(1.0, str(os.path.abspath(new_file)))
        Label(new_window, text=f"New translations : {str(get_counter('tr_count'))}").grid(row=5, column=0)
        Label(new_window, text=f"Translation reuse: {str(get_counter('reuse_count'))}").grid(row=6, column=0)
        Button(new_window, text="Open First File", command=open_1,
               width=12, height=2).grid(row=7, column=0)
        Button(new_window, text="Open Result", command=open_2,
               width=12, height=2).grid(row=7, column=1)
        Button(new_window, text="Quit", command=new_window.destroy,
               width=12, height=2).grid(row=7, column=2)

    root = Tk('600x400+200+200')
    root.title('Excel translator')
    root.resizable(False, False)
    root.geometry('400x280')
    root.grid_columnconfigure(0, weight=1)
    root.grid_columnconfigure(1, weight=1)
    root.grid_columnconfigure(2, weight=1)
    about_tool = "This app translates Excel file from Chinese to English and Russian"
    Label(text=about_tool).grid(row=0, columnspan=3)
    Button(text="Open Excel file", command=insert_text,
           width=12, height=2).grid(row=1, column=0)
    Label(text='Your File:').grid(row=2, column=1)
    text = Text(width=36, height=4)
    text.grid(row=3, column=0, columnspan=3)
    Label(text=' ').grid(row=6, columnspan=3)
    Button(text="Translate", command=translate_to_en_ru,
           width=12, height=2).grid(row=7, column=0)
    Button(text="Quit", command=root.destroy,
           width=12, height=2).grid(row=7, column=2)
    Button(text="Clear DataBase", command=drop_db,
           width=12, height=2).grid(row=1, column=2)
    Label(text='Progress:').grid(row=8, column=1)
    ttk.Progressbar(orient='horizontal', mode='indeterminate',
                    length=280 ).grid(row=9, columnspan=3)

    root.mainloop()


if __name__ == '__main__':
    main()