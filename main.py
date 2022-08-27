from tkinter import *
from tkinter import filedialog as fd
from tkinter import messagebox as mb
import os
from translate_excel import *


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
        # Results:
        new_window = Toplevel(root)
        Label(new_window, text='New File:').grid(row=1, column=0)
        text2 = Text(new_window, width=35, height=3)
        text2.grid(row=1, column=1, columnspan=2)
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
    # root.resizable(False, False)
    about_tool = "This app translates Excel file from Chinese to English and Russian"
    Label(text=about_tool).grid(row=1, columnspan=3)
    Button(text="Open Excel file", command=insert_text,
           width=12, height=2).grid(row=2, column=0)
    Label(text='Your File:').grid(row=3, column=0)
    text = Text(width=35, height=3)
    text.grid(row=3, column=1, columnspan=2)
    Button(text="to Translate", command=translate_to_en_ru,
           width=12, height=2).grid(row=4, column=0)
    Button(text="Quit", command=root.destroy,
           width=12, height=2).grid(row=4, column=1)

    root.mainloop()


if __name__ == '__main__':
    main()