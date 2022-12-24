# Import Module
import os

from tkinter import *
from tkinter import filedialog
from tkinter.ttk import Progressbar

from sendingConfirmation import sendingConfirmation

file_paths = []


def open_file_sap():
    file = filedialog.askopenfile(mode='r',
                                  filetypes=[('Excel Files', '*.xlsx')])
    if file:
        file_path_sap = os.path.abspath(file.name)
        btn_open_text_cl.set("Client Selected")
        status.configure(text="SAP Selected! Waiting for HJ...")

        file_paths.append(file_path_sap)


def open_file_hj():
    file = filedialog.askopenfile(mode='r',
                                  filetypes=[('Excel Files', '*.xlsx')])

    if file:
        file_path_hj = os.path.abspath(file.name)
        btn_open_text_hj.set("HJ Selected")
        status.configure(text="All Files Selected!")

        file_paths.append(file_path_hj)


# run button
def func_to_run():
    status.configure(text="Running!")
    progress_bar.start(10)
    sendingConfirmation(file_paths, status, progress_bar)


def main_func():
    # Execute Tkinter
    root.mainloop()


# create root window
root = Tk()

# root window title and dimension
root.title("Sending Confirmation")

# Image settings
image = PhotoImage(file=r"efl3pl.png")
image = image.subsample(3, 3)
Label(root, image=image).grid(row=0,
                              column=0,
                              columnspan=1,
                              rowspan=1,
                              padx=5,
                              pady=5)
# status label settings
status = Label(root, text="Waiting for Files...")
status.grid(row=1, column=1, sticky=W, pady=5)

# Credits settings
creator = Label(root, text="Developed by Zain Zameer")
creator.config(font=("Courier", 6))
creator.grid(row=7, column=0, sticky=W, pady=5)

#  open client file button
btn_open_text_cl = StringVar()
btn_open_text_cl.set("Client File Not Selected!")
btn_open_client = Button(root,
                         textvariable=btn_open_text_cl,
                         command=open_file_sap)
#  open HJ file button
btn_open_text_hj = StringVar()
btn_open_text_hj.set("HJ File Not Selected!")
btn_open_hj = Button(root,
                     textvariable=btn_open_text_hj,
                     command=open_file_hj)

btn_run = Button(root,
                 text="Run",
                 command=func_to_run)

# progress bar
progress_bar = Progressbar(root, orient=HORIZONTAL, length=100,
                           mode='indeterminate')

btn_open_client.grid(row=3, column=0, pady=2, padx=5)
btn_open_hj.grid(row=3, column=1, pady=2, padx=5)
btn_run.grid(row=3, column=2, pady=2, padx=5)
progress_bar.grid(row=4, column=1, pady=2, padx=5)

if __name__ == '__main__':
    main_func()
