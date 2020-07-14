from ModifyEffCell import main
import os
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from tkinter import *


def resource_path(relative_path):
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


def select_path():
    global path
    curr_directory = os.getcwd()
    filename = filedialog.askdirectory(initialdir=curr_directory, title="Select Folder")
    path.set(filename)


def save_path():
    curr_directory = os.getcwd()
    filename = filedialog.askdirectory(initialdir=curr_directory, title="Select Folder")


def completed_popup():
    messagebox.showinfo(title="Completed", message="Finished",)
    button3 = tk.Button(root, text="Save", command=save_path)
    button3.place(x=95, y=60)

def dummy():
    main(path.get())
    print("Return Success")
    completed_popup()

root = Tk()
root.title('Efficiency')
root.geometry('210x200')
image = PhotoImage(file=resource_path("images.png"))

path = StringVar()

label = tk.Label(root, text="File Path:")
label.place(x=0, y=5)

entry =  tk.Entry(root, width=20, text=path)
entry.place(x=52, y=7)

button1 = tk.Button(root, image=image, width=20, height=20,  command=select_path)
button1.place(x=180, y=3)

button2 = tk.Button(root, text="GO",  command=dummy)
button2.place(x=100, y=30)


root.mainloop()

