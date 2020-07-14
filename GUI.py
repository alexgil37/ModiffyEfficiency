from ModifyEffCell import main as main1
from GetEffCell import main as main2
import os
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from tkinter.ttk import *
from tkinter import *


def resource_path(relative_path):
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


def select_path():
    global path
    curr_directory = os.getcwd()
    filename = filedialog.askdirectory(initialdir=curr_directory, title="Select Folder")
    path.set(filename)


def default_path():
    curr_directory = os.getcwd()
    curr_directory = os.path.join(curr_directory, "Output")
    savePath.set(curr_directory)


def save_path():
    curr_directory = os.getcwd()
    filename = filedialog.askdirectory(initialdir=curr_directory, title="Select Folder")
    savePath.set(filename)


def completed_popup():
    messagebox.showinfo(title="Completed", message="Finished",)

def Loading():
    entry3 = tk.Entry(root, width=15, text=loading)
    entry3.place(x=100, y=125)

def dummy1():
    main1(path.get(), savePath.get())
    completed_popup()

def dummy2():
    main2(path.get(), savePath.get())
    completed_popup()

def progressbar():
    progressbarWindow = Tk()
    progressbarWindow.title('Loading')


    def bar():
        import time
        progress['value'] = 20
        progressbarWindow.update_idletasks()
        time.sleep(0.5)

        progress['value'] = 40
        progressbarWindow.update_idletasks()
        time.sleep(0.5)

        progress['value'] = 50
        progressbarWindow.update_idletasks()
        time.sleep(0.5)

        progress['value'] = 60
        progressbarWindow.update_idletasks()
        time.sleep(0.5)

        progress['value'] = 80
        progressbarWindow.update_idletasks()
        time.sleep(0.5)

        progress['value'] = 100
        progressbarWindow.update_idletasks()
        time.sleep(0.5)

        progress['value'] = 80
        progressbarWindow.update_idletasks()
        time.sleep(0.5)

        progress['value'] = 60
        progressbarWindow.update_idletasks()
        time.sleep(0.5)

        progress['value'] = 50
        progressbarWindow.update_idletasks()
        time.sleep(0.5)

        progress['value'] = 40
        progressbarWindow.update_idletasks()
        time.sleep(0.5)

        progress['value'] = 20
        progressbarWindow.update_idletasks()
        time.sleep(0.5)
        progress['value'] = 0

    progress.pack(pady=10)

root = Tk()
root.title('Efficiency')
root.geometry('300x200')
image = PhotoImage(file=resource_path("images.png"))

path = StringVar()
savePath = StringVar()
loading = StringVar()
loading.set("Processing...")

label = tk.Label(root, text="Folder Path:")
label.place(x=0, y=5)
label2 = tk.Label(root, text="Save Path:")
label2.place(x=0, y=25)

entry =  tk.Entry(root, width=20, text=path)
entry.place(x=67, y=7)
entry2 =  tk.Entry(root, width=20, text=savePath)
entry2.place(x=67, y=27)

button1 = tk.Button(root, image=image, width=20, height=20,  command=select_path)
button1.place(x=190, y=3)
button2 = tk.Button(root, image=image, width=20, height=20,  command=save_path)
button2.place(x=190, y=28)
button2 = tk.Button(root, text="Default", command=default_path)
button2.place(x=218, y=28)

button3 = tk.Button(root, text="Modify Efficiency",  command=lambda : [progressbar(), dummy1()])
button3.place(x=30, y=65)
button3 = tk.Button(root, text="Find Efficiency",  command=dummy2)
button3.place(x=150, y=65)

progress = Progressbar(root, orient = HORIZONTAL, length = 100, mode = 'indeterminate')

root.mainloop()

