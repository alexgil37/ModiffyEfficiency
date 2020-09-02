from ModifyEffCell import main as mainModEff
from GetEffCell import main as mainGetEff
from FindRemovableData import main as main3
from totalActivity import main as main4
from AirSampleData import main as mainAirSamples
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
    mainModEff(path.get(), savePath.get())
    completed_popup()

def dummy2():
    mainGetEff(path.get(), savePath.get())
    completed_popup()

def dummy3():
    main3(path.get(), savePath.get())
    completed_popup()

def dummy4():
    main4(path.get(), savePath.get())
    completed_popup()

def dummyAirSamples():
    mainAirSamples(path.get(), savePath.get())
    completed_popup()



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
label2.place(x=0, y=30)

entry =  tk.Entry(root, width=20, text=path)
entry.place(x=67, y=7)
entry2 =  tk.Entry(root, width=20, text=savePath)
entry2.place(x=67, y=32)

button1 = tk.Button(root, image=image, width=20, height=20,  command=select_path)
button1.place(x=190, y=3)
button2 = tk.Button(root, image=image, width=20, height=20,  command=save_path)
button2.place(x=190, y=28)
button2 = tk.Button(root, text="Default", command=default_path)
button2.place(x=218, y=29)

button3 = tk.Button(root, text="Modify Efficiency",  command=dummy1)
button3.place(x=30, y=65)
button3 = tk.Button(root, text="Find Efficiency",  command=dummy2)
button3.place(x=170, y=65)
button3 = tk.Button(root, text="Find Removable Data",  command=dummy3)
button3.place(x=19, y=95)
button3 = tk.Button(root, text="Find Total Activity",  command=dummy4)
button3.place(x=161, y=95)


airSampleButt = tk.Button(root, text="Air Samples Trending Data",  command=dummyAirSamples)
airSampleButt.place(x=30, y=125)


root.mainloop()

