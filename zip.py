import os
import py7zr
from tkinter import *
from tkinter import filedialog, messagebox
import subprocess

root = Tk()
root.title("ZIP Page")
root.geometry("600x430+300+100")
root.resizable(False, False)

def browse():
    foldername = filedialog.askdirectory(initialdir=os.getcwd(),
                                         title="Select Folder")
    entry1.delete(0, END)  # Clear any previous entry
    entry1.insert(END, foldername)

def protect_zip(folder, protectfile, code):
    try:
        with py7zr.SevenZipFile(protectfile, 'w', password=code) as zip_out:
            # Add all files in the folder to the 7z archive
            for root_dir, _, files in os.walk(folder):
                for file in files:
                    file_path = os.path.join(root_dir, file)
                    arcname = os.path.relpath(file_path, folder)
                    zip_out.write(file_path, arcname=arcname)

        messagebox.showinfo("Info", "Successfully Protected the 7z File!")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to protect 7z file: {e}")

def Protect():
    folder = entry1.get()
    protectfile = entry2.get()
    code = entry3.get()

    if not folder or not protectfile or not code:
        messagebox.showerror("Error", "All fields are required!")
        return

    if os.path.isdir(folder):
        protect_zip(folder, protectfile, code)
    else:
        messagebox.showerror("Error", "Unsupported file type!")
        return

    entry1.delete(0, END)
    entry2.delete(0, END)
    entry3.delete(0, END)

def back_to_main():
    root.destroy()
    subprocess.run(['python', 'main.py'])

image_icon = PhotoImage(file="logop.png")
root.iconphoto(False, image_icon)

Top_image = PhotoImage(file="banner.png")
Label(root, image=Top_image).pack()

frame = Frame(root, width=580, height=290, bd=5, relief=GROOVE)
frame.place(x=10, y=130)

source = StringVar()
Label(frame, text="Source Folder:", font="arial 10 bold", fg="black").place(x=30, y=30)
entry1 = Entry(frame, width=30, textvariable=source, font="arial 15", bd=1)
entry1.place(x=150, y=28)

Button_icon = PhotoImage(file="button.png")
Button(frame, image=Button_icon, width=35, height=24, bg="white", command=browse).place(x=500, y=27)

target = StringVar()
Label(frame, text="Target File:", font="arial 10 bold", fg="black").place(x=30, y=80)
entry2 = Entry(frame, width=30, textvariable=target, font="arial 15", bd=1)
entry2.place(x=150, y=78)

password = StringVar()
Label(frame, text="Set Password:", font="arial 10 bold", fg="black").place(x=15, y=130)
entry3 = Entry(frame, width=30, textvariable=password, font="arial 15", bd=1, show='*')
entry3.place(x=150, y=128)

Protect_button = Button(frame, text="Protect File", compound=LEFT, width=20, fg="black", bg="white", font="arial 14 bold", command=Protect)
Protect_button.place(x=180, y=180)

Back_button = Button(frame, text="Back to Main Page", font="arial 12 bold", width=20, bg="grey", fg="white", command=back_to_main)
Back_button.place(x=195, y=240)

root.mainloop()
