import subprocess
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import os
import win32com.client

# Initialize the main window
root = Tk()
root.title("Excel Page")
root.geometry("600x430+300+100")
root.resizable(False, False)

def browse():
    global filename
    filename = filedialog.askopenfilename(initialdir=os.getcwd(),
                                          title="Select File",
                                          filetypes=(('Excel Files', '*.xlsx'), ('All Files', '*.*')))
    entry1.delete(0, END)  # Clear any previous entry
    entry1.insert(END, filename)

def protect_excel(mainfile, protectfile, code):
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False

        workbook = excel.Workbooks.Open(mainfile)
        workbook.Password = code
        workbook.SaveAs(protectfile)
        workbook.Close()
        excel.Quit()

        messagebox.showinfo("Info", "Successfully Protected the Excel File!")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to protect Excel file: {e}")

def Protect():
    mainfile = entry1.get()
    protectfile = entry2.get()
    code = entry3.get()
    
    if not mainfile or not protectfile or not code:
        messagebox.showerror("Error", "All fields are required!")
        return

    if mainfile.endswith('.xlsx'):
        protect_excel(mainfile, protectfile, code)
    else:
        messagebox.showerror("Error", "Unsupported file type!")
        return

    entry1.delete(0, END)
    entry2.delete(0, END)
    entry3.delete(0, END)
    
# Function to go back to main page
def back_to_main():
    root.destroy()
    subprocess.run(['python', 'main.py'])  # Replace 'main.py' with your main page script name

# Set the icon and top image
image_icon = PhotoImage(file="logop.png")
root.iconphoto(False, image_icon)

Top_image = PhotoImage(file="banner.png")
Label(root, image=Top_image).pack()

frame = Frame(root, width=580, height=290, bd=5, relief=GROOVE)
frame.place(x=10, y=130)

# Source file selection
source = StringVar()
Label(frame, text="Source File:", font="arial 10 bold", fg="black").place(x=30, y=30)
entry1 = Entry(frame, width=30, textvariable=source, font="arial 15", bd=1)
entry1.place(x=150, y=28)

Button_icon = PhotoImage(file="button.png")
Button(frame, image=Button_icon, width=35, height=24, bg="white", command=browse).place(x=500, y=27)

# Target file entry
target = StringVar()
Label(frame, text="Target File:", font="arial 10 bold", fg="black").place(x=30, y=80)
entry2 = Entry(frame, width=30, textvariable=target, font="arial 15", bd=1)
entry2.place(x=150, y=78)

# Password entry
password = StringVar()
Label(frame, text="Set Password:", font="arial 10 bold", fg="black").place(x=15, y=130)
entry3 = Entry(frame, width=30, textvariable=password, font="arial 15", bd=1, show='*')
entry3.place(x=150, y=128)

# Protect button
Protect_button = Button(frame, text="Protect File", compound=LEFT, width=20, fg="black", bg="white", font="arial 14 bold", command=Protect)
Protect_button.place(x=180, y=180)

# Back to Main Page button
Back_button = Button(frame, text="Back to Main Page", font="arial 12 bold", width=20, bg="grey", fg="white", command=back_to_main)
Back_button.place(x=195, y=240)

root.mainloop()
