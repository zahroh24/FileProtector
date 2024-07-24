from tkinter import *
from tkinter import filedialog, messagebox
import os
import docx
import pyAesCrypt
import subprocess

# Initialize the main window
root = Tk()
root.title("Word Page")
root.geometry("600x430+300+100")
root.resizable(False, False)

def browse():
    global filename
    filename = filedialog.askopenfilename(initialdir=os.getcwd(),
                                          title="Select File",
                                          filetype=(('WORD File', '*.docx'), ('all files', '*.*')))
    entry1.delete(0, END)  # Clear the entry widget before inserting new filename
    entry1.insert(END, filename)

def encrypt_file(input_file, output_file, password):
    buffer_size = 64 * 1024  # 64KB buffer size
    try:
        with open(input_file, "rb") as f_in:
            with open(output_file, "wb") as f_out:
                pyAesCrypt.encryptStream(f_in, f_out, password, buffer_size)
        return True
    except Exception as e:
        print(f"Encryption failed: {e}")
        return False

def protect_word(mainfile, protectfile, code):
    try:
        # Save the original document to the target location
        doc = docx.Document(mainfile)
        temp_file = protectfile + ".temp"
        doc.save(temp_file)
        
        # Encrypt the document
        if encrypt_file(temp_file, protectfile, code):
            os.remove(temp_file)  # Remove the temporary unencrypted file
            messagebox.showinfo("Info", "Successfully Protected the Word File!")
        else:
            os.remove(temp_file)  # Remove the temporary unencrypted file in case of failure
            messagebox.showerror("Invalid", "Failed to protect Word file.")
    except Exception as e:
        messagebox.showerror("Invalid", f"Failed to protect Word file: {e}")

def Protect():
    mainfile = source.get()
    protectfile = target.get()
    code = password.get()
    
    if not mainfile or not protectfile or not code:
        messagebox.showerror("Invalid", "All fields are required!")
        return

    if mainfile.endswith('.docx'):
        protect_word(mainfile, protectfile, code)
    else:
        messagebox.showerror("Invalid", "Unsupported file type!")
        return

    source.set("")
    target.set("")
    password.set("")

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
