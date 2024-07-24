from tkinter import *
import subprocess

# Initialize the main window
root = Tk()
root.title("File Protector")
root.geometry("600x430+300+100")
root.resizable(False, False)

def open_word_page():
    root.destroy()
    subprocess.run(['python', 'nyoba.py'])

def open_excel_page():
    root.destroy()
    subprocess.run(['python', 'coba.py'])

def open_pdf_page():
    root.destroy()
    subprocess.run(['python', 'excel.py'])

def open_zip_page():
    root.destroy()
    subprocess.run(['python', 'zip.py'])

# Set the icon and top image
image_icon = PhotoImage(file="logop.png")
root.iconphoto(False, image_icon)

Top_image = PhotoImage(file="banner.png")
Label(root, image=Top_image).pack()

# Create buttons for each page
Button(root, text="Word Page", font="arial 15 bold", width=10, height=3 , command=open_word_page, bg='#3375FF', fg='black').place(x=100, y=150)
Button(root, text="PDF Page", font="arial 15 bold", width=10, height=3, command=open_excel_page, bg='#FF5733', fg='black').place(x=350, y=150)
Button(root, text="Excel Page", font="arial 15 bold", width=10, height=3, command=open_pdf_page, bg='#33FF57', fg='black').place(x=100, y=250)
Button(root, text="ZIP Page", font="arial 15 bold", width=10, height=3, command=open_zip_page, bg='yellow', fg='black').place(x=350, y=250)

root.mainloop()
