from tkinter import filedialog
from tkinter import *

window = Tk ()
window.geometry('600x450')
window.title("Gus Health")
window.filename = filedialog.askopenfilename \
    (initialdir="/", title="Select file",
     filetypes=(("xlsx files", "*.xlsx"), ("all files", "*.*")))
print (window.filename)
