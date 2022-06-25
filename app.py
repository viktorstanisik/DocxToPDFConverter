import subprocess
from tkinter import messagebox
from tkinter import *
import docx2pdf
import os, glob
import shutil


def getUrl():
    if part_text.get() == '' or part2_text.get() == '':
        messagebox.showerror('Required Fields', 'Please include all fields')
    else:
        inputDD = part_text.get()
        ouputDD = part2_text.get()
        outstring = convertFiles(inputDD)
        result = checkFiles(outstring)
        if result == 1:
            messagebox.showinfo("Processed", "Your files have been converted")
            subprocess.Popen(os.startfile(os.path.abspath(ouputDD)))
        else:
            messagebox.showwarning("Processed", "There are no word files to convert inside your input directory folder. Please enter different path")
      
def checkFiles(string): 
    os.chdir(string)  
    for i in glob.glob("*.pdf"):    
        if not i:
            return 0
        else:
            return 1
                
def convertFiles(input):
     docx2pdf.convert(input)
     return input
 
# Create window object
app = Tk()

# Input 1
part_text = StringVar()
part_label = Label(app, text='Input Directory', font=('Open Sans SemiBold',  15), pady=20, bg='#485460', foreground='#d2dae2')
part_label.grid(row=0, column=1, sticky=NSEW)
part_entry = Entry(app, textvariable=part_text, font=('Open Sans', 10),foreground='#485460', width=60, highlightthickness=2, highlightbackground="#079992")
part_entry.grid(row=0, column=6)
# Input 2
part2_text = StringVar()
part2_label = Label(app, text='Output Directory', font=(
    'Open Sans SemiBold', 15), bg='#485460', foreground='#d2dae2')
part2_label.grid(row=1, column=1, sticky=NSEW)
part2_entry = Entry(app, textvariable=part2_text, font=('Open Sans', 10),foreground='#485460', width=60, highlightthickness=2, highlightbackground="#079992")
part2_entry.grid(row=1, column=6)

# Buttons
add_btn = Button(app, text='Convert', font=('Open Sans SemiBold',  12),width=10, foreground='#d2dae2', bg='#218c74', command=getUrl)
add_btn.grid(row=2, column=6, pady=20)


app.title('Word to Pdf Converter')
app.geometry('650x200')

app.configure(bg='#485460')

# Start program
app.mainloop()
