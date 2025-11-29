import tkinter as tk
from tkinter import messagebox as mb
from tkinter import filedialog
import PyInstaller.__main__
import os
class MyGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.geometry("800x400")
        self.root.title("Transform any pyPrograms to an executable")
        self.root.resizable(False, False)
        
        self.FileSelected = tk.Label(self.root, text="", font=('Arial', 15))
        self.FileSelected.place(x=120, y=135)
        self.label1 = tk.Label(self.root, text="python File:", font=('Arial', 15))
        self.label1.place(x=100, y=100)
        self.label2 = tk.Label(self.root, text="Console ?", font=('Arial', 15))
        self.label2.place(x=300, y=100)
        self.label3 = tk.Label(self.root, text="Name:", font=('Arial', 15))
        self.label3.place(x=500, y=100)
        self.filepath = None
        def choosefile():
            self.filepath = filedialog.askopenfilename()
            filename = self.filepath.split("/")[-1]
            self.FileSelected.config(text=f"{filename}")
            print(filename)
            print(self.filepath)
            print('--console' if self.console_state.get() else '--noconsole')
        self.filechoose = tk.Button(self.root, text="choose file", command=choosefile, font=('Arial', 15))
        self.filechoose.place(x=96, y=170)
        
        self.console_state = tk.IntVar()
        
        self.consolecheck = tk.Checkbutton(self.root, text="active ?", variable=self.console_state, font=('Arial', 15))
        self.consolecheck.place(x=296, y=170)
        
        self.nameText = tk.Text(self.root, width=20, height=1, font=('Arial', 18))
        self.nameText.place(x=470, y=170)
        
        
        def turnexec():
         try:
            if not self.filepath:
                mb.showinfo(title='Messages', message="no file selected")
                return
            safe_path = self.filepath.replace('\\', '/')
            console = '--console' if self.console_state.get() else '--noconsole'
            
            dirname = os.path.dirname(safe_path)
            fileName = self.nameText.get('1.0', tk.END).strip()
            
            if not fileName:
                mb.showinfo(title='Messages', message="Please enter a name")
                return
            
            distdir = os.path.join(dirname, "dist")
            workdir = os.path.join(dirname, "build")
            specdir = dirname
            try:
                PyInstaller.__main__.run([safe_path, '--onefile', console, f'--name={fileName}', f"--distpath={distdir}", f"--workpath={workdir}", f"--specpath={specdir}"])
                mb.showinfo(title='turning file into exec finished', message="file created")
            except Exception as e:
                mb.showinfo(title='turning file into exec finished', message="an error happened")
                print(e)
         except Exception as ee:
             mb.showinfo(title='turning file into exec finished', message=f"an error happened {ee}")
          
        self.createbtn = tk.Button(self.root, text="Create file", command=turnexec, font=('Arial', 15))
        
        self.createbtn.place(x=400, y=250)
        self.root.mainloop()

        
        
    def show_message(self):
        if self.check_state.get() == 0:
            print(self.textbox.get('1.0', tk.END))
        else:
            mb.showinfo(title='Messages', message=self.textbox.get('1.0', tk.END))
        #print(self.check_state.get())
MyGUI()