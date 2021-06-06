import tkinter as tk
import tkinter.filedialog
from tkinter import *
from tkinter.ttk import *
import tkinter.messagebox
import os
import sheetmaker

class GUIWindow:

    def pickDirectory(self):
        a = tkinter.filedialog.askdirectory()
        if a:
            self.selectedDirectory = a
            self.runButton['state'] = NORMAL


    def generateSheet(self):
        files = [('Excel Files', '*.xls')]
        if self.selectedDirectory != "":
            if not any(fname.endswith('.csv') for fname in os.listdir(self.selectedDirectory)):
                self.badDirectory()
                return
            self.progbar = Progressbar(self.window, orient=HORIZONTAL, length=100, mode='determinate')
            self.progbar.pack()
            self.csvparser.run(self.selectedDirectory)
            tk.messagebox.showinfo(message="Sheet Generated!")
            self.progbar.pack_forget()
            self.saveButon['state'] = NORMAL
            makeSheet = sheetmaker.sheetMaker(self.csvparser.firstName, self.csvparser.midName, self.csvparser.lastName,
                                              self.csvparser.email, self.csvparser.semester, self.csvparser.phone)
            self.wb = makeSheet.run()
        else:
            tk.messagebox.showerror("Error", message="No Folder Selected.\n Please Click 'Select Folder' to Set "
                                                     "Folder")

    def badDirectory(self):
        tk.messagebox.showerror("Error", message="No .csv files found in selected directory: {}.\nTry selecting a "
                                                 "different folder.".format(self.selectedDirectory))

    def saveSheet(self):
        files = [('Excel Files', '*.xls')]
        b = tkinter.filedialog.asksaveasfile(filetypes=files, defaultextension=files)
        if b and self.wb:
            self.wb.save(b.name)
            tk.messagebox.showinfo(message="File Saved!")

    def startGUI(self, csvparser):
        self.csvparser = csvparser

        self.window = tk.Tk()
        self.window.title("Dean's List V1.0")
        self.window.geometry("640x480")
        self.window.resizable(width=False, height=False)
        self.window.configure(bg='gray19')
        self.selectedDirectory = ""

        self.folderText = tk.Label(self.window, text="Click 'Select Folder' and select the folder where the .csv files "
                                                     "are located.\nClick 'Generate Excel Sheet', then click 'Save "
                                                     "As' to save the generated sheet",
                                   fg='yellow', bg='gray19', font=25, pady=50)

        self.folderButton = tk.Button(self.window, text="Select Folder", command=self.pickDirectory)

        self.runButton = tk.Button(self.window, text="Generate Excel Sheet", command=self.generateSheet, state=DISABLED)

        self.saveButon = tk.Button(self.window, text="Save As", command=self.saveSheet, state=DISABLED)

        self.folderText.pack()
        self.folderButton.pack(pady=25)
        self.runButton.pack()
        self.saveButon.pack(pady=25)

        self.window.mainloop()
