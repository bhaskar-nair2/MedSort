import tkinter as tk
from tkinter import ttk, messagebox, Text, filedialog
import sqlite3 as sql
from openpyxl import load_workbook as LoadBook
import threading
from SearchFileMaker import ReDataMaker


# base GUI Class
class GUI:
    def __init__(self, root, runCommand):
        mf = ttk.Frame(root, padding="5 5 5 5")
        mf.grid(column=0, row=0)
        mf.columnconfigure(0, weight=1)
        mf.rowconfigure(0, weight=1)

        # Global Values
        self.Fnm = tk.StringVar(root, "data/SEARCH FILE 2019-20.xlsx")
        self.Ncol = tk.StringVar(root, "D")
        self.Vcol = tk.StringVar(root, "C")
        self.GPAcnt = tk.IntVar(root, 6)
        self.SPAcnt = tk.IntVar(root, 2)
        self.RCcnt = tk.IntVar(root, 1)

        # Label
        tk.Label(mf, text="File Name").grid(column=1, row=1, pady=6)
        tk.Label(mf, text="Name Col").grid(column=1, row=2, pady=6)
        tk.Label(mf, text="Value Col").grid(column=3, row=2, pady=6)
        tk.Label(mf, text="Total GPA").grid(column=1, row=3, pady=6)
        tk.Label(mf, text="Total SPA").grid(column=3, row=3, pady=6)
        tk.Label(mf, text="Total RC").grid(column=5, row=3, pady=6)

        # components
        self.fname = ttk.Entry(mf, width=18, textvariable=self.Fnm)
        self.nmCol = ttk.Entry(mf, width=6, textvariable=self.Ncol)
        self.fsch = ttk.Button(
            mf, text="Browse", command=lambda: self.getFile(), width=7)
        self.valCol = ttk.Entry(mf, width=6, textvariable=self.Vcol)
        self.GPA = ttk.Entry(mf, width=6, textvariable=self.GPAcnt)
        self.SPA = ttk.Entry(mf, width=6, textvariable=self.SPAcnt)
        self.RC = ttk.Entry(mf, width=6, textvariable=self.RCcnt)

        self.but = ttk.Button(mf, text="Refresh", command=runCommand)
        self.pgbar = ttk.Progressbar(
            mf, orient="horizontal", mode="determinate")
        self.log = Text(mf, width=40, height=10)

        # Design
        self.fname.grid(column=2, row=1, pady=3, columnspan=4, sticky='we')
        self.fsch.grid(column=6, row=1, pady=3, columnspan=1)
        self.nmCol.grid(column=2, row=2, pady=3)
        self.valCol.grid(column=4, row=2, pady=3)
        self.GPA.grid(column=2, row=3, pady=3)
        self.SPA.grid(column=4, row=3, pady=3)
        self.RC.grid(column=6, row=3, pady=3)

        self.but.grid(column=3, row=4, columnspan=2, sticky='we', pady=3)
        self.pgbar.grid(column=1, row=5, columnspan=6, sticky='we')
        self.log.grid(column=1, row=6, columnspan=6, sticky='we')

    def refresh(self):
        pass

    def get(self):
        return [self.Fnm.get(),
                self.Ncol.get(),
                self.Vcol.get(),
                [int(self.GPA.get()),
                 int(self.SPA.get()),
                 int(self.RC.get())]
                ]

    def getFile(self):
        self.fname.delete(0, len(self.Fnm.get()))
        fnm = filedialog.askopenfilename(filetypes=(("Microsoft Excel", "*.xlsx"),
                                                    ("All files", "*.*")))
        try:
            self.fname.insert(0, fnm)
        except:  # <- naked except is a bad idea
            self.showerror("Open Source File",
                                 "Failed to read file\n'%s'" % fnm)

    def showError(self, title="Error", message="Some error Occoured"):
        messagebox.showerror(title, message)



# Base Application Class
class App:
    def __init__(self, master):
        self.master = master
        self.gui = GUI(self.master, self.runit)

    def runit(self):
        self.search = ReDataMaker(self.gui)
        self.thread1 = threading.Thread(target=self.search.refresh)
        self.thread1.start()


def main():
    app = tk.Tk()
    gui = App(app)
    app.title("Refresh Search File")
    app.mainloop()


if __name__ == '__main__':
    main()
