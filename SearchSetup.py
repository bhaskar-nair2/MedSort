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

        self.root = root

        # Global Values
        self.Fnm = tk.StringVar(root, "data/SEARCH FILE 2019-20.xlsx")
        self.GPAcnt = tk.IntVar(root, 6)
        self.SPAcnt = tk.IntVar(root, 2)
        self.RCcnt = tk.IntVar(root, 1)

        self.ContCol = tk.StringVar(root, "C")
        self.NameCol = tk.StringVar(root, "D")
        self.UnitCol = tk.StringVar(root, "E")
        self.CoyCol = tk.StringVar(root, "F")
        self.RateCol = tk.StringVar(root, "G")
        self.GstCol = tk.StringVar(root, "J")
        self.SuppCol = tk.StringVar(root, "L")
        self.ToCol = tk.StringVar(root, "M")
        self.FromCol = tk.StringVar(root, "N")

        # Label
        tk.Label(mf, text="File Name").grid(
            column=1, row=1, pady=6, sticky='w')
        tk.Label(mf, text="Total GPA").grid(
            column=1, row=2, pady=6, sticky='w')
        tk.Label(mf, text="Total SPA").grid(
            column=3, row=2, pady=6, sticky='w')
        tk.Label(mf, text="Total RC").grid(column=5, row=2, pady=6, sticky='w')

        tk.Label(mf, text="Contract Column").grid(
            column=1, row=3, pady=6, sticky='w')
        tk.Label(mf, text="Name Column").grid(
            column=3, row=3, pady=6, sticky='w')
        tk.Label(mf, text="Unit Column").grid(
            column=5, row=3, pady=6, sticky='w')
        tk.Label(mf, text="Company Column").grid(
            column=1, row=4, pady=6, sticky='w')
        tk.Label(mf, text="Rate Column").grid(
            column=3, row=4, pady=6, sticky='w')
        tk.Label(mf, text="GST Column").grid(
            column=5, row=4, pady=6, sticky='w')
        tk.Label(mf, text="Supplier Column").grid(
            column=1, row=5, pady=6, sticky='w')
        tk.Label(mf, text="To Column (Only RC)").grid(
            column=3, row=5, pady=6, sticky='w')
        tk.Label(mf, text="From Column (Only RC)").grid(
            column=5, row=5, pady=6, sticky='w')

        # components
        self.fname = ttk.Entry(mf, width=18, textvariable=self.Fnm)
        self.fsch = ttk.Button(
            mf, text="Browse", command=lambda: self.getFile(), width=7)
        self.GPA = ttk.Entry(mf, width=6, textvariable=self.GPAcnt)
        self.SPA = ttk.Entry(mf, width=6, textvariable=self.SPAcnt)
        self.RC = ttk.Entry(mf, width=6, textvariable=self.RCcnt)

        self.contField = ttk.Entry(mf, width=6, textvariable=self.ContCol)
        self.nameField = ttk.Entry(mf, width=6, textvariable=self.NameCol)
        self.unitField = ttk.Entry(mf, width=6, textvariable=self.UnitCol)
        self.coyField = ttk.Entry(mf, width=6, textvariable=self.CoyCol)
        self.rateField = ttk.Entry(mf, width=6, textvariable=self.RateCol)
        self.gstField = ttk.Entry(mf, width=6, textvariable=self.GstCol)
        self.supplierField = ttk.Entry(mf, width=6, textvariable=self.SuppCol)
        self.toField = ttk.Entry(mf, width=6, textvariable=self.ToCol)
        self.fromField = ttk.Entry(mf, width=6, textvariable=self.FromCol)

        self.but = ttk.Button(mf, text="Refresh", command=runCommand)
        self.pgbar = ttk.Progressbar(
            mf, orient="horizontal", mode="determinate")
        self.log = Text(mf, width=100, height=15)

        # Design
        self.fname.grid(column=2, row=1, pady=3, columnspan=4, sticky='we')
        self.fsch.grid(column=6, row=1, pady=3, columnspan=1)
        self.GPA.grid(column=2, row=2, pady=3)
        self.SPA.grid(column=4, row=2, pady=3)
        self.RC.grid(column=6, row=2, pady=3)

        self.contField.grid(column=2, row=3, pady=3)
        self.nameField.grid(column=4, row=3, pady=3)
        self.unitField.grid(column=6, row=3, pady=3)
        self.coyField.grid(column=2, row=4, pady=3)
        self.rateField.grid(column=4, row=4, pady=3)
        self.gstField.grid(column=6, row=4, pady=3)
        self.supplierField.grid(column=2, row=5, pady=3)
        self.toField.grid(column=4, row=5, pady=3)
        self.fromField.grid(column=6, row=5, pady=3)

        self.but.grid(column=3, row=6, columnspan=2, sticky='we', pady=3)
        self.pgbar.grid(column=1, row=7, columnspan=6, sticky='we')
        self.log.grid(column=1, row=8, columnspan=6, sticky='we')

    def refresh(self):
        pass

    def get(self):
        return {
            "file": self.Fnm.get(),
            "PASheets": [int(self.GPA.get()),
                         int(self.SPA.get()),
                         int(self.RC.get())],
            "FileCols": [
                self.ContCol.get(),
                self.NameCol.get(),
                self.UnitCol.get(),
                self.CoyCol.get(),
                self.RateCol.get(),
                self.GstCol.get(),
                self.SuppCol.get(),
                self.ToCol.get(),
                self.FromCol.get(),
            ]
        }

    def getFile(self):
        self.fname.delete(0, len(self.Fnm.get()))
        fnm = filedialog.askopenfilename(parent=self.root, filetypes=(("Microsoft Excel", "*.xlsx"),
                                                                      ("All files", "*.*")))
        try:
            self.fname.insert(0, fnm)
        except:  # <- naked except is a bad idea
            self.showError("Open Source File",
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
    app.attributes('-topmost', True)
    app.update()
    app.mainloop()


if __name__ == '__main__':
    main()
