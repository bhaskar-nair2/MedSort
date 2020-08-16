import tkinter as tk
from tkinter import ttk, Menu, messagebox, filedialog
import SearchSetup as ss
import threading

from scrollFrame import VerticalScrolledFrame as scrollFrame
import re
from Searcher import Searcher

from pathlib import Path


class GUI:
    def __init__(self, master, runcommand, AddToList):
        # Initial Configuration
        self.root = master
        self.re = ss
        N, W, E, S = tk.N, tk.W, tk.E, tk.S

        # MenuBar
        self.menubar = Menu(self.root)
        self.root.config(menu=self.menubar)
        self.fileMenu = Menu(self.menubar)
        self.fileMenu.add_command(
            label='Refresh Search List', command=self.refresh)
        self.fileMenu.add_command(label='Exit', command=self.destroy)
        self.menubar.add_cascade(label="File", menu=self.fileMenu)
        self.AddToList = AddToList

        # MainFrame
        self.mainframe = ttk.Frame(
            self.root, padding="10 10 12 12", relief='groove')
        self.mainframe.grid(column=0, row=0, sticky=(N, W, E, S))
        self.mainframe.columnconfigure(0, weight=1)
        self.mainframe.rowconfigure(0, weight=1)

        # ChildFrame
        self.childFrame = scrollFrame(self.root)
        self.childFrame.grid(
            column=5, row=0, columnspan=4, sticky=(N, W, E, S))
        self.childFrame.interior.grid_rowconfigure(0, weight=1)
        self.childFr = {}
        self.item = 0

        # newFrame
        self.newFrame = ttk.Frame(
            self.root, padding="10 10 12 12", relief='raised')
        self.newFrame.grid(column=6, row=7)

        # Global Variables
        self.Fnm = tk.StringVar(
            self.root, "/home/bhaskar/Documents/websort/MedSort/data/INDENT NO 60 SORT.xlsx")
        self.IndentCol = tk.StringVar(self.root, "A")
        self.NameCol = tk.StringVar(self.root, "B")
        self.QuantityCol = tk.IntVar(self.root, "D")

        # Labels
        ttk.Label(self.mainframe, text="Indent File ").grid(
            column=1, row=1, sticky=W)
        ttk.Label(self.mainframe, text="Indent Number Column").grid(
            column=1, row=2, sticky=W)
        ttk.Label(self.mainframe, text="Name Column").grid(
            column=1, row=3, sticky=W)
        ttk.Label(self.mainframe, text="Quantity Column").grid(
            column=1, row=4, sticky=W)

        # Components
        self.IndFile = ttk.Entry(
            self.mainframe, width=70, textvariable=self.Fnm)
        self.fsch = ttk.Button(self.mainframe, text="Browse",
                               command=lambda: self.getFile(), width=7)
        # Make Sure you write IndentCol and not Indentcol
        self.IField = ttk.Entry(self.mainframe, width=7,
                                textvariable=self.IndentCol)
        self.NField = ttk.Entry(self.mainframe, width=7,
                                textvariable=self.NameCol)
        self.QField = ttk.Entry(self.mainframe, width=7,
                                textvariable=self.QuantityCol)

        self.sep = ttk.Separator(self.mainframe, orient='horizontal')
        self.btn = ttk.Button(self.mainframe, text="Start", command=runcommand)
        # self.Chk = ttk.Checkbutton(
        #     self.mainframe, text="Human Check?", onvalue=0, offvalue=1)
        self.text = tk.Text(self.mainframe, width=40, height=10)
        self.pgbar = ttk.Progressbar(
            self.mainframe, orient="horizontal", mode="determinate")

        # Grid Style
        self.IndFile.grid(column=2, row=1, sticky=(W, E),
                          columnspan=2, pady="10")
        self.fsch.grid(column=4, row=1, sticky=(W, E), columnspan=1, pady="10")
        self.IField.grid(column=2, row=2, sticky=(
            W, E), columnspan=1, pady="10")
        self.NField.grid(column=2, row=3, sticky=(W, E), pady="10")
        self.QField.grid(column=2, row=4, sticky=(W, E), pady="10")
        self.sep.grid(column=1, row=5, columnspan=4, sticky='ew', pady="5")
        self.btn.grid(column=2, row=6, sticky=N, columnspan=2, pady="5")
        # self.Chk.grid(column=4, row=6, sticky=E, pady="5")
        self.text.grid(column=1, row=7, sticky=(W, E), columnspan=4, pady="5")
        self.pgbar.grid(column=1, row=8, sticky=(W, E), columnspan=4, pady="5")

        self.lsx = 'red'
        self.color = ['red', 'blue']
        self.co = 0

    def refresh(self):
        if __name__ == '__main__':
            self.re.main()

    def get(self):
        return [self.IndFile.get(), self.IField.get(), self.NField.get(), self.QField.get()]

    def destroy(self):
        self.root.destroy()

    def maketablet(self, bValue, cValue, indVal):
        if not self.lsx == bValue:
            self.co = 1 if self.co == 0 else 0
            self.lsx = bValue

        wraplen = 300
        self.childFr[self.item] = tk.Frame(
            self.childFrame.interior, relief='sunken')
        id = self.item
        self.childFr[self.item].item = [id, bValue, indVal]
        self.childFr[self.item].grid_columnconfigure(2, weight=1)
        self.childFr[self.item].grid_rowconfigure(2, weight=1)
        self.childFr[self.item].grid(
            column=1, sticky='n', padx='2', columnspan=3)
        ttk.Label(self.childFr[self.item], text=(
            cValue).upper(), wraplength=wraplen).grid(row=1, columnspan=3)
        ttk.Label(self.childFr[self.item], text=(bValue + "?").upper(), wraplength=wraplen, foreground=self.color[self.co]).grid(row=3,
                                                                                                                                 columnspan=3)
        ttk.Label(self.childFr[self.item], text="is similar to").grid(
            row=2, columnspan=3)
        ttk.Label(self.childFr[self.item], text="At" + indVal,
                  wraplength=wraplen).grid(row=4, columnspan=3)
        self.childFr[self.item].btn = ttk.Button(
            self.childFr[self.item], text="Yes", command=lambda: self.putVals(id))
        self.childFr[self.item].btn.grid(row=5, columnspan=3)
        ttk.Separator(self.childFr[self.item], orient='horizontal').grid(
            row=6, columnspan=3, sticky='ew')
        self.item += 1

    def putVals(self, id):
        self.id = id
        dets = self.childFr[id].item
        self.AddToList(dets, self.childFr[id].btn)

    def remWidget(self, id):
        list = self.childFr[id].grid_slaves()
        for l in list:
            l.destroy()
            self.childFrame.interior.update()
            self.childFrame.update()
        nm = self.childFr[id].item[1]
        self.childFr[id].item = []
        try:
            if self.childFr[id-1].item[1] == nm:
                self.remWidget(id-1)
        except KeyError:
            pass
        except IndexError:
            pass
        try:
            if self.childFr[id+1].item[1] == nm:
                self.remWidget(id+1)
        except KeyError:
            pass
        except IndexError:
            pass

    def getFile(self):
        self.IndFile.delete(0, len(self.Fnm.get()))
        fnm = filedialog.askopenfilename(filetypes=(("Microsoft Excel", "*.xlsx"),
                                                    ("All files", "*.*")))
        try:
            self.IndFile.insert(0, fnm)
        except:  # <- naked except is a bad idea
            showError("Open Source File",
                      "Failed to read file\n'%s'" % fnm)

    def showError(self, title="Error", message="Some error Occoured"):
        messagebox.showerror(title, message)


class App:
    def __init__(self, master):
        Path("./DB").mkdir(parents=True, exist_ok=True)
        Path("./outputs").mkdir(parents=True, exist_ok=True)
        self.master = master
        self.gui = GUI(self.master, self.runit, self.put)

    def runit(self):
        try:
            self.search = Searcher(self.gui)
            self.thread1 = threading.Thread(target=self.search.orcestrator)
            self.thread1.start()
        except FileNotFoundError:
            messagebox.showinfo('Error', 'File not Found')

    def put(self, dets):
        self.gui.btn['state'] = 'disabled'
        # self.thread2 = threading.Thread(
        #     target=lambda: self.search.addToLst(dets))
        # self.thread2.start()


def main():
    app = tk.Tk()
    gui = App(app)
    app.title("Med Sort")
    app.mainloop()


if __name__ == '__main__':
    main()
