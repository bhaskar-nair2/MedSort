import tkinter as tk
from tkinter import ttk, Menu, messagebox
import searchSetup as ss
import threading
import sqlite3 as sql
import openpyxl as xl
from scrollFrame import VerticalScrolledFrame as scrollFrame
import re


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
        self.fileMenu.add_command(label='Refresh', command=self.refresh)
        self.fileMenu.add_command(label='Exit', command=self.destroy)
        self.menubar.add_cascade(label="File", menu=self.fileMenu)
        self.AddToList = AddToList

        # MainFrame
        self.mainframe = ttk.Frame(self.root, padding="10 10 12 12", relief='groove')
        self.mainframe.grid(column=0, row=0, sticky=(N, W, E, S))
        self.mainframe.columnconfigure(0, weight=1)
        self.mainframe.rowconfigure(0, weight=1)

        # ChildFrame
        self.childFrame = scrollFrame(self.root)
        self.childFrame.grid(column=5, row=0, columnspan=4, sticky=(N, W, E, S))
        self.childFrame.interior.grid_rowconfigure(0,weight=1)
        self.childFr = {}
        self.item = 0

        # newFrame
        self.newFrame = ttk.Frame(self.root, padding="10 10 12 12", relief='raised')
        self.newFrame.grid(column=6, row=7)

        # Global Variables
        self.Fnm = tk.StringVar(self.root, "Demand.xlsx")
        self.Ncol = tk.StringVar(self.root, "C")
        self.Vcol = tk.StringVar(self.root, "R")
        self.tots = tk.IntVar(self.root, 200)

        # Labels
        ttk.Label(self.mainframe, text="Indent File ").grid(column=1, row=1, sticky=W)
        ttk.Label(self.mainframe, text="Indent: Name Column ").grid(column=1, row=2, sticky=W)
        ttk.Label(self.mainframe, text="Indent: Edit Column ").grid(column=1, row=3, sticky=W)
        ttk.Label(self.mainframe, text="Total Values").grid(column=1, row=4, sticky=W)

        # Components
        self.IndFile = ttk.Entry(self.mainframe, width=70, textvariable=self.Fnm)
        self.NCol = ttk.Entry(self.mainframe, width=7, textvariable=self.Ncol)  # Make Sure you write NCol and not Ncol
        self.VCol = ttk.Entry(self.mainframe, width=7, textvariable=self.Vcol)
        self.tot = ttk.Entry(self.mainframe, width=7, textvariable=self.tots)
        self.sep = ttk.Separator(self.mainframe, orient='horizontal')
        self.btn = ttk.Button(self.mainframe, text="Start", command=runcommand)
        self.Chk = ttk.Checkbutton(self.mainframe, text="Human Check?", onvalue=0, offvalue=1)
        self.text = tk.Text(self.mainframe, width=40, height=10)
        self.pgbar = ttk.Progressbar(self.mainframe, orient="horizontal", mode="determinate")

        # Grid Style
        self.IndFile.grid(column=2, row=1, sticky=(W, E), columnspan=3, pady="10")
        self.NCol.grid(column=2, row=2, sticky=(W, E), columnspan=1, pady="10")
        self.VCol.grid(column=2, row=3, sticky=(W, E), pady="10")
        self.tot.grid(column=2, row=4, sticky=(W, E), pady="10")
        self.sep.grid(column=1, row=5, columnspan=4, sticky='ew', pady="5")
        self.btn.grid(column=2, row=6, sticky=N, columnspan=2, pady="5")
        self.Chk.grid(column=4, row=6, sticky=E, pady="5")
        self.text.grid(column=1, row=7, sticky=(W, E), columnspan=4, pady="5")
        self.pgbar.grid(column=1, row=8, sticky=(W, E), columnspan=4, pady="5")


    def refresh(self):
        if __name__ == '__main__':
            self.re.main()

    def get(self):
        return [self.IndFile.get(), self.NCol.get(), self.VCol.get(), self.tot.get()]

    def destroy(self):
        self.root.destroy()

    def maketablet(self, bValue, cValue, indVal):
        wraplen = 300
        self.childFr[self.item] = tk.Frame(self.childFrame.interior, relief='sunken')
        id = self.item
        self.childFr[self.item].item = [id, bValue, indVal]
        self.childFr[self.item].grid_columnconfigure(2, weight=1)
        self.childFr[self.item].grid_rowconfigure(2, weight=1)
        self.childFr[self.item].grid(column=1, sticky='n', padx='2', columnspan=3)
        ttk.Label(self.childFr[self.item], text=(bValue).upper(), wraplength=wraplen).grid(row=1, columnspan=3)
        ttk.Label(self.childFr[self.item], text="is similar to").grid(row=2, columnspan=3)
        ttk.Label(self.childFr[self.item], text="At"+indVal,wraplength=wraplen).grid(row=4, columnspan=3)
        self.childFr[self.item].btn = ttk.Button(self.childFr[self.item], text="Yes", command=lambda: self.putVals(id))
        self.childFr[self.item].btn.grid(row=5, columnspan=3)
        ttk.Label(self.childFr[self.item], text=(cValue + "?").upper(), wraplength=wraplen).grid(row=3, columnspan=3)
        ttk.Separator(self.childFr[self.item], orient='horizontal').grid(row=6, columnspan=3, sticky='ew')
        self.item += 1

    def putVals(self, id):
        self.id = id
        dets = self.childFr[id].item
        self.AddToList(dets,self.childFr[id].btn)

    def remWidget(self, id):
        list = self.childFr[id].grid_slaves()
        for l in list:
            l.destroy()
            self.childFrame.interior.update()
            self.childFrame.update()




class Searcher:
    def __init__(self, frame, data, prog, but, text, maketablet, remWid):
        self.frame = frame
        self.childFr = {}
        self.item = 0
        self.butt = but
        self.progg = prog
        self.text = text
        self.db = 'medSort'

        self.indFile = xl.load_workbook(data[0])
        self.NCol = data[1]
        self.VCol = data[2]
        self.max = data[3]
        self.progg["maximum"] = self.max
        self.maketablet = maketablet
        self.remWid = remWid
        self.FCount = 0
        self.rCount = 0

        self.Fdata = self.indFile.active
        self.nameslst = self.Fdata[self.NCol][1:]
        self.vallst = self.Fdata[self.VCol][1:]

    def startSearch(self):
        self.text.insert("end", "Process Started...\n")
        tbls = ['rcSearchList', 'paSearchList']
        self.butt['state'] = 'disabled'
        # Start Search
        rcList = self.iniList(self.db, tbls[0])
        paList = self.iniList(self.db, tbls[1])
        lists={0:rcList,1:paList}
        for nameCell, valCell in zip(self.nameslst, self.vallst):
            flag=False
            name = nameCell.value
            if name == None:
                break
            for dets in rcList:
                step = Searcher.isSimilar(name, dets[1])
                if step == 0:  #
                    valCell.value = dets[0]
                    self.text.insert('end', "Found " + name + " at " + dets[0] + "\n\n")
                    self.text.yview_pickplace("end")
                    self.FCount += 1
                    flag = True
                elif step == 1:  # Make Tablet
                    self.maketablet(name, dets[1], dets[0])
                    self.rCount += 1
            if not flag:
                for dets in paList:
                    step = Searcher.isSimilar(name, dets[1])
                    if step == 0:  #
                        valCell.value = dets[0]
                        self.text.insert('end', "Found " + name + " at " + dets[0] + "\n\n")
                        self.text.yview_pickplace("end")
                        self.FCount += 1
                    elif step == 1:  # Make Tablet
                        self.maketablet(name, dets[1], dets[0])
                        self.rCount += 1


            self.progg.step(1)
        # Search Over
        self.butt['state'] = 'enabled'
        self.text.insert('end', '\nFound: ' + str(self.FCount) + " of " + str(self.max))
        self.text.insert('end', '\nReported: ' + str(self.rCount)+" possible similar values")
        self.text.yview_pickplace("end")
        self.indFile.template = False
        while True:
            try:
                self.indFile.save('Demand.xlsx')
                break
            except PermissionError:
                messagebox.showinfo("Permission Error!!", "Please close any instances of the fileopen, and press OK")
        messagebox.showinfo("Process Done", "Computational Search Over!\nPlease check values Manually")
        self.text.insert("end", "\nProcess Over...\n")
        self.text.yview_pickplace("end")

    @staticmethod
    def isSimilar(v1, v2):
        if v1.lower() == v2.lower():
            return 0
        if v1.replace(' ','') == v2.replace(' ',''):
            return 0
        else:
            a = re.findall(r"[\w]+", v1.lower())
            b = re.findall(r"[\w]+", v2.lower())
            a = list(set(a) - set(trash))
            b = list(set(b) - set(trash))
            if a == b:
                return 0
            else:
                for _ in a[1:len(a) - 2]:
                    if _.isalpha():
                        for p in b:
                            if _ == p and len(_) > 5:
                                # print(' '.join(a),' '.join(b))
                                return 1
                                # TODO: Make the function to ask matching seperately
        return 2

    @staticmethod
    def iniList(db, tbl):
        con = sql.connect(db)
        cur = con.cursor()
        que = "select * from " + tbl
        cur.execute(que)
        return cur.fetchall()

    def addToLst(self, dets):
        print(dets[2],dets[1])
        #Add to file
        self.add=threading.Thread(target= lambda: self.SaveVals(dets[1],dets[2]))
        self.add.start()
        #end
        self.text.insert('end',"\n\n"+dets[1]+" Declared at-"+dets[2])
        self.text.yview_pickplace("end")
        self.FCount+=1
        self.remWid(dets[0])

    def SaveVals(self,name,value):
       for namesCell,valCell in zip(self.nameslst, self.vallst):
           if namesCell.value == name:
               valCell.value = value
               self.indFile.template = False
               while True:
                   try:
                       self.indFile.save('Demand.xlsx')
                       break
                   except PermissionError:
                       messagebox.showinfo("Permission Error!!",
                                           "Please close any instances of the fileopen, and press OK")


class App:
    def __init__(self, master):
        self.master = master
        self.gui = GUI(self.master, self.runit, self.putVals)

    def runit(self):
        self.search = Searcher(self.gui.childFrame.interior, self.gui.get(), self.gui.pgbar,
                               self.gui.btn, self.gui.text, self.gui.maketablet, self.gui.remWidget)
        self.thread1 = threading.Thread(target=self.search.startSearch)
        self.thread1.start()

    def putVals(self, dets,but):
        but['state']='disabled'
        self.thread2 = threading.Thread(target=lambda: self.search.addToLst(dets))
        self.thread2.start()


def main():
    app = tk.Tk()
    gui = App(app)
    app.title("Med Sort")
    app.mainloop()


trash = open('trash.txt', 'r').read().split("\n")
if __name__ == '__main__':
    main()
