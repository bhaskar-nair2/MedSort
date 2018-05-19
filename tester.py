from tkinter import *
from tkinter import ttk, messagebox

top = Tk()
def hello():
   print(brute.get())

brute=StringVar(top,0)
C2 = ttk.Checkbutton(text="Human Check?", variable=brute, onvalue=0, offvalue=1)
C2.pack()

B1 = ttk.Button(top, text = "Say Hello", command = hello)
B1.pack()

top.mainloop()

#trash = ['tabs', 'inj', 'bottle', 'syp', 'bot', 'bott', 'cap', 'doses','with','ml','mg','in', 'methyl',\
 #        'containing','antibiotic', 'sodium', 'chloride', 'fluoride', 'phosphate','without', \
  #       'chloride', 'ammonium', 'citrate','adrenaline','gluconate','propionate','absorbent',\
   #      'unmedicated','sulphate','eye drops','lactate','disposable','lignocaine']

#x=open('trash.txt','r+')
#for _ in trash:
#   x.write(_+"\n")
