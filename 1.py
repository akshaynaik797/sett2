from tkinter import *

root = Tk()

scrollbar = Scrollbar(root)
scrollbar.pack(side=RIGHT, fill=BOTH)
listbox = Listbox(root)
listbox.pack()


for i in range(100):
    j = "aks, ads,aks, ads,aks, ads,aks, ads, ads "+str(i)
    listbox.insert(END, j)

# bind listbox to scrollbar
listbox.config(yscrollcommand=scrollbar.set, width=20)
scrollbar.config(command=listbox.yview)
mainloop()