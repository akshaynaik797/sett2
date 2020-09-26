from datetime import datetime
from tkinter import Tk, Label, Entry, StringVar, Button, ttk, messagebox
from backend import process_values, accept_values

my_date = ''
mygui = Tk(className='Process Failed Settlements')
mygui.geometry("900x200")

format_label = Label(mygui, text='dd/mm/yyyy hh:mm:ss').grid(row=2, column=1)
format_label2 = Label(mygui, text='dd/mm/yyyy hh:mm:ss').grid(row=2, column=4)

fromtime = StringVar()
fromtime_label = Label(mygui, text="From", width=10).grid(row=1, column=0)
fromtime_entry = Entry(mygui, textvariable=fromtime).grid(row=1, column=1, pady=15)
fromtime.set("01/08/2020 00:00:01")

totime = StringVar()
totime_label = Label(mygui, text="To", width=10).grid(row=1, column=3)
totime_entry = Entry(mygui, textvariable=totime).grid(row=1, column=4)
totime.set("01/09/2020 00:00:01")

insname = StringVar()
ttk.Label(mygui, text="Insurer Name", width=10).grid(column=0, row=4, pady=15, padx=15)
inslist = ttk.Combobox(mygui, width=12, textvariable=insname)
inslist['values'] = ('all', 'aditya', 'apollo', 'bajaj', 'big', 'east_west', 'fgh', 'fhpl', 'Good_health', 'hdfc',
                     'health_heritage', 'health_india', 'health_insurance', 'icici_lombard', 'MDINDIA', 'Medi_Assist',
                     'Medsave', 'Paramount', 'Raksha', 'reliance', 'religare', 'small', 'united', 'Universal_Sompo',
                     'vidal', 'vipul')
inslist.grid(row=4, column=1)
inslist.current(0)


def send_values():
    if fromtime.get() != '' and totime.get() != '' and insname.get() != '':
        result = process_values(fromtime.get(), totime.get(), insname.get())
        if result is True:
            print('Report generated')
            messagebox.showinfo(title='Success', message='Report generated')
        else:
            messagebox.showerror(title='Error', message='Job failed, see logs')
    else:
        messagebox.showerror(title="Error", message='Enter values in all fields')

def process_pdfs():
    if fromtime.get() != '' and totime.get() != '' and insname.get() != '':
        result = accept_values(fromtime.get(), datetime.now().strftime("%d/%m/%Y %H:%M:%S"), insname.get())
        if result is True:
            print('Job finished')
            messagebox.showinfo(title='Success', message='Job finished')
        else:
            messagebox.showerror(title='Error', message='Job failed, see logs')
    else:
        messagebox.showerror(title="Error", message='Enter values in all fields')


loginButton = Button(mygui, text="Check", command=send_values).grid(row=5, column=0)
closeButton = Button(mygui, text="Close", command=mygui.destroy).grid(row=5, column=2)
processButton = Button(mygui, text="Process", command=process_pdfs).grid(row=5, column=1)
mygui.mainloop()