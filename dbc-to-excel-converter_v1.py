from tkinter import *
from tkinter import filedialog, messagebox
import re

import ctypes.wintypes
CSIDL_PERSONAL = 5       # My Documents
SHGFP_TYPE_CURRENT = 0   # Get current, not default value

from openpyxl import Workbook
from openpyxl.worksheet.table import Table
import cantools

inp= ""
out= ""

def setTextInput(text):
    e1.delete(0,"end")
    e1.insert(0, text)
    op_name = re.split("/", text)[-1].replace('.dbc', '.xlsx')
    e2.insert(0, op_name)
    return op_name

def resetTextInput():
    e1.delete(0,"end")
    e2.delete(0,"end")

def browseFiles():
    try:
        filename = filedialog.askopenfilename(initialdir="/", title="Select a DBC File",
                                              filetypes=(("DBC files", "*.dbc*"), ("all files", "*.*")))
        op_name = setTextInput(filename)
        buf = ctypes.create_unicode_buffer(ctypes.wintypes.MAX_PATH)
        ctypes.windll.shell32.SHGetFolderPathW(None, CSIDL_PERSONAL, None, SHGFP_TYPE_CURRENT, buf)

        global inp, out
        inp = filename
        out = str(buf.value).replace('\\', '/') + '/' + op_name

    except:
        messagebox.showerror("Error", "Unable to load File")


def convert_to_excel():
    try:
        global inp, out
        out2=out
        temp_str = re.split("/", out2)[-1]
        op_name_from_e2=e2.get()
        if op_name_from_e2!=temp_str:
            out=out2.replace(temp_str, op_name_from_e2)


        db = cantools.database.load_file(inp)

        wb = Workbook()
        ws = wb.create_sheet('messages')

        sheet = wb["Sheet"]
        wb.remove(sheet)

        data1 = []
        for msg in db.messages:
            data1.append([msg.name, msg.frame_id, msg.is_extended_frame, msg.length, msg.comment,
                          re.sub('\[|\]|\'', '', str(msg.senders)), msg.send_type, msg.cycle_time, msg.bus_name])

        ws.append(['name', 'frame_id', 'is_extended_frame', 'length', 'comment',
                   'senders', 'send_type', 'cycle_time', 'bus_name'])
        for row in data1:
            ws.append(row)
        ref_str = "A1" + ":" + "I" + str(len(data1) + 1)
        tab1 = Table(displayName="Table1", ref=ref_str)

        ws.add_table(tab1)

        ws = wb.create_sheet('signals')

        data2 = []
        for msg in db.messages:
            for sig in msg.signals:
                data2.append(
                    [sig.name, sig.start, sig.length, "Motorola" if sig.byte_order == "big_endian" else "Intel",
                     "Unsigned" if sig.is_signed == "FALSE" else "Signed", sig.initial, sig.scale, sig.offset,
                     sig.minimum,
                     sig.maximum, sig.unit, re.sub('\[|\]|\'', '',
                                                   str(sig.choices)[:-3].replace('OrderedDict', '').replace('(',
                                                                                                            '').replace(
                                                       ',', ':').replace('):', '  ')),
                     sig.comment, re.sub('\[|\]|\'', '', str(sig.receivers)), sig.is_multiplexer])

        ws.append(['name', 'start', 'length', 'byte_order', 'is_signed', 'initial', 'scale', 'offset', 'minimum',
                   'maximum', 'unit', 'value_table', 'comment', 'receivers', 'is_multiplexer'])
        for row in data2:
            ws.append(row)
        ref_str = "A1" + ":" + "I" + str(len(data2) + 1)
        tab2 = Table(displayName="Table2", ref=ref_str)

        ws.add_table(tab2)
        wb.save(out)
        resetTextInput()
        messagebox.showinfo("Export Successful", "File saved to "+out)

        print("Export successful")

    except:
        print("failed")
        resetTextInput()
        messagebox.showerror("Error", "Unable to Export!!")

master = Tk()
master.title("DBC to Excel Converter")
master.minsize(width=400, height=200)

l1=Label(master, text="Input File Path").grid(row=0)
l2=Label(master, text="Output File Name").grid(row=1)

e1 = Entry(master, width=30)
e2 = Entry(master, width=30)

button_explore = Button(master, text="Browse File", command=browseFiles).grid(row=0, column=2)

e1.grid(row=0, column=1)
e2.grid(row=1, column=1)

Button(master, text='Convert',command=convert_to_excel).grid(row=3, column=1,pady=4)
Button(master, text='Reset', command=resetTextInput).grid(row=3, column=1, sticky=W,pady=4)

mainloop()