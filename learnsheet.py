#!/usr/bin/env python
#
# Python GUI to test memorization from an Excel Spreadsheet
#
# DJS Apr 2019
#

import Tkinter as tk

from tkinter import filedialog
from tkinter import *
import openpyxl
import random
import copy



class OptionMenu(tk.OptionMenu):
    """
        Extend the tkinter Options Menu to add the addOption method which doesn't seem to be present
    """

    def __init__(self, *args, **kw):
        self._command = kw.get("command")
        self.variable = args[1]
        tk.OptionMenu.__init__(self, *args, **kw)
    def addOption(self, label):
        self["menu"].add_command(label=label,
            command=tk._setit(self.variable, label, self._command))

class Item:
    def __init__(self,row,column,value):
        self.row = row
        self.column = column
        self.value = value

    def get(self):
        return self.row,self.column, self.value


class GUI(Frame):
    def __init__(self, master):
        Frame.__init__(self, master=None)

        #self.protocol("WM_DELETE_WINDOW", master.destroy)

        self.itemlist = list()

        self.master = master
        self.master.geometry("500x400+300+300")

        self.destination = tk.StringVar(self.master, value='')
        self.labeldestination = Label(self.master, text="Workbook:")
        self.labeldestination.place(x=20, y=20)
        self.entrydistinationdir = Entry(self.master, textvariable=self.destination, state='readonly')
        self.entrydistinationdir.place(x=150, y=18)
        self.btn_browse = Button(self.master, text="Browse", command=self.find_file)
        self.btn_browse.place(x=350, y=18)


        self.labelmailfolder = Label(self.master, text="Sheet List:")
        self.labelmailfolder.place(x=20, y=170)
        self.selectedfolder = tk.StringVar(self.master)
        self.selectedfolder.set("<Choose Sheet>")
        self.optionMenu = OptionMenu(self.master, self.selectedfolder, '<Choose Sheet>')
        self.optionMenu.place(x=148, y=168)
        self.load_button = Button(self.master, text="Load",state=DISABLED, command=self.load_worksheet)
        self.load_button.place(x=350, y=168)


        self.test_button = Button(self.master, text="Test",state=DISABLED, command=self.tb_test)
        self.test_button.place(x=150, y=250)


        self.loadstatus = Label(self.master, text="")
        self.loadstatus.place(x=20, y=300)







    def tb_test(self):
        d = TestWindow(root, self.itemlist, self.max_rows, self.max_cols)
        root.wait_window(d.top)
        # self.valor.set(d.ejemplo)

    def update_sheetlist(self):
        """
            Method to get the list of folders and populate the options Menu
        """
        menu = self.optionMenu.children["menu"]

        menu.delete(0, 'end')
        selectone = False
        for sheet in self.workbook.worksheets:
            self.optionMenu.addOption(sheet.title)
            if (selectone == False):
                self.selectedfolder.set(sheet.title)
                selectone = True

        self.load_button.configure(state=NORMAL)

    def find_file(self):
        """
            Method to open a Directory selection dialog
        """

        file_path = filedialog.askopenfilename(title="Select File")
        self.destination.set(file_path)

        print "Reading workbook...."
        self.workbook = openpyxl.load_workbook(file_path)
        print "Done Reading workbook...."
        self.update_sheetlist()


    def load_worksheet(self):

        self.itemlist = list()
        print "Finding Sheet...."
        sheet = self.workbook.get_sheet_by_name(self.selectedfolder.get())
        print "Found Sheet...."

        print "Calculating Total Rows...."
        print "Total Rows: %s " % sheet.max_row

        self.max_rows = sheet.max_row

        print "Calculating Total Cols...."
        print "Total Cols: %s " % sheet.max_column

        self.max_cols = sheet.max_column

        row_count = 0
        print "Loading List...."
        for row_num in range(2,self.max_rows+1):
            for col_num in range(2, self.max_cols + 1):
                row_heading = sheet.cell(row=row_num, column=1).value
                col_heading = sheet.cell(row=1, column=col_num).value
                item_value = sheet.cell(row=row_num, column=col_num).value
                self.itemlist.append(Item(row_heading, col_heading, item_value))

        self.loadstatus.configure(text="Loaded: %s items" % (len(self.itemlist)))
        self.test_button.configure(state=NORMAL)





class TestWindow:
    def __init__(self, parent, itemlist,rows,cols):
        self.safelist = copy.copy(itemlist)
        self.itemlist = copy.copy(itemlist)
        self.missedlist = list()
        random.shuffle(self.itemlist)

        self.itemindex = 0
        self.misscount = 0
        self.totalmisscount = 0
        self.hitcount = 0
        self.totalhitcount = 0
        self.missedit = False
        self.review = False
        self.listcount = len(self.itemlist)


        self.top = Toplevel(parent)
        self.top.transient(parent)
        self.top.grab_set()
        self.top.title("Test")
        self.top.geometry("500x400+300+300")
        self.labelitem = Label(self.top, text="")
        self.labelitem.place(x=150, y=50)
        #self.labelitem.pack()


        self.entry = tk.StringVar(self.top, value='')

        self.entryfield = Entry(self.top, textvariable=self.entry)
        self.entryfield.place(x=150, y=75)
        self.btn_check = Button(self.top, text="Check", command=self.check_entry)
        self.btn_check.place(x=350, y=75)

        self.lblprogress = Label(self.top, text="")
        self.lblprogress.place(x=150, y=155)
        #self.progress.pack()


        self.lblreview = Label(self.top, text="")
        self.lblreview.place(x=150, y=175)
        #self.progress.pack()

        self.lblaccuracy = Label(self.top, text="")
        self.lblaccuracy.place(x=150, y=195)

        self.lblhint= Label(self.top, text="Hint?")
        self.lblhint.place(x=150, y=225)

        self.btn_hint = Button(self.top, text="Hint", command=self.hintitem)
        self.btn_hint.place(x=50, y=225)


        self.drawindex()

        self.top.bind('<Return>', (lambda e, b=self.btn_check: b.invoke()))

        self.btn_quit = Button(self.top, text="Done", command=self.quit)
        self.btn_quit.place(x=150, y=250)
        self.btn_quit.place_forget()

        self.btn_reset = Button(self.top, text="Reset", command=self.reset)
        self.btn_reset.place(x=350, y=250)
        self.btn_reset.place_forget()


        # b = Button(self.top, text="Next", command=self.nextitem)
        # b.pack(pady=5)
        # b = Button(self.top, text="Previous", command=self.previousitem)
        # b.pack(pady=5)

    def reset(self):
        self.itemlist = copy.copy(self.safelist)
        self.missedlist = list()
        random.shuffle(self.itemlist)

        self.itemindex = 0
        self.misscount = 0
        self.totalmisscount = 0
        self.hitcount = 0
        self.totalhitcount = 0
        self.missedit = False
        self.review = False
        self.listcount = len(self.itemlist)
        self.btn_quit.place_forget()
        self.btn_reset.place_forget()
        self.btn_hint.configure(state=NORMAL)
        self.btn_check.configure(state=NORMAL)

        self.drawindex()

        self.drawprogress()

    def misseditem(self):
        self.missedit = True
        self.misscount = self.misscount + 1
        if self.review == False:
            self.totalmisscount = self.totalmisscount + 1
        self.missedlist.append(self.itemlist[self.itemindex])

    def check_entry(self):
        item = self.itemlist[self.itemindex]
        row,col,value = item.get()
        print "Testing: %s" % (self.itemtext())
        print "Correct Value: %s" % (value)
        if (value == self.entry.get()):
            print "Correct Entry: %s" % (self.entry.get())
            if (self.missedit == False):
                self.hitcount  =self.hitcount + 1
                if self.review == False:
                    self.totalhitcount = self.totalhitcount + 1
            self.nextitem()
        else:
            print "Incorrect Entry: %s" % (self.entry.get())
            if (self.missedit == False):
                self.misseditem()
            self.drawindex()

    def hintitem(self):
        item = self.itemlist[self.itemindex]
        row,col,value = item.get()
        print "Correct Value: %s" % (value)

        self.lblhint.config(text=value)
        self.misseditem()


    def nextitem(self, event=None):
        self.lblhint.config(text="Hint?")
        self.missedit = False
        if self.itemindex + 1 < self.listcount:
            self.itemindex = self.itemindex +1
        else:
            if (len(self.missedlist) > 0):
                self.review = True
                self.itemlist = copy.copy(self.missedlist)
                self.itemindex = 0
                self.misscount = 0
                self.hitcount = 0
                self.listcount = len(self.missedlist)
                self.missedlist = list()

        self.drawindex()

    def previousitem(self, event=None):
        self.missedit = False
        if self.itemindex > 0:
            self.itemindex = self.itemindex -1

        self.drawindex()

    def drawprogress(self):
        progresstext = "Missed: %s Correct: %s Progress: %s of %s" % (
        self.misscount, self.hitcount, self.itemindex + 1, self.listcount)
        if (self.review == False):
            self.lblprogress.config(text=progresstext)
            self.lblreview.config(text="")
        else:
            progresstext = "Review! %s" % (progresstext)
            self.lblreview.config(text=progresstext)

        if (self.totalmisscount + self.totalhitcount > 0):
            accuracy = (float(self.totalhitcount) / float(self.totalmisscount + self.totalhitcount)) * 100
            accuracytext = "Accuracy: %3.2f%%" % ( accuracy)
            self.lblaccuracy.config(text=accuracytext)
        else:
            self.lblaccuracy.config(text="")

    def itemtext(self):

        item = self.itemlist[self.itemindex]
        row, col, value = item.get()
        itemtext = "%s : %s" % (row, col)
        return itemtext

    def drawindex(self):
        if (self.missedit == False):
            self.entryfield.configure(background='white')
        else:
            self.entryfield.configure(background='red')

        if(self.hitcount == self.listcount ):
            #progresstext = "Finished!"
            self.drawprogress()
            self.btn_hint.configure(state=DISABLED)
            self.btn_check.configure(state=DISABLED)

            self.btn_quit.place(x=150, y=250)
            self.btn_reset.place(x=350, y=250)

        else:
            self.entry.set("")
            self.labelitem.config(text=self.itemtext())
            self.drawprogress()


    def quit(self, event=None):
        self.top.destroy()

    def cancel(self, event=None):
        self.top.destroy()

root = Tk()
root.title("Learn from Spreadsheet")
main_ui = GUI(root)
root.mainloop()