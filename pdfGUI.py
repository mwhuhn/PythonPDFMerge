# -*- coding: utf-8 -*-
#!/usr/bin/env python 1

"""
This is a GUI pdf merger program. You can select individual files or
whole folders to add to a  merger list. You can move files up or down
in the merger order and remove files.
"""

import Tkinter as tk
import tkFileDialog
import os
from PyPDF2 import PdfFileReader, PdfFileMerger
from tkMessageBox import showinfo
#import re # add if using numerical order for folder


class MainView(tk.Tk):
    """ 
        Main view of GUI
        Included to make adding new frames easy
        To add frame, add frame name to initialization command and add class(tk.Frame)
    """
    def __init__(self, *args, **kwargs):
        
        tk.Tk.__init__(self, *args, **kwargs)
        
        container = tk.Frame(self)
        container.grid(row=0)

        self.frames = {}
        
        """ Initialize subpages
        Creates a dictionary connecting string page names (key) to the
        frames (value)
        """
        for F in [SelectFilesPage]:
            page = F.__name__
            frame = F(parent=container, controller=self)
            self.frames[page] = frame
            frame.grid(row=0, column=0, sticky="nsew")
        
        # Show first page
        self.showFrame("SelectFilesPage")

    """ Raise frame function
    Uses frames dictionary to call a frame from the page name (key)    
    """
    def showFrame(self, page):
        frame = self.frames[page]
        frame.tkraise()

class SelectFilesPage(tk.Frame):
    """ Page allows you to select the pdf files to merge in order
    
    """
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller
        self.lastdir = os.path.expanduser('~') # set home as initial directory

        spacer1 = tk.Label(self, text="")
        spacer1.grid(row=0, column=0, columnspan=3, sticky="we")        

        label = tk.Label(self, text="Select your files to merge.")
        label.grid(row=1, column=0, columnspan=3, sticky="w")

        spacer2 = tk.Label(self, text="")
        spacer2.grid(row=2, column=0, columnspan=3, sticky="we")        


        # Filenames listbox, updates with added files
        
        scrollbarV = tk.Scrollbar(self)
        scrollbarH = tk.Scrollbar(self, orient="horizontal")
        listbox = tk.Listbox(self, width = 70,
                             yscrollcommand=scrollbarV.set,
                             xscrollcommand=scrollbarH.set,
                             selectmode=tk.EXTENDED)
        scrollbarV.config(command=listbox.yview)
        scrollbarH.config(command=listbox.xview)

        addFileButton = tk.Button(self, text="Add File",
                                  command=lambda: self.addFile(listbox))
        addFolderButton = tk.Button(self, text="Add Entire Folder",
                                  command=lambda: self.addFolder(listbox))
        clearButton = tk.Button(self, text="Clear All",
                                command=lambda: listbox.delete(0,tk.END))

        addFileButton.grid(row=3, column=0, sticky="we")
        addFolderButton.grid(row=3, column=1, sticky="we")
        clearButton.grid(row=3, column=2, sticky="we")
        listbox.grid(row=4, column=0, columnspan=3, sticky="we")
        scrollbarV.grid(row=4, column=3, sticky="ns")
        scrollbarH.grid(row=5, column=0, columnspan=3, sticky="we")

        # interact with list buttons
        moveUpButton = tk.Button(self, text="Move Up",
                                 command=lambda: self.moveSelectedUp(listbox))
        moveUpButton.grid(row=6, column=0, sticky="we")
        moveDownButton = tk.Button(self, text="Move Down",
                                   command=lambda: self.moveSelectedDown(listbox))
        moveDownButton.grid(row=6, column=1, sticky="we")
        removeButton = tk.Button(self, text="Remove",
                                 command=lambda: self.removeSelected(listbox))
        removeButton.grid(row=6, column=2, sticky="we")

        # Buttons to navigate pages
                                                               
        quitButton = tk.Button(self,
                               text="Quit",
                               command=controller.destroy)
        mergeButton = tk.Button(self,
                                text="Merge Files",
                                command=lambda: self.mergePDF(listbox))
                               
        quitButton.grid(row=7, column=0, sticky="we")
        mergeButton.grid(row=7, column=1, sticky="we")
    
    
    # Add filename/directory to listbox    
    def addFile(self, listbox):
        name = tkFileDialog.askopenfilename(filetypes=[('PDF files', '.pdf')],
                                            initialdir = self.lastdir)
        if not name:
            return # cancels if no file selected
        self.lastdir = os.path.split(name)[0]
        listbox.insert("end", name)
    
    # Add all files in folder to listbox, in str.lower order
    def addFolder(self, listbox):
        directory = tkFileDialog.askdirectory(initialdir = self.lastdir)
        self.lastdir = directory # update last directory used
        pdf_files = [f for f in os.listdir(directory) if f.endswith("pdf")]
        if len(pdf_files) == 0:
            return  # cancels if no pdf files in folder
        def fileKeyAlph(filename):
            f = filename.encode("utf-8") # filenames from unicode to str
            f = f.lower()
            return f
        # change default sort to numeric here (by first full number)
#        def fileKeyNum(filename):
#            pattern = re.compile("[0-9]+")
#            num = pattern.search(filename)
#            if not num:
#                return 0 # puts non-numbered files first
#            return int(num.group())
        sortedFiles = sorted(pdf_files, key=fileKeyAlph)
        for filename in sortedFiles:
            listbox.insert("end", os.path.join(directory, filename).replace("\\","/"))
    
    # Moves selected file up one position
    def moveSelectedUp(self, listbox):
        selected = listbox.curselection()
        if not selected:
            return # cancels if no file selected
        nomovelast = False
        for i in selected:
            if i == 0: # don't move up if i is 0
                last = 0
                nomovelast = True
                continue
            elif nomovelast: # don't move up if last didn't move up
                if i - 1 == last: # but only if they're 1 away from eachother
                    last = i
                    continue
                else:
                    text = listbox.get(i)
                    listbox.delete(i)
                    listbox.insert(i-1, text)
                    listbox.selection_set(i-1)                    
            else: # move selected up one and keep it selected
                text = listbox.get(i)
                listbox.delete(i)
                listbox.insert(i-1, text)
                listbox.selection_set(i-1)

        
    # Moves selected file down one position
    def moveSelectedDown(self, listbox):
        selected = listbox.curselection()
        if not selected:
            return # cancels if no file selected
        nomovelast = False
        selected = selected[::-1]
        for i in selected:
            if i == listbox.size() - 1: # don't move down if i is last
                last = i
                nomovelast = True
                continue
            elif nomovelast: # don't move up if last didn't move up
                if i + 1 == last: # but only if they're 1 away from eachother
                    last = i
                    continue
                else:
                    text = listbox.get(i)
                    listbox.delete(i)
                    listbox.insert(i+1, text)
                    listbox.selection_set(i+1)                    
            else: # move selected up one and keep it selected
                text = listbox.get(i)
                listbox.delete(i)
                listbox.insert(i+1, text)
                listbox.selection_set(i+1)

    
    # Removes selected file
    def removeSelected(self, listbox):
        selected = listbox.curselection()
        if not selected:
            return # cancels remove if none selected
        selected = selected[::-1]
        for i in selected:
            listbox.delete(i)
        
     # Merges all pdf files in list
    def mergePDF(self, listbox):
        # ensure only pdf files included
        name = tkFileDialog.asksaveasfilename(defaultextension=".pdf",
                                              filetypes=[('PDF files', '.pdf')],
                                              initialdir = self.lastdir)
        if not name:
            return # cancels merge if no save name
        sortedFiles = listbox.get(0,tk.END)
        merger = PdfFileMerger()
        for filename in sortedFiles:
            merger.append(PdfFileReader(filename), "rb")
        merger.write(name)
        showinfo("Finished", "Files merged.")

if __name__ == "__main__":
    root = MainView()
    root.mainloop()
