from datetime import date
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Cm, Inches, Mm, Emu

import tkinter as tk
from tkinter import ttk
from tkinter.scrolledtext import ScrolledText
from tkinter import filedialog

import pandas as pd
import numpy as np
import os

from generateLabels import *
from generateCoC import *
from generatePackingList import *

class Page(tk.Frame):
    def __init__(self, container, app):
        tk.Frame.__init__(self, container)

        self.app = app

        coc_template = ttk.Button(self, text = "Choose CoC\nTemplate", command = 
                           self.getFolder)
        coc_template.grid(row=1, column=0, padx = 10, pady = 10)

        plist_template = ttk.Button(self, text="Choose Packing\nList Template", 
                                  command = self.getFolder)
        plist_template.grid(row=1, column=1, padx = 10, pady = 10)

        generate = ttk.Button(self, text = "Generate\nDocument", command = 
                           lambda: app.reg_read(False))
        generate.grid(row = 3, column = 0, padx = 10, pady = 10)
    
    def getFolder(self):
        return filedialog.askdirectory(initialdir = "/",title = "Select folder")

class Application(tk.Tk):
    def __init__(self, *args, **kwargs):
        # __init__ function for class Tk
        tk.Tk.__init__(self, *args, **kwargs)
            
        # creating a container
        container = tk.Frame(self) 
        container.pack(side = "top", fill = "both", expand = True)
    
        container.grid_rowconfigure(0, weight = 1)
        container.grid_columnconfigure(0, weight = 1)

        self.frames = {}
        page = Page(container, self)
        self.frames[Page] = page

        self.book = ""
        self.coc_template = ""
        self.plist_template = ""
                
        self.show_frame(Page)

    # Show a page on the window
    def show_frame(self, page):
        frame = self.frames[page]
        frame.tkraise()
        self.update()

app = Application()

app.mainloop()