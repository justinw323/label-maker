from datetime import date
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Cm, Inches, Mm, Emu
import pandas as pd
import numpy as np
import os

from generateLabels import *
from generateCoC import *
from generatePackingList import *



def generateDocs(input, plist, coc, dest):

    info = pd.read_excel(input, nrows=3, usecols = 'A, B', header=None)
    po_num = info.loc[0][1]
    batch = info.loc[1][1]
    thedate = info.loc[2][1]

    packs = pd.read_excel(input, skiprows = 6)

    # Parts in each pack
    packlist = []

    i = 0
    counter = 1
    while i < packs.shape[1]:
        pack = packs.iloc[:, [i,i+1]]
        makeLabels(pack, po_num, batch, thedate, counter, dest)
        pack.columns = ['Code', 'S/N']
        packlist.append(pack)
        i += 3
        counter += 1

    # All parts
    parts = pd.concat(packlist, ignore_index=True)


    properties = pd.read_excel(input, sheet_name = 'Property', nrows=3, usecols = 'A, B', header=None)
    ppty = properties.loc[0][1]
    unit = properties.loc[1][1]

    part_ppty = pd.read_excel(input, sheet_name = 'Property', skiprows=4)

    makeCoC(parts, batch, thedate, part_ppty, ppty, unit, po_num, coc, dest)

    summary = pd.read_excel(input, sheet_name = 'Summary', skiprows=2, usecols = "B:E")
    summary = summary.dropna().reset_index(drop=True)


    makePList(input, parts, part_ppty, summary, ppty, batch, po_num, thedate, 
              counter-1, plist, dest)
    
    return dest + '/' + batch