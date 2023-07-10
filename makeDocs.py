from datetime import date
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Cm, Inches, Mm, Emu
import pandas as pd
import numpy as np
import os

from generateLabels import *
from generateCoC import *
from generatePackingList import *

######################################
##### Files Directories Are Here #####
######################################

TempDir = "C:\\Justin\\TreadStone\\Project\\InfinityDocs\\Templates"
ParentDir = "C:\\Justin\\TreadStone\\Project\\InfinityDocs"


### Label Template 3 and CoC Template 2 or 3
LabelTempFile = 'Label Template3.docx'
CoCTempFile = 'CoC IFC Template3.docx'
PLTempFile = 'Packing List Template.docx'

#############################################
##### Previous Shipment Inputs Are Here #####
#############################################

# print('Enter the parts file name:')
# File = input()
# PartsFile = File + '.xlsx'
PartsFile = "New input file format.xlsx"
sheet = "C:\\Justin\\TreadStone\\Project\\testing\\New input file format.xlsx"

Labels = DocxTemplate(TempDir+'\\'+LabelTempFile)
# sheet = ParentDir + "\\" + PartsFile

info = pd.read_excel(sheet, nrows=3, usecols = 'A, B', header=None)
po_num = info.loc[0][1]
batch = info.loc[1][1]
thedate = info.loc[2][1]

packs = pd.read_excel(sheet, skiprows = 6)

# Parts in each pack
packlist = []

i = 0
counter = 1
while i < packs.shape[1]:
    pack = packs.iloc[:, [i,i+1]]
    # makeLabels(pack, po_num, batch, thedate, counter)
    pack.columns = ['Code', 'S/N']
    packlist.append(pack)
    i += 3
    counter += 1

# All parts
parts = pd.concat(packlist, ignore_index=True)


properties = pd.read_excel(sheet, sheet_name = 'Property', nrows=3, usecols = 'A, B', header=None)
ppty = properties.loc[0][1]
unit = properties.loc[1][1]

part_ppty = pd.read_excel(sheet, sheet_name = 'Property', skiprows=4)

# makeCoC(parts, batch, thedate, part_ppty, ppty, unit, po_num)

summary = pd.read_excel(sheet, sheet_name = 'Summary', skiprows=2, usecols = "B:E")
summary = summary.dropna().reset_index(drop=True)


# makePList(parts, part_ppty, summary, ppty, batch, po_num, sheet, thedate, counter-1)