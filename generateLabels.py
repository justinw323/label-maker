from datetime import date
from docxtpl import DocxTemplate, InlineImage
from docx import Document
from docx.shared import Cm, Inches, Mm, Emu, Pt
from docx.oxml.shared import OxmlElement
from docx.oxml.ns import qn
import pandas as pd
import numpy as np
import os



#######################################################
##### This section pertains to the label template #####
#######################################################
def makeLabels(parts, po_num, batch, thedate, num, dest):
    parts.columns = ['Code', 'S/N']
    context = dict()
    context['Batch'] = batch
    numParts = parts.shape[0]
    keysL = []
    serialL =[]
    #this generates the pos and code sections of the labels
    for i in range(len(parts)):
        keysL.append('pos'+str(i))
        keysL.append('code'+str(i))
        serialL.append(f"{int(parts.loc[i].at['S/N']):03}")
        code = str(parts.loc[i].at['Code'])
        serialL.append(code)
        
        #This generates the PN section of the labels
        ##############################################
        ### THIS SECTION is for the P/N for labels ###
        ##############################################
        # Code corresponds to a P/N
        keysL.append('PN'+str(i))
        if code == 'M38':
            serialL.append('1110-038-2')
        elif code == 'M39':
            serialL.append('1110-039-2')
        elif code == 'M26':
            serialL.append('1120-026-2')
        elif code == 'M30':
            serialL.append('1130-030-2')
        elif code == 'M17':
            serialL.append('1141-017-2')
        elif code == 'M18':
            serialL.append('1141-018-2')
        elif code == 'M40':
            serialL.append('1142-040-2')
        elif code == 'M37':
            serialL.append('1142-037-2')
        elif code == 'M42':
            serialL.append('1142-042-2')
        elif code == 'M43':
            serialL.append('1142-043-2')

    context.update(dict(zip(keysL, serialL)))

    doc = Document()
    style = doc.styles['Normal']
    font = style.font
    style.paragraph_format.space_after = Pt(0)
    font.name = 'Times New Roman'
    font.size = Pt(12)
    for section in doc.sections:
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)

    table = doc.add_table(rows = numParts//3 + (0 if numParts % 3 == 0 else 1), 
                          cols = 3)

    i = 0
    # Make each row of the label sheet
    for r in range(len(table.rows)):
        row = table.rows[r].cells
        for c in range(3):
            if(i >= numParts):
                break        
            code = str(parts.loc[i].at['Code'])
            pn = ''
            if code == 'M38':
                pn = '1110-038-2'
            elif code == 'M39':
                pn = '1110-039-2'
            elif code == 'M26':
                pn = '1120-026-2'
            elif code == 'M30':
                pn = '1130-030-2'
            elif code == 'M17':
                pn = '1141-017-2'
            elif code == 'M18':
                pn = '1141-018-2'
            elif code == 'M40':
                pn = '1142-040-2'
            elif code == 'M37':
                pn = '1142-037-2'
            elif code == 'M42':
                pn = '1142-042-2'
            elif code == 'M43':
                pn = '1142-043-2'
            row[c].text = ('Batch #: ' + str(context['Batch']) + '\n' + 
                        'P/N: ' + pn + '\n' + 
                        'Serial #: ' + code + '-' + 
                                f"{int(parts.loc[i].at['S/N']):04}" + '\n' + 
                        'Parts #: ' + str(i+1) + '\n')
            i += 1


    batchpath = dest + '\\' +  context['Batch']
    savepath = dest + '\\' +  context['Batch'] + '\\Labels'
    filepath = savepath + '\\' + context['Batch'] +'_Pack' + str(num) + '_labels.docx'

    if not os.path.exists(batchpath):
        os.mkdir(batchpath)

    if not os.path.exists(savepath):
        os.mkdir(savepath)

    if os.path.isfile(filepath):
        filename, extension = os.path.splitext(filepath)
        counter = 1
        while os.path.isfile(filepath):
            filepath = filename + " (" + str(counter) + ")" + extension
            counter += 1
        return (filepath, doc)
    else:
        return (filepath, doc)