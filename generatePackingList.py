from datetime import date
from docxtpl import DocxTemplate, InlineImage
from docx import Document
from docx.shared import Cm, Inches, Mm, Emu, Pt
from docx.oxml.shared import OxmlElement
from docx.oxml.ns import qn
import pandas as pd
import numpy as np
import xlwings as xw
from xlrd import open_workbook
from xlutils.copy import copy
import os
import shutil

def makePList(sheet, parts, part_ppty, summary, property, batch, po_num, 
              thedate, numboxes, template, dest):
    context = dict()
    # Makes a dataframe where a code is matched to each of its S/N's
    unique_parts = parts.groupby('Code')['S/N'].apply(list).reset_index(name = 'S/N')

    # print(unique_parts)
    # print(summary)

    numcodes = unique_parts.shape[0]

    desc = ""

    # Generate description (all parts)
    for i in range(numcodes):
        code_desc = ""
        code = unique_parts.loc[i].at['Code']
        sns = unique_parts.loc[i].at['S/N']
        start = -1000
        last = -1000
        counter = 0
        for sn in sorted(sns):
            if sn == last+1:
                if(counter == len(sns)-1):
                    code_desc += (' to ' + str(code) + '-' + str(last))
                pass
            else:
                if(start != last):      # Last code was consecutive
                    code_desc += (' to ' + str(code) + '-' + str(last))
                    code_desc += (', ' + str(code) + '-' + str(sn))
                else:
                    code_desc += (', ' + str(code) + '-' + str(sn))
                start = sn
            last = sn
            counter += 1
        desc += code_desc
    desc = desc[2:]
    context["Desc"] = desc

    context["numcodes"] = range(numcodes)

    context["Num_Tested"] = part_ppty.shape[0]
    context["Property"] = property
    context["Batch"] = batch
    context["PO"] = po_num
    context["Date"] = str(thedate)[:10]  # Slice out the time
    context['nboxes'] = str(numboxes)
    if(numboxes == 1):
        context['boxes'] = "box"
    else:
        context['boxes'] = 'boxes'
    part_specs = dict()

    updates = []
    totals = []
    
    for i in range(numcodes):
        specs = dict()
        code = unique_parts.loc[i].at['Code']
        # print("code ", code)
        # List of s/n's
        sns = unique_parts.loc[i].at['S/N']
        specs['this'] = len(sns)
        code_summ = summary.loc[summary['Code'] == code]
        # print(code_summ)
        specs['quantity'] = int(code_summ.iloc[0].at['Quantity Ordered'])
        specs['prev'] = int(code_summ.iloc[0].at['Previous Shipment'])
        specs['total'] = specs['this'] + specs['prev']
        specs['rem'] = specs['quantity'] - specs['total']
        specs['code'] = code
        if code == 'M38':
            specs['PN'] = '1110-038-2'
        elif code == 'M39':
            specs['PN'] = '1110-039-2'
        elif code == 'M26':
            specs['PN'] = '1120-026-2'
        elif code == 'M30':
            specs['PN'] = '1130-030-2'
        elif code == 'M17':
            specs['PN'] = '1141-017-2'
        elif code == 'M18':
            specs['PN'] = '1141-018-2'
        elif code == 'M40':
            specs['PN'] = '1142-040-2'
        elif code == 'M37':
            specs['PN'] = '1142-037-2'
        elif code == 'M42':
            specs['PN'] = '1142-042-2'
        elif code == 'M43':
            specs['PN'] = '1142-043-2'

        part_specs[code] = specs

        total = pd.DataFrame({'Total': [specs['total']]})
        totals.append(total)

        data = pd.DataFrame([[specs['this'], specs['total'], specs['rem']]])
        updates.append(data)
    context['specs'] = part_specs

    totals = pd.concat(totals)
    totals.index = [x for x in range(totals.shape[0])]
    updates = pd.concat(updates)
    updates.index = [x for x in range(updates.shape[0])]

    pList = DocxTemplate(template)
    pList.render(context)

    savepath = dest + '\\' +  context['Batch']
    filepath = savepath + '\\Packing List' + '_'+str(batch) + '.docx'

    if not os.path.exists(savepath):
        os.mkdir(savepath)

    if os.path.isfile(filepath):
        filename, extension = os.path.splitext(filepath)
        counter = 1
        while os.path.isfile(filepath):
            filepath = filename + " (" + str(counter) + ")" + extension
            counter += 1
        pList.save(filepath)
    else:
        pList.save(filepath)

    app = xw.App(visible=False)
    new_loc = os.path.split(sheet)[0]
    print(new_loc)

    # Copy template sheet into batch folder
    batchsheet = savepath + '\\' + str(batch) + '.xlsx'
    shutil.copy(sheet, batchsheet)

    bb = xw.Book(batchsheet)
    bs = bb.sheets['Summary']

    # # Fill in the summary columns
    bs.range('F4').options(index=False, header=False).value = updates
    bb.save()
    bb.close()

    # # Clear existing sheet for new template
    wb = xw.Book(sheet)

    # # Clear batch and date
    wb.sheets['Part #'].range('B2').clear_contents()
    wb.sheets['Part #'].range('B3').clear_contents()
    wb.sheets['Part #'][7:,:].clear_contents()

    wb.sheets['Property'][5:,:].clear_contents()

    wb.sheets['Summary'][3:,1:].clear_contents()

    wb.sheets['Summary'].range('E4').options(index=False, header=False).value = totals

    wb.save()
    wb.close()

    app.quit()

    

