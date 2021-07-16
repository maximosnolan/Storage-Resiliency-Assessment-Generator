from docx import Document
from numpy import NaN
import pandas
import csv
import PIL
from PIL import ImageGrab, Image
from docx.shared import Cm, Pt, RGBColor
import os
import sys
from copy import deepcopy
from pandas.io import excel
from multiprocessing import Pool
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
import glob
from docx.enum.text import WD_ALIGN_PARAGRAPH
#import win32com.client as win32
class Recomendation:
    def __init__(self, importance, recommendation, framework, rank):
        self.importance = importance
        self.recommendation = recommendation
        self.framework = framework
        self.rank = rank

    #importance = 0
    #rec = ""
    #framework = ""

major_reccomendation = [] #major recomendations
minor_reccomendation = [] #minor recomendations
final_scores = {}

color_map = {"Respond": 15120823, "Detect": 16570036, "Protect": 14999532, "Identify": 14279154, "Recover": 15528671}

class company:
    def __init__(self, assessment, name, date):
        self.assessment = assessment
        self.company_name = name
        self.date = date  
    date = ""
    company_name = ""
    assessment = ""


def get_information():
    assessment = input("Please enter the assessment file name:")
    name = input("Enter the companies name:")
    date = input("Enter the date the assessment will be presented")
    victim_comapny = company(assessment, name, date)
    return victim_comapny

def print_intro():
    print("Welcome to the IBM Storage Resiliency Report Automation Software!")
    print("Use guide:")
    print("Please have the assessment in the folder where you run this program and the templated word doc provided.")
    print("Once you run the script a new document will appear named 'final report'")
    return

def replace_all(company, document, randsom_percent, final_scores):
    document.add_paragraph(company.company_name)
    style = document.styles['Normal']
    font = style.font
    font.name = 'IBM Plex Sans Light'
    font.size = Pt(9)
    for paragraph in document.paragraphs:
        if 'Executive Summary – Summary View' in paragraph.text:
            df = copy_total_score()
            document.tables[1].cell(1,1).text = str(round(df.loc[2, 'SCORE'], 2))
            for x in range (4,35):
                if x != 10 and x != 17 and x != 22 and x != 29 and x!= 35:
                 
                   document.tables[1].cell(x - 1, 1).text = str(round(df.loc[x, 'SCORE'], 2))
                   document.tables[1].cell(x - 1, 1).allignment = 1
                   document.tables[1].cell(x - 1, 2).text = str(df.loc[x, 'LEVEL'])
                   document.tables[1].cell(x - 1, 2).allignment = 1
                
        if 'CUSTOMER_NAME' in paragraph.text:
            orig_text = paragraph.text
            new_text = str.replace(orig_text, "CUSTOMER_NAME", company.company_name)
            paragraph.text = new_text
            style = document.styles['Normal']
            font = style.font
            font.name = 'IBM Plex Sans Light'
            font.size = Pt(9)
            paragraph.style = document.styles['Normal']
        if 'PRESENTATION_DATE' in paragraph.text:
            orig_text = paragraph.text
            new_text = str.replace(orig_text, "PRESENTATION_DATE", company.date)
            paragraph.text = new_text
            style = document.styles['Normal']
            font = style.font
            font.name = 'IBM Plex Sans Light'
            font.size = Pt(9)
            paragraph.style = document.styles['Normal']
        if 'PERCENT_RANSOM' in paragraph.text:
            orig_text = paragraph.text
            percent = str(randsom_percent) + '%'
            new_text = str.replace(orig_text, "PERCENT_RANSOM", percent)
            paragraph.text = new_text
            style = document.styles['Normal']
            font = style.font
            font.name = 'IBM Plex Sans Light'
            font.size = Pt(9)
            paragraph.style = document.styles['Normal']
        if 'PERCENT _MISS' in paragraph.text:
            orig_text = paragraph.text
            missing = str(100 - randsom_percent) + '%'
            new_text = str.replace(orig_text, "PERCENT _MISS", missing)
            paragraph.text = new_text
            style = document.styles['Normal']
            font = style.font
            font.name = 'IBM Plex Sans Light'
            font.size = Pt(9)
            paragraph.style = document.styles['Normal']
        if 'EX_1' in paragraph.text:
            orig_text = paragraph.text
            #set new text depending on terminating char
            new_text = str.replace(orig_text, "EX_1", list(final_scores.keys())[15])
            paragraph.text = new_text
            style = document.styles['Normal']
            font = style.font
            font.name = 'IBM Plex Sans Light'
            font.size = Pt(9)
            paragraph.style = document.styles['Normal']
        if 'EX_2' in paragraph.text:
            orig_text = paragraph.text
            #set new text depending on terminating char
            new_text = str.replace(orig_text, "EX_2", list(final_scores.keys())[16])
            paragraph.text = new_text
            style = document.styles['Normal']
            font = style.font
            font.name = 'IBM Plex Sans Light'
            font.size = Pt(9)
            paragraph.style = document.styles['Normal']
        if 'EX_3' in paragraph.text:
            orig_text = paragraph.text
            #set new text depending on terminating char
            new_text = str.replace(orig_text, "EX_3", list(final_scores.keys())[17])
            paragraph.text = new_text
            style = document.styles['Normal']
            font = style.font
            font.name = 'IBM Plex Sans Light'
            font.size = Pt(9)
            paragraph.style = document.styles['Normal']
        if 'FIRST ISSUE' in paragraph.text:
            orig_text = paragraph.text
            #set new text depending on terminating char
            new_text = str.replace(orig_text, "FIRST ISSUE", list(final_scores.keys())[0])
            paragraph.text = new_text
            style = document.styles['Normal']
            font = style.font
            font.name = 'IBM Plex Sans Light'
            font.size = Pt(9)
            paragraph.style = document.styles['Normal']
        if 'SECOND ISSUE' in paragraph.text:
            orig_text = paragraph.text
            print(orig_text)
            new_text = str.replace(orig_text, "SECOND ISSUE", list(final_scores.keys())[1])
            paragraph.text = new_text
            style = document.styles['Normal']
            font = style.font
            font.name = 'IBM Plex Sans Light'
            font.size = Pt(9)
            paragraph.style = document.styles['Normal']
        if 'THIRD ISSUE' in paragraph.text:
            orig_text = paragraph.text
            #set new text depending on terminating char
            new_text = str.replace(orig_text, "THIRD ISSUE", list(final_scores.keys())[2])
            paragraph.text = new_text
            style = document.styles['Normal']
            font = style.font
            font.name = 'IBM Plex Sans Light'
            font.size = Pt(9)
            paragraph.style = document.styles['Normal']
        if 'FOURTH ISSUE' in paragraph.text:
            orig_text = paragraph.text
            #set new text depending on terminating char
            new_text = str.replace(orig_text, "FOURTH ISSUE", list(final_scores.keys())[3])
            paragraph.text = new_text
            style = document.styles['Normal']
            font = style.font
            font.name = 'IBM Plex Sans Light'
            font.size = Pt(9)
            paragraph.style = document.styles['Normal']
        if 'SCORE  is slightly below average' in paragraph.text:
            orig_text = paragraph.text
            #set new text depending on terminating char
            conjunction = ""
            if round(df.loc[2, 'SCORE'], 2) < 6:
                conjunction = " is below average"
            elif round(df.loc[2, 'SCORE'], 2) < 7:
                conjunction = " is slightly below average"
            elif round(df.loc[2, 'SCORE'], 2) > 9:
                conjunction = " is above average"
            elif round(df.loc[2, 'SCORE'], 2) > 8:
                conjunction = " is slightly above average"
            else:
                conjunction = " is average"
            new_text = str.replace(orig_text, "SCORE  is slightly below average",str(round(df.loc[2, 'SCORE'], 2)) + conjunction )
            paragraph.text = new_text
            style = document.styles['Normal']
            font = style.font
            font.name = 'IBM Plex Sans Light'
            font.size = Pt(9)
            paragraph.style = document.styles['Normal']
    x = 0
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    if x == 0: #if this is the first table
                        continue
                    if 'CUSTOMER_NAME' in paragraph.text:
                        paragraph.text = paragraph.text.replace("CUSTOMER_NAME", company.company_name)
                    if 'PRESENTATION_DATE' in paragraph.text:
                        paragraph.text = paragraph.text.replace("PRESENTATION_DATE", company.date)
        x+=1

    
    return

def copy_total_score():
    excel_data_df = pandas.read_excel('CRAT_example.xlsx', sheet_name= 'Results')
    return excel_data_df
def pull_level(score):
    level = 0
    if score == "Identify":
        level = 1
    elif score == "Protect":
        level = 2
    elif score == "Detect":
        level = 3
    elif score == "Respond":
        level = 4
    else:
        level = 5
    return level

def pull_reccomendations():
    excel_data_df = pandas.read_excel('CRAT_example.xlsx', sheet_name= 'Internal Data')
    #227 corresponds to first question "Inventory storage physical devices...."
    for x in range (227, 399):
        #1 determine if yes 
        level = pull_level(str(excel_data_df.loc[x, 'Score']))
        if str(excel_data_df.loc[x, 'Answer']) == 'Yes':
            continue
        elif str(excel_data_df.loc[x, 'Answer']) == 'No':
            a = Recomendation(str(excel_data_df.loc[x, 'Count']), str(excel_data_df.loc[x, 'Questions']), str(excel_data_df.loc[x, 'Score']), level)
            major_reccomendation.append(a)
        else:
            a = Recomendation(str(excel_data_df.loc[x, 'Count']), str(excel_data_df.loc[x, 'Questions']), str(excel_data_df.loc[x, 'Score']), level)
            minor_reccomendation.append(a)

    return

def pull_score():
    scores = {}
    excel_data_df = pandas.read_excel('CRAT_example.xlsx', sheet_name= 'Internal Data')
    for x in range (1, 19):
        scores[excel_data_df.loc[x, 'Category']] = excel_data_df.loc[x, 'Percent_Ach']
    final_scores = {k: v for k, v in sorted(scores.items(), key=lambda item: item[1])}
    for x, y in final_scores.items():
        print(x, y) 
    return final_scores

def copy_table_after(table, paragraph):
    tbl, p = table._tbl, paragraph._p
    new_tbl = deepcopy(tbl)
    p.addnext(new_tbl)

def generate_table(document, maj_min, number):
    table = document.tables[number]
    for x in range(0, len(maj_min) ):
            if(len(maj_min) <= x):
                continue
            rec = maj_min[x]
            importance = rec.importance
            text = rec.recommendation
            type = rec.framework
            document.tables[number].cell(x,0).text = importance
            document.tables[number].cell(x,0).paragraphs[0].runs[0].font.size = Pt(9)
            document.tables[number].cell(x,0).paragraphs[0].runs[0].font.name = 'Arial'
            
            document.tables[number].cell(x,1).text = text
            document.tables[number].cell(x,1).paragraphs[0].runs[0].font.size = Pt(9)
            document.tables[number].cell(x,1).paragraphs[0].runs[0].font.name = 'Arial'
            
           
            if type == "Identify":
                shading_elm = parse_xml(r'<w:shd {} w:fill="D9E1F2"/>'.format(nsdecls('w')))
                document.tables[number].cell(x,0)._tc.get_or_add_tcPr().append(parse_xml(r'<w:shd {} w:fill="D9E1F2"/>'.format(nsdecls('w'))))
                p2 = table.cell(x,1)
                p2._tc.get_or_add_tcPr().append(shading_elm)
            elif type == "Protect":
                shading_elm = parse_xml(r'<w:shd {} w:fill="E4DFEC"/>'.format(nsdecls('w')))
                p1 = table.cell(x,0)
                p2 = table.cell(x,1)
                document.tables[number].cell(x,0)._tc.get_or_add_tcPr().append(parse_xml(r'<w:shd {} w:fill="E4DFEC"/>'.format(nsdecls('w'))))
                p2._tc.get_or_add_tcPr().append(shading_elm)
            elif type == "Respond":
                shading_elm = parse_xml(r'<w:shd {} w:fill="E6B9B7"/>'.format(nsdecls('w')))
                p1 = table.cell(x,0)
                p2 = table.cell(x,1)
                document.tables[number].cell(x,0)._tc.get_or_add_tcPr().append(parse_xml(r'<w:shd {} w:fill="E6B9B7"/>'.format(nsdecls('w'))))
                p2._tc.get_or_add_tcPr().append(shading_elm)
            elif type == "Detect":
                shading_elm = parse_xml(r'<w:shd {} w:fill="FCD6B4"/>'.format(nsdecls('w')))
                p1 = table.cell(x,0)
                p2 = table.cell(x,1)
                document.tables[number].cell(x,0)._tc.get_or_add_tcPr().append(parse_xml(r'<w:shd {} w:fill="FCD6B4"/>'.format(nsdecls('w'))))
                p2._tc.get_or_add_tcPr().append(shading_elm)
            else: #is recover
                shading_elm = parse_xml(r'<w:shd {} w:fill="ECF2DF"/>'.format(nsdecls('w')))
                p1 = table.cell(x,0)
                p2 = table.cell(x,1)
                document.tables[number].cell(x,0)._tc.get_or_add_tcPr().append(parse_xml(r'<w:shd {} w:fill="ECF2DF"/>'.format(nsdecls('w'))))
                p2._tc.get_or_add_tcPr().append(shading_elm)
    #remove extra rows
    for x in range(len(maj_min), 32): #Change if you add more rows
        row = table.rows[len(maj_min)]
        remove_row(table, row)
    return

def remove_row(table, row):
    tbl = table._tbl
    tr = row._tr
    tbl.remove(tr)
    return

def sort_reccomendations(maj_min):
    maj_min.sort(key = lambda x: x.importance)
    maj_min.sort(key = lambda x: x.rank)
    return

def set_font_all(document, paragraph_num):
    document.paragraph[paragraph_num].font.name = 'IBM Plex Sans Light'
    document.paragraph[paragraph_num].font.size = Pt(9)
    return

def save_images(document, number, pic):
    table = document.tables[number]
    row_cells = table.add_row().cells
    paragraph = row_cells[0].paragraphs[0]
    run = paragraph.add_run()
    if pic == "Randsom.jpg" or pic == "%Ach.jpg":
        table.rows[1].height = Cm(11.5)
        run.add_picture(pic, width = 5400000, height = 4000000)
    if pic == "Spider.jpg" or pic == "NIST.jpg":
        if pic == "Spider.jpg":
            table.rows[1].height = Cm(10.39)
            run.add_picture(pic, width = 5400000, height = 3500000)
        else: 
            table.rows[1].height = Cm(12)
            run.add_picture(pic, width = 5400000, height = 3900000)
    run.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for col in table.columns:
        for cell in col.cells:
            for par in cell.paragraphs:
                par.alignment = WD_ALIGN_PARAGRAPH.CENTER
    return

def develope_images(document):
    images = {"NIST.jpg" : 2, "Spider.jpg" : 1, "%Ach.jpg": 4, "Randsom.jpg" : 3}
    for x, y in images.items():
        save_images(document= document, number = y, pic = x)


def add_footer(document, company):
    """section = document.sections[0]
    footer = section.footer
    footer_para = footer.paragraphs[0]
    footer_para.text = "\tIBM -" + company.company_name + " CONFIDENTIAL"""
    for section in document.sections:
        footer_section = section
        footer = footer_section.footer
        footer_text = footer.paragraphs[0]
        footer_text.text = "\tIBM -" + company.company_name + " CONFIDENTIAL"
        footer.paragraphs[0].runs[0].font.name = 'IBM Plex Sans Light'
        footer.paragraphs[0].runs[0].font.size = Pt(9)
   

    return

def pull_ransom():
    excel_data_df = pandas.read_excel('CRAT_example.xlsx', sheet_name= 'Internal Data')
    return ((excel_data_df.loc[6, 'Ransom_percent'])  * 100) 
   
def change_intro(document, company):
    #access table 0
    count = 0
    table = document.tables[0]
    
    document.tables[0].cell(1,0).text = '        ' + company.company_name
    document.tables[0].cell(1,0).paragraphs[0].runs[0].font.size = Pt(23)
    document.tables[0].cell(1,0).paragraphs[0].runs[0].font.bold = True
    document.tables[0].cell(1,0).paragraphs[0].runs[0].font.color.rgb = RGBColor(255,255,255)
    document.tables[0].cell(1,0).paragraphs[0].runs[0].font.name = 'IBM Plex Sans Light'

    
    document.tables[0].cell(3,0).text = '                        '  + company.date
    document.tables[0].cell(3,0).paragraphs[0].runs[0].font.size = Pt(14)
    document.tables[0].cell(3,0).paragraphs[0].runs[0].font.color.rgb = RGBColor(255,255,255)
    document.tables[0].cell(3,0).paragraphs[0].runs[0].font.name = 'IBM Plex Sans Light'
    
    return




def main():
    print_intro()
    victim_company = get_information()
    document = Document('assessment_template.docx')
    pull_reccomendations()
    sort_reccomendations(major_reccomendation)
    sort_reccomendations(minor_reccomendation)
    final_scores = pull_score()
    generate_table(document, major_reccomendation, 7)
    generate_table(document, minor_reccomendation, 9)
    ransom_percent = pull_ransom()
    replace_all(victim_company, document, ransom_percent, final_scores)
    copy_total_score()
    save_images(document, 4, "Randsom.jpg")
    save_images(document, 2, "Spider.jpg")
    save_images(document, 3, "NIST.jpg")
    save_images(document, 5, "%Ach.jpg")
    add_footer(document, victim_company)
    change_intro(document, victim_company)
    
    document.save('new document.docx')
    return

if __name__ == "__main__":

    main()
