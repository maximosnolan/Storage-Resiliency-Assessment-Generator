from docx import Document
from numpy import NaN
import pandas
import csv
import PIL
from PIL import ImageGrab, Image
from docx.shared import Cm, Pt
import os
import sys
from copy import deepcopy
from pandas.io import excel
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
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

def replace_all(company, document):
    document.add_paragraph(company.company_name)
    for paragraph in document.paragraphs:
        #print(paragraph.text)
        if 'TOTAL_SCORE_REPLACE' in paragraph.text:
            paragraph.text = ''
            df = copy_total_score()
            document.tables[0].cell(1,1).text = str(round(df.loc[2, 'SCORE'], 2))
            print(df)
            print(document.tables[0].cell(10, 2).text)
            for x in range (4,35):
                if x != 10 and x != 17 and x != 22 and x != 29 and x!= 35:
                   print(x)
                   document.tables[0].cell(x - 1, 1).text = str(round(df.loc[x, 'SCORE'], 2))
                   document.tables[0].cell(x - 1, 1).allignment = 1
                   document.tables[0].cell(x - 1, 2).text = str(df.loc[x, 'LEVEL'])
                   document.tables[0].cell(x - 1, 2).allignment = 1
                
        if 'CUSTOMER_NAME' in paragraph.text:
            orig_text = paragraph.text
            new_text = str.replace(orig_text, "CUSTOMER_NAME", company.company_name)
            paragraph.text = new_text
        if 'PRESENTATION_DATE' in paragraph.text:
            orig_text = paragraph.text
            new_text = str.replace(orig_text, "PRESENTATION_DATE", company.date)
            paragraph.text = new_text
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    if 'CUSTOMER_NAME' in paragraph.text:
                        paragraph.text = paragraph.text.replace("CUSTOMER_NAME", company.company_name)
                    if 'PRESENTATION_DATE' in paragraph.text:
                        paragraph.text = paragraph.text.replace("PRESENTATION_DATE", company.date)

    
    return

def copy_total_score():
    excel_data_df = pandas.read_excel('CRAT_example.xlsx', sheet_name= 'Results')
    print(excel_data_df)

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
            print(a.rank)
            major_reccomendation.append(a)
        else:
            a = Recomendation(str(excel_data_df.loc[x, 'Count']), str(excel_data_df.loc[x, 'Questions']), str(excel_data_df.loc[x, 'Score']), level)
            print(a.rank)
            minor_reccomendation.append(a)
        #print(excel_data_df.loc[x, 'Questions'])

    return

def pull_score():
    scores = {}
    excel_data_df = pandas.read_excel('CRAT_example.xlsx', sheet_name= 'Internal Data')
    for x in range (1, 19):
        scores[excel_data_df.loc[x, 'Category']] = excel_data_df.loc[x, 'Percent_Ach']
    final_scores = {k: v for k, v in sorted(scores.items(), key=lambda item: item[1])}
    #for x in final_scores.items():
        #print(x)
    return

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
    for x in range(len(maj_min), 32):
        row = table.rows[len(maj_min)]
        remove_row(table, row)
    return

def remove_row(table, row):
    tbl = table._tbl
    tr = row._tr
    tbl.remove(tr)

def sort_reccomendations(maj_min):
    maj_min.sort(key = lambda x: x.importance)
    maj_min.sort(key = lambda x: x.rank)

def set_font_all(document, paragraph_num):
    document.paragraph[paragraph_num].font.name = 'IBM Plex Sans Light'
    document.paragraph[paragraph_num].font.size = Pt(9)
def main():
    print_intro()
    victim_company = get_information()
    document = Document('assessment_template.docx')
    pull_reccomendations()
    sort_reccomendations(major_reccomendation)
    sort_reccomendations(minor_reccomendation)
    for x in major_reccomendation:
        print(x.importance, x.recommendation)
    pull_score()
    generate_table(document, major_reccomendation, 5)
    generate_table(document, minor_reccomendation, 7)
    set_font_all(document, 4)
    
    replace_all(victim_company, document)
    copy_total_score()
    document.save('new document.docx')
    return

if __name__ == "__main__":

    main()