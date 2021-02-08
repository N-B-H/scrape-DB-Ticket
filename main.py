import os
import pdfplumber
from openpyxl import Workbook
from datetime import datetime
import re


file_names = os.listdir("tickets")
file_paths = [f"tickets\\{file_name}" for file_name in file_names]

#file_name = file_names[0]
#file_path = f"tickets\\{file_name}"


def pdf_to_data(file_path):

    ### pdfplumber ###
    pdf = pdfplumber.open(file_path)
    page = pdf.pages[0]  # only first page

    ### page.extract_text() gives bugged results. Therefore going for the table data

    tables = page.extract_tables() #extract tables

    text = page.extract_text()  # used for international tickets
    last_line = text.splitlines()[-1]  # to extract the Ticket ID

    if tables[0][0][0].startswith('Positionen'): international = True


    str_lines = []

    for table in tables:
        for row in table:
            for cell in row:
                if cell: str_lines.extend(cell.split("\n"))

    ticket_string = '\n'.join(str_lines+[last_line])

    #print(ticket_string)

    ### parse all tickets from ticket_string ###

    # find items with regular expressions
    # dict with attribute as key and re-marker as value
    re_marker = {
        "date": r'(?:(?:Gültigkeit: a.)|(?:Fahrtantritt a.)) (.*)\n',
        "start": 'Hinfahrt: (.*)   ',

        "dest": r'Hinfahrt: .+   (.+?)[,\n]',

        "cost": 'Betrag (.*)€',
        "ID": r'\n(.+?) Seite 1 / 1$'
    }

    re_marker_international = r'Gütigkeit: a.(.+).+VON ->NACH XX\n.* (.+) ->(.+)' ### work in progress...

    travel = {}  # will have the same keys as the regular expression markers
    for attribute in re_marker:
        m = re.search(re_marker[attribute], ticket_string)
        try:
            travel[attribute] = m.group(1)
        except:
            print(travel)
            print(attribute)
            print(ticket_string)
            print("\n\n", page.extract_text())
    #print(ticket_string)
    # convert date and cost
    travel["date"] = str(datetime.strptime(travel["date"], '%d.%m.%Y').date())
    travel["cost"] = float(travel["cost"].replace(',', '.'))

    #print(tickets)
    return travel

### read all tickets in /tickets folder

tickets = [pdf_to_data(file_path) for file_path in file_paths]
for ticket in tickets: print(ticket)

### save to an excel-file
def tickets_to_excel(tickets, xls_file_path):

    wb = Workbook()
    sheet = wb.active

    #initialise the worksheet

    sheet.title = "DB-Reisen"

    for y in range(len(tickets)): # as many rows as there are tickets
        #for x in range(1,5): # 4 attributes
        sheet.cell(row=y+1, column=1).value = tickets[y]["date"]

        sheet.cell(row=y+1, column=2).value = f'{tickets[y]["start"]} - {tickets[y]["dest"]}'

        sheet.cell(row=y+1, column=3).value = tickets[y]["cost"]
        #sheet.cell(row=y, column=3).style = 'Currency'
        sheet.cell(row=y+1, column=3).number_format = '#,##0.00€'

        sheet.cell(row=y + 1, column=4).value = tickets[y]["ID"]

    ### manual cell size adjustment not yet supported


    ### autosize not yet supported
    '''
    for column_cells in sheet.columns:
    sheet.column_dimensions[column_cells[0].column].width = max(max_length, cur_length)

        cur_length = sheet.column_dimensions[column_cells[0].column].width
        max_length = max(len(str(cell.value)) for cell in column_cells)

        #adjust
    '''

    wb.save(xls_file_path)

### sort all tickets by date
tickets = sorted(tickets, key=lambda x: x['date'], reverse=True)

tickets_to_excel(tickets, 'output.xlsx')
