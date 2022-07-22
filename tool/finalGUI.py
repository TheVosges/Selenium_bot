from ctypes import alignment
from msilib.schema import Error
import PySimpleGUI as sg
import os

# to import function from Arek's code
from generate import Invoice_Adjustment_Audit_Report as audit
from generate import Invoice_Adjustment_Accrual_Report as accrual
from merge import Services as services
from generate import PDFtoJPG as pdf
from subprocess import call

# Colors
dark_blue=('#2e6fa4')
light_blue=('#7aacc8')

# Fonts
font = ("Bahnschrift SemiLight", 13)
font2 = ("Bahnschrift SemiLight", 10)

#List of countries that you can see 
country_list=["1.Poland", "2.Netherlands", "3.Belgium", "4.Germany", "5.Spain", "6.France", "7.United Kingdom", "8.Italy",
			"Portugal", "Sweden", "Turkiye"]

Selection=['JUN-22','JUL-22','AUG-22','SEP-22','OCT-22','NOV-22','DEC-22','JAN-23','FEB-23','MAR-23','APR-23','MAY-23']

# Theme
sg.theme("Reddit")

# Window layout
layout_1 = [
[sg.Text(" ", size=(5, 1))],
[sg.Image(filename=r'I:\ICS\IC\Automation\16. Reconciliation Tool\tool\logo.png')],
[sg.Text(" ", size=(5, 1))],
[sg.Text("Reconciliation Tool".upper(), size=(20, 1), font=font)],
[sg.Text(" ", size=(5, 1),)],

[sg.Text("Step 1: ", size=(6, 1), font=font2), sg.Text('Date:', font=font2), sg.Combo(Selection, key='period')],
[sg.Text('Step 2: Output folder', font=font2), sg.In(key='path', size=(24, 1)), sg.FolderBrowse(button_color=dark_blue, font=font2)],

[sg.Text("Step 3: ", size=(8, 1), font=font2), sg.Button("Download Accural Report", size=(22, 1), button_color=dark_blue, font=font2)],
[sg.Text("Step 4: ", size=(8, 1), font=font2), sg.Button("Download Audit Report", size=(22, 1), button_color=dark_blue, font=font2)],

[sg.Text("Step 5: ", size=(8, 1), font=font2), sg.Button("PDF/JPG xls/xlsx", size=(22, 1), button_color=dark_blue, font=font2)],
#[sg.Text(" ", size=(5, 1), key='-text-')],
#[sg.Text('Step 3: Choose countries', font=font2)],
[sg.Checkbox('1.Poland', default=False, enable_events=True, key='1.Poland'), sg.Checkbox('2.Netherlands', default=False, enable_events=True,  key='2.Netherlands'), sg.Checkbox('4.Germany', enable_events=True,  default=False, key='4.Germany')],
[sg.Checkbox('5.Spain', default=False, enable_events=True,  key='5.Spain'), sg.Checkbox('6.France', default=False, key='6.France'), sg.Checkbox('7.United Kingdom', default=False, enable_events=True,  key='7.United Kingdom'), sg.Checkbox('8.Italy', default=False, enable_events=True,  key='8.Italy')],
[sg.Checkbox('Portugal', default=False, enable_events=True,  key='Portugal'), sg.Checkbox('Sweden', default=False, key='Sweden'), sg.Checkbox('Turkiye', default=False, enable_events=True,  key='Turkiye'), sg.Checkbox('3.Belgium', default=False, enable_events=True,  key='3.Belgium')],
[sg.Checkbox('All', default=False, enable_events=True, key='All'), sg.Button("Clear all", key='Clear', button_color=light_blue)],
# [sg.Button("Download from Oracle", size=(24, 1), button_color=dark_blue, font=font2), sg.Button("Download from IFS", size=(24, 1), button_color=light_blue, font=font2)],
# [sg.Button("Download from SAP", size=(24, 1), button_color=dark_blue, font=font2), sg.Button("Download from SN", size=(24, 1), button_color=light_blue, font=font2)],
#[sg.Text(" ", size=(5, 1), key='-text-')],
[sg.Text("Step : 6", size=(10, 1), key='-text-'), sg.Button("Merge files", size=(24, 1), button_color=dark_blue, font=font2)],
]

# Second window for data input
# def window2():
#     layout_2 = [
#     [sg.Text(" ", size=(10, 1), key='-text-')],
#     [sg.Text("Please provide data:".upper(), size=(20, 1), key='-text-', font=font)],
#     [sg.Text(" ", size=(10, 1), key='-text-')],
#     [sg.Text('Date:', font=font2), sg.Combo(Selection, key='board')],
#     [sg.Button("Next", size=(19, 1), button_color=dark_blue, font=font2), sg.Exit(size=(19, 1), button_color=light_blue, font=font2)],
#     ]
# # sg.InputText(key='period', default_text="example JUN-22")
#     window = sg.Window("Data input", layout_2, size=(320, 185), element_justification='center')
#     while True:
#         event, value = window.read()
#         if event == sg.WINDOW_CLOSED or event == "Exit":
#             break
#         if event == "Next":
#             # event == "Exit"
#             mes = "Success"
#             try:
#                 accrual.start(value['period'])
#                 audit.start(value['period'])
#                 pdf.start()
#             except Exception as e:
#                 print(e)
#                 mes = e
#             sg.Popup(mes)

#     window.close()


# Initialize start window and define events
window = sg.Window("Reconciliation Tool", layout_1, size=(420, 550),icon=r'I:\ICS\IC\Automation\16. Reconciliation Tool\tool\logo.png', element_justification='center')
while True:
    event, value = window.read()
    mes = "Success"
    if event == sg.WINDOW_CLOSED:
        break
    if event == "Download Audit Report":
        try:
            audit.start(value['period'])  #calling audit module (which was imported above)
        except Exception as e:
            print(e)
            mes = e
            sg.Popup(mes)
    if event == "Download Accural Report":
        try:
            accrual.start(value['period'])
        except Exception as e:
            print(e)
            mes = e
            sg.Popup(mes)
    if event == "PDF/JPG xls/xlsx":
    
        print("PDFtoJPG")
        pdf.start(value['path'])
    if value['All'] == True:
        window['1.Poland'].update(True)
        window['8.Italy'].update(True)
        window['5.Spain'].update(True)
        window['3.Belgium'].update(True)
        window['4.Germany'].update(True)
        window['6.France'].update(True)
        window['Sweden'].update(True)
        window['2.Netherlands'].update(True)
        window['Portugal'].update(True)
        window['Turkiye'].update(True)
        window['7.United Kingdom'].update(True)
    if event == 'Clear':
            window['All'].update(False)
            window['1.Poland'].update(False)
            window['8.Italy'].update(False)
            window['5.Spain'].update(False)
            window['3.Belgium'].update(False)
            window['4.Germany'].update(False)
            window['6.France'].update(False)
            window['Sweden'].update(False)
            window['2.Netherlands'].update(False)
            window['Portugal'].update(False)
            window['Turkiye'].update(False)
            window['7.United Kingdom'].update(False)
    if event == "Merge files":
        list = []
        for item in country_list:    # checking if country on layout was checked and if yes adding it to list
            if value[item] == True:
                list.append(item)
        print(list)
        path = value['path']
        finalPath = path.replace("/", "\\") + "\\"
        if path == "":
            mes = "Please provide the path"
        else:
            mes = services.start(finalPath, list)
        sg.Popup(mes)

window.close()

