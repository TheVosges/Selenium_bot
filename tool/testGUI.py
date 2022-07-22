from ctypes import alignment
from msilib.schema import Error
import PySimpleGUI as sg
import os

from generate import Invoice_Adjustment_Audit_Report as audit
from generate import InvoiceAdjustmentAccrualReport as accrual
from merge import Services as services
# from generate import PDFtoJPG as pdf
from subprocess import call

# Colors
dark_blue=('#2e6fa4')
light_blue=('#7aacc8')

# Fonts
font = ("Bahnschrift SemiLight", 13)
font2 = ("Bahnschrift SemiLight", 10)

country_list=["Poland", "Netherlands", "Belgium", "Germany", "Spain", "France", "UK", "Italy",
			"Portugal", "Sweden", "Turkiye"]

# Theme
sg.theme("Reddit")

# Window layout
layout_1 = [
[sg.Text(" ", size=(10, 1), key='-text-')],
[sg.Image(filename="logo.png")],
[sg.Text(" ", size=(10, 1), key='-text-')],
[sg.Text("Reconciliation Tool".upper(), size=(20, 1), key='-text-', font=font)],
[sg.Text(" ", size=(10, 1), key='-text-')],
[sg.Text("STEP 1: ", size=(10, 1), key='-text-'), sg.Button("Download from SN", size=(24, 1), button_color=dark_blue, font=font2)],
[sg.Text(" ", size=(10, 1), key='-text-')],
[sg.Text('STEP 2: Output folder', font=font2), sg.In(key='path', size=(24, 1)), sg.FolderBrowse(button_color=dark_blue, font=font2)],
[sg.Text(" ", size=(10, 1), key='-text-')],
[sg.Text('STEP 3: Choose countries', font=font2)],
[sg.Checkbox('Poland', default=False, enable_events=True, key='Poland'), sg.Checkbox('Netherlands', default=False, enable_events=True,  key='Netherlands'), sg.Checkbox('Germany', enable_events=True,  default=False, key='Germany')],
[sg.Checkbox('Spain', default=False, enable_events=True,  key='Spain'), sg.Checkbox('France', default=False, key='France'), sg.Checkbox('UK', default=False, enable_events=True,  key='UK'), sg.Checkbox('Italy', default=False, enable_events=True,  key='Italy')],
[sg.Checkbox('Portugal', default=False, enable_events=True,  key='Portugal'), sg.Checkbox('Sweden', default=False, key='Sweden'), sg.Checkbox('Turkiye', default=False, enable_events=True,  key='Turkiye'), sg.Checkbox('Belgium', default=False, enable_events=True,  key='Belgium')],
[sg.Checkbox('All', default=False, enable_events=True, key='All')],
# [sg.Button("Download from Oracle", size=(24, 1), button_color=dark_blue, font=font2), sg.Button("Download from IFS", size=(24, 1), button_color=light_blue, font=font2)],
# [sg.Button("Download from SAP", size=(24, 1), button_color=dark_blue, font=font2), sg.Button("Download from SN", size=(24, 1), button_color=light_blue, font=font2)],
[sg.Text(" ", size=(10, 1), key='-text-')],
[sg.Text("STEP 4: ", size=(10, 1), key='-text-'), sg.Button("Merge files", size=(24, 1), button_color=dark_blue, font=font2)],
]

# Second window for data input
def window2():
    layout_2 = [
    [sg.Text(" ", size=(10, 1), key='-text-')],
    [sg.Text("Please provide data:".upper(), size=(20, 1), key='-text-', font=font)],
    [sg.Text(" ", size=(10, 1), key='-text-')],
    [sg.Text('ID:     ', font=font2), sg.InputText(key='id')],
    [sg.Text('Date:', font=font2), sg.InputText(key='period')],
    [sg.Button("Next", size=(19, 1), button_color=dark_blue, font=font2), sg.Exit(size=(19, 1), button_color=light_blue, font=font2)],
    ]

    window = sg.Window("Data input", layout_2, size=(320, 185), element_justification='center')
    while True:
        event, value = window.read()
        if event == sg.WINDOW_CLOSED or event == "Exit":
            break
        if event == "Next":
            # event == "Exit"
            accrual.start(value['id'])
            audit.start(value['period'], value['id'])
            try:
                accrual.start(value['id'])
                audit.start(value['period'], value['id'])
            except Error:
                print("Error")

    window.close()


# Initialize start window and define events
window = sg.Window("Start", layout_1, size=(420, 570), element_justification='center')
while True:
    event, value = window.read()
    if event == sg.WINDOW_CLOSED:
        break
    if event == "Download from SN":
        # sg.Popup("Error")
        window2()
        if value['All'] == True:
            window['Poland'].update(True)
            window['Italy'].update(True)
            window['Spain'].update(True)
            window['Belgium'].update(True)
            window['Germany'].update(True)
            window['France'].update(True)
            window['Sweden'].update(True)
            window['Netherlands'].update(True)
            window['Portugal'].update(True)
            window['Turkiye'].update(True)
            window['UK'].update(True)
        if value['All'] == False:
            window['Poland'].update(False)
            window['Italy'].update(False)
            window['Spain'].update(False)
            window['Belgium'].update(False)
            window['Germany'].update(False)
            window['France'].update(False)
            window['Sweden'].update(False)
            window['Netherlands'].update(False)
            window['Portugal'].update(False)
            window['Turkiye'].update(False)
            window['UK'].update(False)
    if event == "Merge files":
        list = []
        for item in country_list:
            if value[item] == True:
                list.append(item)
        print(list)
        path = value['path']
        finalPath = path.replace("/", "\\") + "\\"
        mes = services.start(finalPath, list)
        sg.Popup(mes)


window.close()



