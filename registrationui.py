import os.path

import PySimpleGUI as sg
import time
import ard_data
import Linktoexcel

clock = sg.Text(time.strftime("%d %b %Y"), key="clock")
sg.theme('BluePurple')
label_1 = sg.Text("Enter Your Name--:")
label_2 = sg.Text("Registration Number --:")
card_no = sg.Text("", text_color="Red", key="card_no")
result = sg.Text("", text_color="Green", key="result")
input_1 = sg.InputText("", key="name", tooltip="Enter your Name")
input_2 = sg.InputText("", key="regno", tooltip="Enter your Registration number")
label_3 = sg.Radio("IOT SEC A", "sec", default=False, key="IOT_SEC_A")
label_4 = sg.Radio("IOT SEC B", "sec", default=False, key="IOT_SEC_B")
label_5 = sg.Radio("IOT SEC C", "sec", default=False, key="IOT_SEC_C")
label_6 = sg.Radio("IOT SEC D", "sec", default=False, key="IOT_SEC_D")
button_1 = sg.Button("Submit", size=15, key="submit")
button_2 = sg.Button("Scan Your Card", key="scan", size=15)
col1 = sg.Column([[label_1], [label_2]])
col2 = sg.Column([[input_1], [input_2]])
window = sg.Window("Registration", layout=[[col1, col2],
                                           [label_3, label_4, label_5, label_6],
                                           [button_2, card_no],
                                           [button_1],
                                           [result]],
                   font=('Helvetica', 15))
while True:
    event, value = window.read(timeout=1000)
    match event:
        case "scan":
            window["card_no"].update(value="Scanning Activated")
            data = ard_data.scanning_card()
            window["card_no"].update(value=data)
        case "submit":
            try:
                if value["IOT_SEC_A"]:
                    if not os.path.exists("IOT_SEC_A.xlsx"):
                        Linktoexcel.createfile("IOT_SEC_A.xlsx")
                    result = Linktoexcel.writedata("IOT_SEC_A.xlsx", value['name'], value['regno'],"Love")
                    if result == 'Success':
                        window["result"].update(value="Data Registered Successfully")
                    else:
                        window["result"].update(text_color="Red", value="Error in registering Data")
            except:
                window["result"].update(text_color="Red", value="Error in registering Data")
window.close()
