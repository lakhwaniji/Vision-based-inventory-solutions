import serial
import time

device = "COM4"


def scanning_card():
    try:
        print("Trying ...", device)
        arduino = serial.Serial(9600)
    except:
        print("Failed To Connect")
    try:
        data = arduino.readline()
        data = str(data)
        data = data[3:-5]
        return data
    except:
        return "Error In Reading"
