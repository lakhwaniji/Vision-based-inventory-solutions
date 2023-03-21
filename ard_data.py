import serial


def scanning_card():
    arduino = serial.Serial(port='<YOUR CONNECTED PORT TO ARDUINO>', baudrate=9600)
    try:
        data = arduino.readline()
        data = str(data)
        data = data[3:-5]
        return data
    except:
        return "Error In Reading"
