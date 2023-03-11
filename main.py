import serial
import time

device = 'COM6'  # this will have to be changed to the serial port you are using
try:
    print("Trying...", device)
    arduino=serial.Serial(device)
except ValueError:
    print("Failed to connect on", device)
while True:
    time.sleep(1)
    try:
        data = arduino.readline()
        print(data)
        pieces = data.split(" ")
    except ValueError:
        print("Processing")
