import cv2

from PIL import Image

from pytesseract import pytesseract

camera = cv2.VideoCapture(0)

while True:
    _,Image=camera.read()
    cv2.imshow('Text Detection',Image)
    if cv2.waitKey(1)& 0xFF==ord('s'):
        cv2.imwrite('test1.jpg',Image)
        break
camera.release()
cv2.destroyAllWindows()
