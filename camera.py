import cv2
from PIL import Image
from pytesseract import pytesseract

camera=cv2.VideoCapture(0)

while True:
    _,image=camera.read()
    cv2.imshow('Text detection', image)
    if cv2.waitKey(1)& 0xFF==ord('s'):
        cv2.imwrite('test1.jpg', image)
        break                       #image captured 
camera.release()
cv2.destroyAllWindows()    

def tesseract():
    path_to_tesseract=r"/opt/homebrew/bin/tesseract"
    image_path='test1.jpg'
    pytesseract.tesseract_cmd=path_to_tesseract
    text=pytesseract.image_to_string(Image.open(image_path))
       #convert captured image to string
    print(text[:-1])
tesseract()    


