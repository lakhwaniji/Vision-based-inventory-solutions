import cv2

def scanqrcode():
    cap=cv2.VideoCapture(0)
    detector=cv2.QRCodeDetector()
    while True:
        _,img=cap.read()
        data,one, _=detector.detectAndDecode(img)
        if data:
            return data
            break
        if cv2.waitKey(1)==ord('q'):
            break
    cv2.destroyAllWindows()
