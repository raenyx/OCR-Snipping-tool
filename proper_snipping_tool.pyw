import cv2 as cv
import pytesseract
import os
import pyautogui, time
import numpy
from infi.systray import SysTrayIcon
from PyDictionary import PyDictionary
import win32clipboard
from datetime import date
from io import BytesIO
import PIL as pil

dictionary = PyDictionary()
time.sleep(2)
ref_point = []
crop = False

def send_to_clipboard(image):
    output = BytesIO()
    image.convert('RGB').save(output, 'BMP')
    data = output.getvalue()[14:]
    output.close()

    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
    win32clipboard.CloseClipboard()

def snip(systray):

    def shape_selection(event, x, y, flags, param):
        global ref_point, crop

        if event == cv.EVENT_LBUTTONDOWN:
            ref_point = [(x,y)]
        
        elif event == cv.EVENT_LBUTTONUP:
            ref_point.append((x,y))
            cv.rectangle(image, ref_point[0], ref_point[1], (0, 250, 0), 2)
            cv.imshow("image", image)     


    image = pyautogui.screenshot()
    opencv_image = numpy.array(image)
    image = opencv_image[:, :, ::-1].copy() 
    cv.imshow('image', image)
    clone = image.copy()
    cv.namedWindow("image")
    cv.setMouseCallback("image", shape_selection)

    while True:
        cv.imshow('image', image)
        key = cv.waitKey(1)         

        if key == ord('q'):
            image = clone.copy

        elif key == ord('c'):
            break

    if len(ref_point) == 2:
        crop_image = clone[ref_point[0][1]:ref_point[1][1], ref_point[0][0]:ref_point[1][0]]
        cv.imshow("crop_image", crop_image)
        text = pytesseract.image_to_string(crop_image)
        print(text)

    
    keyy = cv.waitKey()

    if keyy == ord('g'):
        string = fr"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
        os.startfile(string)
        time.sleep(2)
        text = text.replace("\n", " ")
        pyautogui.write(text, interval = 0)
        pyautogui.press('enter')
        cv.waitKey(0)

    elif keyy == ord('d'):
        text = text.split("\n")
        print(text)
        meaning = dictionary.meaning(text[0])
        pyautogui.alert(text = meaning, title = 'Dictionary', button = 'OK')    

    elif keyy == ord('s'):
        today = str(date.today())
        today = fr"D:\Snip\{today}.png"

        cv.imwrite(today, crop_image)

    elif keyy == ord('i'):
        pil_image = cv.cvtColor(crop_image, cv.COLOR_BGR2RGB)
        pil_image = pil.Image.fromarray(pil_image)
        send_to_clipboard(pil_image)

    elif keyy == ord('t'):
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardText(text)
        win32clipboard.CloseClipboard()

    cv.destroyAllWindows()

menu_options = (("Snip", None, snip),)
systray = SysTrayIcon("icon.ico", "Example tray icon", menu_options)
systray.start()
