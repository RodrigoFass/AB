################################
#                              #
#             BOT              #
#            VENDER            #
#                              #
################################

#region Import e misc

import win32api
import win32con
from PIL import ImageGrab
import time
import pytesseract
from pynput.keyboard import Controller

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
keyboard = Controller()

#endregion

#region Funçõesen


def clicar(x,y):

    time.sleep(0.1)
    win32api.SetCursorPos((x,y))
    time.sleep(0.1)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, x,y, 0, 0)
    time.sleep(0.1)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, x,y, 0, 0)
    return True


#endregion

#region Coordenadas

vnd_1 = {'x':1275,'y':435}
vnd_2 = {'x':880,'y':735}

srch_bar = {'x':820,'y':400}
pric_bar = {'x':655,'y':640}
clse_bar = {'x':937,'y':310}

normal = {'x':820,'y':430}

pric_dcrs = {'x':560,'y':640}

tt = (80,530,128,542)
teste = [970,440,1050,460]
preco = [1020,370,1090,390]

#endregion

#region Main

time.sleep(1)   

while True:

    clicar(vnd_1['x'],vnd_1['y'])
    clicar(srch_bar['x'],srch_bar['y'])
    clicar(normal['x'],normal['y'])

    time.sleep(0.3)
    ScrnSht = ImageGrab.grab(preco)
    valor = pytesseract.image_to_string(ScrnSht, config='--psm 7')

    clicar(clse_bar['x'], clse_bar['y'])
    clicar(vnd_1['x'],vnd_1['y'])
    clicar(pric_bar['x'],pric_bar['y'])

    keyboard.type(valor)

    clicar(pric_dcrs['x'],pric_dcrs['y'])
    clicar(vnd_2['x'],vnd_2['y'])
    clicar(clse_bar['x'], clse_bar['y'])

    time.sleep(0.1)


#endregion