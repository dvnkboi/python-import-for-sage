import keyboard as kb
import threading
from pynput.keyboard import Key, Controller
import time as tm
import pyautogui as pg
import pandas as pd

# read excel file
df = pd.read_excel('esmag.xlsx', header=0)

df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
df = df.rename(columns=lambda x: x.strip().replace(' ', '').upper())

ids = df["NUM.BON"]
df[ids.isin(ids[ids.duplicated()])].sort_values("NUM.BON")

# duplicates = set()
# for index, row in df[ids.isin(ids[ids.duplicated()])].sort_values("NUM.BON").iterrows():
#     duplicates.add(row["NUM.BON"])

keyboard = Controller()

origin = pg.position()
delay = 0.1


def mouseClick(x, y):
    tm.sleep(delay)
    pg.moveTo(x, y)
    pg.click()
    tm.sleep(delay)


def mouseMove(x, y):
    tm.sleep(delay)
    pg.moveTo(x, y)
    tm.sleep(delay)


def setOrigin():
    global origin
    origin = pg.position()


def moveToOrigin():
    global origin
    tm.sleep(delay)
    pg.moveTo(origin)
    tm.sleep(delay)


def pressTab():
    tm.sleep(delay)
    keyboard.tap(Key.tab)
    tm.sleep(delay)


def pasteText(text):
    tm.sleep(delay)
    keyboard.type(text)
    tm.sleep(delay)


def pressEnter():
    tm.sleep(delay)
    keyboard.tap(Key.enter)
    tm.sleep(delay)


def removeAllText():
    tm.sleep(delay)
    keyboard.press(Key.ctrl)
    keyboard.tap("a")
    keyboard.release(Key.ctrl)
    keyboard.tap(Key.backspace)
    tm.sleep(delay)


stopEvent = threading.Event()


def stop():
    stopEvent.set()


kb.add_hotkey("ctrl+alt+shift+s", stop)


def write():
    stopEvent.clear()
    mouseClick(1132, 1053)

    for index, row in df.head(1).iterrows():

        if stopEvent.is_set():
            break
        mouseClick(175, 158)
        mouseClick(973, 208)
        removeAllText()

        if stopEvent.is_set():
            break
        pasteText(str(row['SITE']))
        pressTab()

        if stopEvent.is_set():
            break
        pasteText(str(row['NUM.BON']))
        pressTab()

        if stopEvent.is_set():
            break
        pasteText(str(row['DATEBON'].strftime('%d/%m/%Y')))
        pressTab()
        pressTab()
        pressTab()

        if stopEvent.is_set():
            break
        pasteText(str(row['CODEART.']))
        pressTab()
        pressTab()
        pressTab()

        if stopEvent.is_set():
            break
        pasteText(str(row['QUANTITE']))
        pressTab()
        pressTab()
        pressTab()

        if stopEvent.is_set():
            break
        pasteText(str(row['ENT']))
        pressTab()

        if stopEvent.is_set():
            break
        pasteText(str(row['PARCAUTO']))
        pressEnter()
        pressEnter()

        if stopEvent.is_set():
            break
        tm.sleep(1.5)


write()
