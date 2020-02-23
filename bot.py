from pynput.mouse import Listener
import logging
import pandas as pd
import pyautogui
from openpyxl import load_workbook
from time import sleep
from tkinter import *



def monitor_myszy():
    # filemode - w - nadpisuje plik od początku
    logging.basicConfig(filename="mysz.txt",filemode="w" ,level=logging.DEBUG, format="%(message)s")
    def on_click(x,y, button, pressed):
        if pressed:
            logging.info("klik \t{0}\t {1}\t{2}".format(x, y, button))

    def on_scroll(x, y):
        logging.info("scroll \t{0}\t {1}".format(x, y))

    with Listener(on_click=on_click, on_scroll=on_scroll) as listener:
        listener.join()

def konwersja_logu_do_xls():
    log = pd.read_table("mysz.txt")
    log.to_excel("mysz.xlsx", index=False, header=True)

def dodanie_wiersza_tytulowego():
    skoroszyt = load_workbook("mysz.xlsx")
    arkusz = skoroszyt.active
    arkusz.insert_rows(1)
    arkusz.cell(1,1,"KolumnaA")
    arkusz.cell(1,2,"KolumnaB")
    arkusz.cell(1,3,"KolumnaC")
    arkusz.cell(1,4,"KolumnaD")
    skoroszyt.save("mysz.xlsx")

def odczyt_x():
    x = pd.read_excel("mysz.xlsx")
    licznik = 0
    ilosc_wierszy = x["KolumnaC"].size
    tablicax = []
    while licznik < ilosc_wierszy:
        asd = x["KolumnaC"].loc[licznik]
        licznik += 1
        tablicax.append(asd)
    return tablicax

def odczyt_y():
    y = pd.read_excel("mysz.xlsx")
    licznik = 0
    ilosc_wierszy = y["KolumnaB"].size
    tablicay = []
    while licznik < ilosc_wierszy:
        asd = y["KolumnaB"].loc[licznik]
        licznik += 1
        tablicay.append(asd)
    return tablicay

def odczyt_przycisku():
    od = pd.read_excel("mysz.xlsx")
    licznik = 0
    ilosc_wierszy = od["KolumnaD"].size
    tablicaod = []
    while licznik < ilosc_wierszy:
        asd = od["KolumnaD"].loc[licznik]
        licznik += 1
        string_przycisku = str(asd)
        string_przycisku = string_przycisku.replace("Button.left", "left")
        string_przycisku = string_przycisku.replace("Button.right", "right")
        tablicaod.append(string_przycisku)

    return tablicaod

# funkcja korzysta z odczytx i odczyty

def start_bota():
    konwersja_logu_do_xls()
    dodanie_wiersza_tytulowego()
    licznikx = 0
    liczniky = 0
    licznikod = 0
    licznik_wykonania = pd.read_excel("mysz.xlsx")["KolumnaB"].size-1
    while licznik_wykonania > ile_razy_wykonac:
        while licznik_wykonania >= liczniky:
            y = (odczyt_y()[liczniky])
            x = (odczyt_x()[licznikx])
            button = (odczyt_przycisku()[licznikod])
            pyautogui.click(x,y, button=button)
            licznikx += 1
            liczniky += 1
            sleep(2)


window = Tk()
window.title("Rejestrator bota")
opcja1 = Button(window, text="Rejestrator bota", activeforeground= "#FF0000", command=monitor_myszy).grid(sticky=E+W, padx = 5, pady = 5)
opcja2 = Button(window,text="Uruchom bota", activeforeground= "#FF0000", command=start_bota).grid(sticky=E+W, padx = 5, pady = 5)
tytul = Label(window, text="Wybierz ile razy zostanie powtórzone zadanie:").grid(sticky=E+W, padx=5, pady=5)
ilosc_powtorzen = Spinbox(window, from_=1, to=20).grid(sticky=E + W, padx=5, pady=5)
opcja3 = Button(window, text="Zakończ").grid(sticky=E+W, padx = 5, pady = 5)
ilosc_powtorzen = IntVar()
ile_razy_wykonac = ilosc_powtorzen.get()
window.mainloop()

