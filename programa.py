from tkinter import *
import webbrowser
from PIL import ImageTk, Image
import openpyxl
import xlrd
from openpyxl import Workbook
import pathlib

programa = Tk()
programa.geometry("850x450")
pavadinimas = Label(programa, text="Siunčiamos siuntos registravimas", font="ar 10 bold")
programa.title("Siuntų registravimo forma")
programa.configure(highlightbackground="black", highlightthickness=3)


#exelio kurimas ir informacijos issaugojimas i ji
failas = pathlib.Path("Siuntos_registravimas.xlsx")
if failas.exists():
    pass
else:
    failas=Workbook()
    sheet=failas.active
    sheet["A1"]="Siuntėjas"
    sheet["B1"]="Gavėjas"
    sheet["C1"]="Gavėjo adresas"
    sheet["D1"]="Siuntos išmatavimai"
    sheet["E1"]="Siuntos svoris"
    failas.save("Siuntos_registravimas.xlsx")


def issaugoti():
    workbook_name = 'Siuntos_registravimas.xlsx'
    wb = openpyxl.load_workbook(workbook_name)
    page = wb.active
    a=laukas1.get()
    b=laukas2.get()
    c=laukas3.get()
    d=laukas4.get()
    e=laukas5.get()
    row = [a, b, c, d, e]
    page.append(row)
    wb.save(workbook_name)


def callback(url):
    webbrowser.open_new(url)



#Lenteles pagrindiniu laukeliu kurimas
uzrasas1 = Label(programa, text="Siuntėjas", font="ar 8 bold")
laukas1 = Entry(programa)
uzrasas2 = Label(programa, text="Gavėjas", font="ar 8 bold")
laukas2 = Entry(programa)
uzrasas3 = Label(programa, text="Gavėjo adresas", font="ar 8 bold")
laukas3 = Entry(programa)
uzrasas4 = Label(programa, text="Siuntos išmatavimai (cm)", font="ar 8 bold")
laukas4 = Entry(programa)
uzrasas5=Label(programa, text="Siuntos svoris (kg)", font="ar 8 bold")
laukas5= Entry(programa)



uzrasas1.grid(row=1, column=2)
laukas1.grid(row=1, column=3)
uzrasas2.grid(row=2, column=2)
laukas2.grid(row=2, column=3)
uzrasas3.grid(row=3, column=2)
laukas3.grid(row=3, column=3)
uzrasas4.grid(row=4, column=2)
laukas4.grid(row=4, column=3)
uzrasas5.grid(row=5, column=2)
laukas5.grid(row=5, column=3)
pavadinimas.grid(row=0, column=3)


#Registracijos laukelio kurimas
mygtukas = Button(programa, text="Registruoti!", command=issaugoti)
mygtukas.grid(row=6, column=3)

#naudingos nuorodos kurimas
link1 = Label(programa, text="Tai Jums gali būti naudinga - Lietuvos žemėlapis", fg="red", font="ar 8 bold", cursor="dotbox")
link1.bind("<Button-1>", lambda e: callback("https://maps.lt/map/"))
link1.grid(column=1, row=8)
link2 = Label(programa, text="Mus rekomenduoja", fg="red", font="ar 8 bold", cursor="dotbox")
link2.bind("<Button-1>", lambda e: callback("https://codeacademy.lt/"))
link2.grid(column=1, row=9)


#logotipo ikelimas
logo = ImageTk.PhotoImage(Image.open("fast.png"))
panel = Label(programa, image=logo)
panel.grid(column=1, row=0)


programa.mainloop()




