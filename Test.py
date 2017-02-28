#!/usr/bin/python
# -*- coding: iso-8859-15 -*-
import Tkinter,xlrd, tkMessageBox, sys, write

reload(sys)
sys.setdefaultencoding('utf-8')
#####Добавление данных#####
rb = xlrd.open_workbook('1.xls',formatting_info=True)


##########КОНКУРЕНТЫ##########
#выбираем активный лист
sheetForeign = rb.sheet_by_index(0)

#Множество продуктов конкурентов
ProductsForeign = list()

#получаем список продуктов из всех записей
for rownum in range(sheetForeign.nrows):
  #первая запись - это название магазина
  if rownum !=0:
    ProductsForeign.append(sheetForeign.row_values(rownum)[0])

######НАШ###############
#выбираем активный лист
sheetOur = rb.sheet_by_index(1)

#Множество наших продуктов
ProductsOur = list()

#получаем список продуктов из всех записей
for rownum in range(sheetOur.nrows):
  #первая запись - это название магазина
  if rownum !=0:
    ProductsOur.append(sheetOur.row_values(rownum)[0])

#Удалить повторения
ProductsOur = dict(zip(ProductsOur, ProductsOur)).values()



ListProd=list()

#Ищем наши продукты у конкурентов
for prodour in ProductsOur:
  ListProdFor = list()
  ListProdFor.append(prodour)
  for prodfor in ProductsForeign:
     key = prodour.find(' ')
     #Вытаскиваем название продукта
     NameProd = prodour[0:key]

     #Либо ищем по первому слову продукта
     if (NameProd in prodfor):
       #DictProd[prodour] = prodfor
       ListProdFor.append(prodfor)
     #Либо по полному названию
     elif (prodour in prodfor):
       #DictProd[prodour] = prodfor
       ListProdFor.append(prodfor)
  ListProd.append(ListProdFor)
#####Добавление данных#####

#####ОКНО#####
from Tkinter import *
master = Tk()

#####Выбраные пары будем хранить в словаре#####
ChoiceList=dict()

#####Событие при выборе элемента списка нашего магазина####
curSelectListOur = 0
def onselectProdOur(evt):
    # Note here that Tkinter passes an event object to onselect()
    w = evt.widget
    index = int(w.curselection()[0])
    value = w.get(index)
    global curSelectListOur
    curSelectListOur = value
    print 'You selected item %d: "%s"' % (index, value)
    #Очистим список
    listboxForeign.delete(0, END)
    listboxForeign.insert(END, "Список товаров конкурентов")
    for item in ListProd:
        if item[0] == value:
           iter=0
           for prod in item:
             #Первый элемент списка - наш товар, его не берем
             if iter !=0 :
                listboxForeign.insert(END, prod)
             iter=iter+1
#####Событие при выборе элемента списка нашего магазина####

#####Событие при выборе элемента списка магазина конкурента####
def onselectProdFor(evt):
    # Note here that Tkinter passes an event object to onselect()
    global ChoiceList
    global curSelectListOur
    w = evt.widget
    index = int(w.curselection()[0])
    value = w.get(index)
    #записать выбранную пару
    if index > 0:
        ChoiceList[curSelectListOur] = value

#####Событие при выборе элемента списка магазина конкурента####

listboxOur = Listbox(master, selectmode=SINGLE, width=70)
listboxOur.bind('<<ListboxSelect>>', onselectProdOur)

listboxOur.insert(END, "Список наших товаров")

for item in ProductsOur:
    listboxOur.insert(END, item)
listboxOur.pack(side = 'left')

listboxForeign = Listbox(master, selectmode=SINGLE, width=70)
listboxForeign.bind('<<ListboxSelect>>', onselectProdFor)
listboxForeign.pack()

listboxForeign.insert(END, "Список товаров конкурентов")
listboxForeign.pack(side = 'left')



#Событие - нажали на кнопку
def onBtClick():
    showStr =""
    global ChoiceList
    for keys,values in ChoiceList.items():
        showStr = showStr + str(keys) + " : "
        showStr = showStr + values + "\n"+ "\n"
    tkMessageBox.showinfo('Список выбранных позиций', showStr)

def onBtSaveClick():
     global ChoiceList
     wrExcel=write.WriteExcel()
     wrExcel.WriteDocEx(**ChoiceList);

#Для просмотра списка выбранных товаров
bt = Button(master, text="Просмотреть выбранные позиции", width = 50,command = onBtClick)
bt.pack(side = 'bottom')

btSave = Button(master, text="Сохранить", width = 50,command = onBtSaveClick)
btSave.pack(side = 'bottom')

mainloop()
