# -*- coding: utf-8 -*-
"""
Created on Tue Mar 26 2015
@author: Karolina Cibor
"""
from twill.commands import *
from Tkinter import *
from tkFileDialog import askopenfile,askdirectory
import random, os
from xlrd import open_workbook

def onclick():
   pass
# sprawdz czy zawiera duplikaty:
def uniq(lista):
  output = []
  for x in lista:
    if x not in output:
      output.append(x)
  return output

def evaluate(event):
    random_ind=[]
    names=""
    #pobierz liczbe zwyciezcow
    n = int(eval(entry.get()))
    if n > int(len(names_array)):
        print("Liczba zwycięzców mniejsza niż liczba zgłoszeń.")
        n=len(names_array)-1
    #wygeneruj liczby pseudolosowe
    for i in range(n*n):
        random_ind.append(random.randint(1, n))
        random_ind=uniq(random_ind)
    #wygeneruj nazwiska z listy
    for i in range(n):
        names=names+"\n"+str(i+1)+". "+str(names_array[random_ind[i]])
    text = Text(w)
    text.insert(INSERT, "Lista "+str(n)+" studentów: "+names)
    text.pack()
        
if __name__ == "__main__":
    w = Tk()
    # wybierz katalog zawiarajacy formularz:
    dirname = askdirectory(parent=w,initialdir="/",title='Please select a directory')
    if len(dirname ) > 0:
        print "You chose %s" % dirname 
    #wybierz plik z nazwiskami:
    f = askopenfile(parent=w,mode='rb',title='Choose a file')
    book = open_workbook(os.path.basename(f.name), on_demand=True)
    
    
    # wybierz piwerszy arkusz
    first_sheet = book.sheet_by_index(0)
    # odczytaj liczbe wierszy(nazwisk)
    n = int(first_sheet.nrows - 1)
    # stworz tablice przechowujaca nazwiska studentow
    names_array = []
    
    # odkoduj polskie znaki
    for x in xrange(0, n):
        s = str(first_sheet.cell(x+1, 1))[7:-1]
        s=s.decode('unicode-escape')
        names_array.append(s.encode('utf-8'))
        print names_array[x]
    
    Label(w, text="Ilu studentów wylosować?").pack()
    entry = Entry(w)
    entry.bind("<Return>", evaluate)
    entry.pack()
    res = Label(w)
    res.pack()
    w.mainloop()
    root.mainloop()
