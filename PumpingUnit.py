# -*- coding: CP1251 -*-
import os
import codecs
import subprocess
from math import *

class SWmodel(object):
    """Клас моделі SolidWorks"""
    d={} # словник параметрів
    fileName=None # ім'я файлу моделі 
    fileType=".SLDPRT" # розширення файлу моделі (".SLDPRT", ".SLDASM")
    
    def create(self):
        """Розраховує параметри моделі"""
        pass
    
    def rebuildModel(self):
        """Перебудовує файл рівнянь та модель SolidWorks"""
        self.write_dict_to_SW_equations()
        self.rebuildAndSaveModel()
        
    def rebuildAndSaveModel(self):
        """Перебудовує та зберігає SolidWorks модель шляхом виконання VBS скрипта"""
        
        vbs=r"""'Скрипт VBS для перебудови моделі SolidWorks 
Dim swApp 'SldWorks.SldWorks
Dim Part 'SldWorks.ModelDoc2
Set swApp = CreateObject("SldWorks.Application")
Set Part = swApp.OpenDoc("{fullFileNameExt}", {docType})
Set Part = swApp.ActivateDoc("{fileNameExt}")
Part.EditRebuild
Part.SaveSilent
Set Part = Nothing
swApp.CloseDoc "{fileName}"
Set swApp=Nothing
        """
        fileName=self.fileName
        fileNameExt=self.fileName+self.fileType
        fullFileNameExt=os.path.join(os.getcwd(), fileNameExt)
        if self.fileType==".SLDPRT":
            docType="1" # якщо деталь
        else:
            docType="2" # якщо зборка
        vbs=vbs.format(fullFileNameExt=fullFileNameExt, fileNameExt=fileNameExt, fileName=fileName, docType=docType)
        scriptFileName=os.path.join(os.getcwd(), "RebuildSWmodelTemp.vbs")
        f=open(scriptFileName, 'w')
        f.write(vbs)
        f.close()
        # виконує процес та чекає його завершення
        subprocess.Popen(r'c:\Windows\system32\wscript.exe '+scriptFileName).wait()
    
    def read_dict_from_SW_equations(self):
        """Додає елементи в словник self.d
        з текстового файлу (utf-8-sig) рівнянь SolidWorks"""
        filename=self.fileName+".txt" # файл рівнянь SolidWorks
        if not os.path.exists(filename): return # якщо файлу не існує, вийти
        f=codecs.open(filename,'r','utf-8-sig') # відкити файл для читання
        for line in f.readlines(): # для усіх рядків у списку
            if '=' in line: # якщо в рядку є символ "="
                pair=line.split('=') # розділити рядок
                pair=pair[0].strip()[1:-1], pair[1].strip() # видалити пробіли і лапки
                
                if pair[1].isdigit(): # якщо всі символи цифри
                    val=int(pair[1])
                else: # інакше
                    try: # якщо дійсне число
                        val=float(pair[1])  
                    except ValueError: # якщо рядок
                        val=pair[1]
                        
                self.d[pair[0].encode('CP1251')]=val # записати в словник
        f.close() # закрити файл
        
    def write_dict_to_SW_equations(self):
        """Записує значення елементів словника self.d
        в текстовий файл (utf-8-sig) рівнянь SolidWorks"""
        d={} # словник з unicode ключами
        for k in self.d: # перетворюємо ключі в unicode
            d[k.decode('CP1251')]=self.d[k]
            
        filename=self.fileName+".txt" # файл рівнянь SolidWorks
        if not os.path.exists(filename): return # якщо файлу не існує, вийти
        f=codecs.open(filename,'r','utf-8-sig') # відкити файл для читання
        oldlines=f.readlines() # старий список рядків
        newlines=[] # новий список рядків
        for line in oldlines: # для усіх рядків у списку
            if '=' in line: # якщо в рядку є символ "="
                pair=line.split('=') # розділити рядок
                pair=pair[0].strip()[1:-1],pair[1].strip() # видалити пробіли і лапки
                if pair[0] in d.keys(): # якщо ліва частина від "=" є серед ключів словника
                    line='"'+pair[0]+'" = '+str(d[pair[0]])+"\n" # формувати новий рядок
            newlines.append(line) # добавити рядок в новий список рядків
        f.close() # закрити файл
        
        f=codecs.open(filename,'w','utf-8-sig') # відкрити файл для запису
        f.writelines(newlines) # записати список нових рядків
        f.close() # закрити файл
        
    def assignParam(self, profil, name, nameList):
        """Присвоює значення заданим елементам словника параметрів.
        profil - конструктор відповідного профілю (див. класи профілів),
        name - назва профілю,
        nameList - список назв елементів SolidWorks.
        """
        p=profil()
        p.create(name=name)
        self.d=p.setSWParam(self.d, nameList)
            
class SWmodelPRT(SWmodel):
    """Клас моделі деталі SolidWorks"""
    fileType=".SLDPRT"

class SWmodelASM(SWmodel):
    """Клас моделі зборки SolidWorks"""
    fileType=".SLDASM"      

class PumpingUnit(SWmodelASM):
    """Клас верстата-гойдалки"""
    fileName="Верстат"
    d={
    "Тип" : "СКД8-3-4000",
    "Максимальне допустиме навантаження" : 80.0,
    "Список довжин ходу полірованого штока" : [1200.0, 1600.0, 2000.0, 2500.0, 3000.0],
    "Довжина ходу полірованого штока" : 2000.0,
    "Список кількості гойдань" : [5.0,12.0],
    "Кількість гойдань" : 5.0,
    "Максимальний крутний момент" : 40.0,
    "Довжина переднього плеча балансира" : 2290.0,
    "Довжина заднього плеча балансира" : 2000.0,
    "Довжина шатуна" : 3000.0,
    "Найбільший радіус кривошипа" : 1290.0,
    "Радіус кривошипа" : None,
    "Горизонтальна відстань між осями опори балансира і тихохідного валу редуктора" : 1345.0,
    "Вертикальна відстань між осями опори балансира і тихохідного валу редуктора" : None,
    "Довжина" : 6900.0,
    "Ширина" : 2250.0,
    "Висота" : 4910.0,
    "Маса" : 11780.0,
    "Систама урівноважування" : "кривошипна",
    "Вага кривошипних противаг" : 750.0,
    "Максимальна кількість кривошипних противаг" : 6,
    "Номінальна потужність електродвигуна" : 18.5,
    "Редуктор" : "Ц2НШ-750 Б"}
    
    def create(self):
        """Розраховує параметри моделі"""
        sList=self.d["Список довжин ходу полірованого штока"]
        s0=self.d["Довжина ходу полірованого штока"]
        # розрахунок вертикальної відстані між осями опори балансира і тихохідного валу редуктора
        smax=max(sList) # найбільша довжина ходу
        rmax=self.d["Найбільший радіус кривошипа"]
        l=self.d["Довжина шатуна"]
        k=self.d["Довжина заднього плеча балансира"]
        k1=self.d["Довжина переднього плеча балансира"]
        l1=self.d["Горизонтальна відстань між осями опори балансира і тихохідного валу редуктора"]
        h=smax*k/k1 # вертикальний хід опори траверси
        b=sqrt(k**2-(h/2)**2) # відстань від опори балансира до вертикалі h
        H=sqrt((l-rmax)**2-(b-l1)**2)+h/2
        self.d["Вертикальна відстань між осями опори балансира і тихохідного валу редуктора"]=H
        
        # розрахунок радіуса кривошипа
        h=s0*k/k1 # вертикальний хід опори траверси
        b=sqrt(k**2-(h/2)**2) # відстань від опори балансира до вертикалі h
        r=l-sqrt((H-h/2)**2+(b-l1)**2)
        self.d["Радіус кривошипа"]=r
        
        # створення та розрахунок деталей
        self.GolovBalansir=GolovBalansir()
        self.GolovBalansir.create(paramPU=self.d)       
        self.Balansir=Balansir()
        self.Balansir.create(paramPU=self.d, paramGolovBalansir=self.GolovBalansir.d)
        self.GolovBalansir.d["Висота між опорами@Профіль головки"]=self.Balansir.d["Висота@Двотавр профіль"]
        self.Shatun=Shatun()
        self.Shatun.create(paramPU=self.d)
        self.Krivoshyp=Krivoshyp()
        self.Krivoshyp.create(paramPU=self.d)
        self.Traversa=Traversa()
        self.Traversa.create(paramPU=self.d)
        self.Val=Val()
        self.Val.create(paramPU=self.d, paramShatun=self.Shatun.d, paramTraversa=self.Traversa.d)
        self.Reduktor=Reduktor()
        self.Reduktor.create()
        self.Stiyka=Stiyka()
        self.Stiyka.create(paramPU=self.d, paramReduktor=self.Reduktor.d, paramVal=self.Val.d, paramKrivoshyp=self.Krivoshyp.d)
        self.Rama=Rama()
        self.Rama.create(paramPU=self.d, paramStiyka=self.Stiyka.d, paramReduktor=self.Reduktor.d)
        self.Protyvaga=Protyvaga()
        self.Protyvaga.create(paramKrivoshyp=self.Krivoshyp.d)
        # зборки:
        self.BalansirVZbori=BalansirVZbori()
        self.Krivoshypy=Krivoshypy()
        self.TraversaVZbori=TraversaVZbori()
        self.Karkas=Karkas()
        
    def rebuildModel(self):
        self.GolovBalansir.rebuildModel()
        self.Balansir.rebuildModel()
        self.Shatun.rebuildModel()
        self.Krivoshyp.rebuildModel()
        self.Traversa.rebuildModel()
        self.Val.rebuildModel()
        self.Reduktor.rebuildModel()
        self.Stiyka.rebuildModel()
        self.Rama.rebuildModel()
        self.Protyvaga.rebuildModel()
        # зборки:
        self.BalansirVZbori.rebuildModel()
        self.Krivoshypy.rebuildModel()
        self.TraversaVZbori.rebuildModel()
        self.Karkas.rebuildModel()

class Profil(object):
    """Клас профілю"""
    fileName=None
    d={} # словник параметрів
    
    def create(self, name):
        self.getParamFromCSV(name)
        
    def setSWParam(self,d,nameList):
        """Повертає словник параметрів для SolidWorks.
        d - початковий словник параметрів SolidWorks,
        nameList - список назв елементів SolidWorks.
        Приклад:
        setSWParam({"Висота@Профіль1":0, "Висота@Профіль2":0},
            ["Профіль1","Профіль2"])
        """
        for k in d: # для усіх ключів словника параметрів для SolidWorks
            if k.count('@')==1: # якщо в ключі є тільки 1 символ '@'
                pair=k.split('@') # розділити на дві частини
                if pair[1] in nameList: # якщо назва елементу в списку
                    d[k]=self.d[pair[0]] # змінити значення
        return d
    
    def getParamFromCSV(self, name):
        """З файлу CSV записує в словник параметрів параметри профілю з назвою name"""
        import csv
        csv_file=open(self.fileName+".csv", "rb")
        reader=csv.DictReader(csv_file, delimiter = ';')
        for row in reader:
            if row["Назва"]==name:
                self.d=row
        csv_file.close()
        
        # перетворює значення словника в дійсні числа
        for k in self.d:
            if k!="Назва": # окрім назви
                if "," in self.d[k]: # якщо є символ ","
                    self.d[k]=float(self.d[k].replace(",", "."))
                else:
                    self.d[k]=float(self.d[k])

class DvotavrProfil(Profil):
    """Клас двотавра"""
    fileName="Двотаври ГОСТ 8239-89"
    
class ShvelerProfil(Profil):
    """Клас швелера"""
    fileName="Швелери серія У ДСТУ 3436-96"

class KutnykProfil(Profil):
    """Клас кутника"""
    fileName="Кутник ДСТУ 2251-93"
                               
class Balansir(SWmodelPRT):
    """Клас балансиру"""
    fileName="Балансир"
    #словник параметрів
    d={
    "Двотавр номер профілю" : None}
    
    def create(self, paramPU, paramGolovBalansir):
        self.read_dict_from_SW_equations() # читати параметри з файлу рівнянь
        
        # профіль двотавра
        self.assignParam(DvotavrProfil, self.d["Двотавр номер профілю"], ["Двотавр профіль"])
            
        # довжина плеч
        lpp=paramPU["Довжина переднього плеча балансира"]
        lzp=paramPU["Довжина заднього плеча балансира"]
        # довжина головки балансира
        lgb=paramGolovBalansir["Ширина@Профіль головки"]-paramGolovBalansir["Глибина опори@Профіль головки"]
        self.d["Довжина переднього плеча балансира"]=lpp-lgb  
        self.d["Довжина@Двотавр"]=lpp-lgb+self.d["Координата@Профіль опори балансира"]+self.d["Ширина@Профіль опори балансира"]/2
        self.d["Координата@Профіль опори балансира"]=lzp-self.d["Ширина@Профіль опори балансира"]/2+self.d["Ширина@Профіль опори"]/2+self.d["Координата@Профіль опори"]
        # координати ребер жорсткості
        self.d["Координата@Ребро профіль"]=100.0
        self.d["Координата@Масив ребер1"]=(lzp-100)/2
        self.d["Координата@Масив ребер2"]=(lzp+100)
        self.d["Координата@Масив ребер3"]=lzp+lpp/2
        
class GolovBalansir(SWmodelPRT):
    """Клас головки балансира"""
    fileName="Головка балансира"
    #словник параметрів
    d={
    "Двотавр номер профілю" : None,        
    "Радіус@Профіль головки" : 2290.0,
    "Висота між опорами@Профіль головки" : 640.0,
    "Ширина@Профіль головки" : 1000.0,
    "Глибина опори@Профіль головки" : 300.0,
    "Зовнішній радіус@Двотавр профіль" : 6.0,
    "Внутрішній радіус@Двотавр профіль" : 14.0,
    "Товщина середини основи@Двотавр профіль" : 12.3,
    "Товщина стінки@Двотавр профіль" : 7.5,
    "Ширина@Двотавр профіль" : 145.0,
    "Висота@Двотавр профіль" : 360.0,
    "Висота@Перемістити грань1" : 70.0,
    "Висота@Перемістити грань2" : 70.0,
    "Висота осі опори@Профіль головки" : 120.0,
    "Кут@Профіль головки" : 75.0}
    
    def create(self, paramPU):
        # профіль двотавра
        self.assignParam(DvotavrProfil, self.d["Двотавр номер профілю"], ["Двотавр профіль"])
        
        r=paramPU["Довжина переднього плеча балансира"]
        self.d["Радіус@Профіль головки"]=r
        s=max(paramPU["Список довжин ходу полірованого штока"]) # найбільша
        self.d["Кут@Профіль головки"]=degrees(s/r)
        
class Shatun(SWmodelPRT):
    """Клас шатуна"""
    fileName="Шатун"
    #словник параметрів
    d={"Довжина@Тіло" : 3000.0,
       "Ширина@Опора" : 70.0}
    def create(self, paramPU=None):
        dsh=paramPU["Довжина шатуна"]
        self.d["Довжина@Тіло"]=dsh

class Krivoshyp(SWmodelPRT):
    """Клас кривошипа"""
    fileName="Кривошип"
    #словник параметрів
    d={"Довжина@Ескіз" : 2000.0,
       "Радіус@Ескіз" : 1290.0,
       "Ширина@Профіль" : 150.0,
       "Висота@Ескіз" : 400.0,
       "Координата отвору@Ескіз" : 300.0}
    
    def create(self, paramPU=None):
        r=paramPU["Радіус кривошипа"]
        rmax=paramPU["Найбільший радіус кривошипа"]
        self.d["Радіус@Ескіз"]=r
        self.d["Довжина@Ескіз"]=rmax+600.0          

class Traversa(SWmodelPRT):
    """Клас траверси"""
    fileName="Траверса"
    #словник параметрів
    d={"Зовнішній радіус@Швелер профіль" : 6.0,
       "Внутрішній радіус@Швелер профіль" : 15.0,
       "Товщина середини основи@Швелер профіль" : 13.5,
       "Товщина стінки@Швелер профіль" : 8.0,
       "Кут основи@Швелер профіль" : 5.73,
       "Ширина@Швелер профіль" : 115.0,
       "Висота@Швелер профіль" : 400.0,
       "Півдовжина@Швелер" : 1125.0,
       "Товщина@Профіль ребра" : 10.0,
       "Координата@Профіль ребра" : 100.0,
       "Координата@Масив ребер" : 562.5,
       "Кількість@Масив ребер" : 2,
       "Діаметр@Ескіз отвору під шатун" : 50.0,
       "Координата@Ескіз отвору під шатун" : 100.0,
       "Висота@Ескіз опори" : 240.0,
       "Діаметр отвору@Ескіз опори" : 100.0,
       "Висота осі@Ескіз опори" : 120.0,
       "Товщина@Опора" : 30.0,
       "Координата@Опора" : 170.0,
       "Товщина@Профіль основи опори" : 10.0,
       "Півдовжина@Основа опори": 240.0}
    
    def create(self, paramPU=None):
        # профіль швелера
        self.assignParam(ShvelerProfil, self.d["Швелер номер профілю"], ["Швелер профіль"])
            
        ln=paramPU["Ширина"]
        self.d["Півдовжина@Швелер"]=ln/2
        self.d["Координата@Масив ребер"]=self.d["Півдовжина@Швелер"]/2

class Val(SWmodelPRT):
    """Клас валу"""
    fileName="Вал"
    #словник параметрів
    d={"Довжина@Вал" : 2400.0}
    def create(self, paramPU, paramShatun, paramTraversa):
        l1=paramShatun["Ширина@Опора"]
        ln=paramTraversa["Півдовжина@Швелер"]
        l2=paramTraversa["Координата@Ескіз отвору під шатун"]
        self.d["Довжина@Вал"]=2*(ln-l2-l1/2)

class Reduktor(SWmodelPRT):
    """Клас редуктора"""
    fileName="Редуктор"
    #словник параметрів
    d={"Висота осі@Профіль отвору" : 350.0,
       "Координата осі@Профіль отвору" : 300.0,
       "Довжина@Профіль корпуса" : 1100.0}
           
class Stiyka(SWmodelPRT):
    """Клас стійки"""
    fileName="Стійка"
    #словник параметрів
    d={"Ширина@Ескіз основи" : 2000.0,
       "Довжина@Ескіз основи" : 2000.0,
       "Висота@Основа" : 4500.0,
       "Висота@Секція1" : 500.0,
       "Висота@Секція2" : 1500.0,
       "Висота@Секція3" : 2500.0,
       "Висота@Секція4" : 3500.0,
       "Ширина@Ескіз опори балансира" : 400.0,
       "Висота@Ескіз опори балансира" : 400.0,
       "Висота осі@Ескіз опори балансира" : 200.0,
       "Товщина@Опора балансира" : 100.0,
       "Координата@Опора балансира" : 170.0,
       "Товщина@Основа опори" : 20.0,
       "Кут" : 8}
    
    def create(self, paramPU, paramReduktor, paramVal, paramKrivoshyp):
        s=paramVal["Довжина@Вал"]-2*paramKrivoshyp["Ширина@Профіль"]-100
        self.d["Ширина@Ескіз основи"]=s
        self.d["Довжина@Ескіз основи"]=s
        h=paramPU["Вертикальна відстань між осями опори балансира і тихохідного валу редуктора"]
        hr=paramReduktor["Висота осі@Профіль отвору"]
        self.d["Висота@Основа"]=h+hr-self.d["Висота осі@Ескіз опори балансира"]-self.d["Товщина@Основа опори"]
        hsect1=self.d["Висота@Основа"]/9
        self.d["Висота@Секція1"]=hsect1
        hsect=(self.d["Висота@Основа"]-hsect1)/4
        self.d["Висота@Секція2"]=hsect1+hsect
        self.d["Висота@Секція3"]=hsect1+2*hsect
        self.d["Висота@Секція4"]=hsect1+3*hsect
        
class Rama(SWmodelPRT):
    """Клас рами"""
    fileName="Рама"
    #словник параметрів
    d={"Довжина@Ескіз" : 5000.0,
       "Ширина@Ескіз" : 2000.0,
       "L1@Ескіз" : 200.0,
       "L2@Ескіз" : 2000.0,
       "L3@Ескіз" : 400.0,
       "L4@Ескіз" : 1000.0,
       "L5@Ескіз" : 300.0}
    
    def create(self, paramPU, paramStiyka, paramReduktor):
        l0=paramPU["Горизонтальна відстань між осями опори балансира і тихохідного валу редуктора"]
        s=paramStiyka["Ширина@Ескіз основи"]
        l2=paramStiyka["Довжина@Ескіз основи"]
        lr=paramReduktor["Довжина@Профіль корпуса"]
        ko=paramReduktor["Координата осі@Профіль отвору"]
        
        self.d["Ширина@Ескіз"]=s
        self.d["L2@Ескіз"]=l2
        self.d["L3@Ескіз"]=l0-l2/2-ko
        self.d["L4@Ескіз"]=lr-100
        self.d["Довжина@Ескіз"]=self.d["L1@Ескіз"]+self.d["L2@Ескіз"]+self.d["L3@Ескіз"]+self.d["L4@Ескіз"]+self.d["L5@Ескіз"]+700     

class Protyvaga(SWmodelPRT):
    """Клас противаги"""
    fileName="Противага"
    #словник параметрів
    d={"Радіус@Эскиз1" : 1625.0,
       "Півширина кривошипа@Эскиз1" : 200.0,
       "Ширина@Эскиз1" : 800.0,
       "Товщина@Бобышка-Вытянуть1" : 150.0}
    
    def create(self, paramKrivoshyp):
        self.d["Радіус@Эскиз1"]=paramKrivoshyp["Довжина@Ескіз"]-paramKrivoshyp["Координата отвору@Ескіз"]
        self.d["Півширина кривошипа@Эскиз1"]=paramKrivoshyp["Висота@Ескіз"]/2
        self.d["Товщина@Бобышка-Вытянуть1"]=paramKrivoshyp["Ширина@Профіль"]
        # ширина залежить від маси
        
class BalansirVZbori(SWmodelASM):
    """Клас балансира в зборі"""
    fileName="Балансир в зборі"

class Krivoshypy(SWmodelASM):
    """Клас кривошипів з валом в зборі"""
    fileName="Кривошипи"    

class TraversaVZbori(SWmodelASM):
    """Клас траверси з шатунами в зборі"""
    fileName="Траверса з шатунами"

class Karkas(SWmodelASM):
    """Клас каркасу в зборі"""
    fileName="Каркас"     

# pu=PumpingUnit()            
# pu.create()
# pu.rebuildModel()

# k=KutnykProfil()
# k.create("150х150х12")
# print k.d["Зовнішній радіус"]
