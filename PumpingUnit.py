# -*- coding: CP1251 -*-
import os
import codecs
import subprocess
from math import *

class SWmodel(object):
    """���� ����� SolidWorks"""
    d={} # ������� ���������
    fileName=None # ��'� ����� ����� 
    fileType=".SLDPRT" # ���������� ����� ����� (".SLDPRT", ".SLDASM")
    
    def create(self):
        """��������� ��������� �����"""
        pass
    
    def rebuildModel(self):
        """���������� ���� ������ �� ������ SolidWorks"""
        self.write_dict_to_SW_equations()
        self.rebuildAndSaveModel()
        
    def rebuildAndSaveModel(self):
        """���������� �� ������ SolidWorks ������ ������ ��������� VBS �������"""
        
        vbs=r"""'������ VBS ��� ���������� ����� SolidWorks 
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
            docType="1" # ���� ������
        else:
            docType="2" # ���� ������
        vbs=vbs.format(fullFileNameExt=fullFileNameExt, fileNameExt=fileNameExt, fileName=fileName, docType=docType)
        scriptFileName=os.path.join(os.getcwd(), "RebuildSWmodelTemp.vbs")
        f=open(scriptFileName, 'w')
        f.write(vbs)
        f.close()
        # ������ ������ �� ���� ���� ����������
        subprocess.Popen(r'c:\Windows\system32\wscript.exe '+scriptFileName).wait()
    
    def read_dict_from_SW_equations(self):
        """���� �������� � ������� self.d
        � ���������� ����� (utf-8-sig) ������ SolidWorks"""
        filename=self.fileName+".txt" # ���� ������ SolidWorks
        if not os.path.exists(filename): return # ���� ����� �� ����, �����
        f=codecs.open(filename,'r','utf-8-sig') # ������ ���� ��� �������
        for line in f.readlines(): # ��� ��� ����� � ������
            if '=' in line: # ���� � ����� � ������ "="
                pair=line.split('=') # �������� �����
                pair=pair[0].strip()[1:-1], pair[1].strip() # �������� ������ � �����
                
                if pair[1].isdigit(): # ���� �� ������� �����
                    val=int(pair[1])
                else: # ������
                    try: # ���� ����� �����
                        val=float(pair[1])  
                    except ValueError: # ���� �����
                        val=pair[1]
                        
                self.d[pair[0].encode('CP1251')]=val # �������� � �������
        f.close() # ������� ����
        
    def write_dict_to_SW_equations(self):
        """������ �������� �������� �������� self.d
        � ��������� ���� (utf-8-sig) ������ SolidWorks"""
        d={} # ������� � unicode �������
        for k in self.d: # ������������ ����� � unicode
            d[k.decode('CP1251')]=self.d[k]
            
        filename=self.fileName+".txt" # ���� ������ SolidWorks
        if not os.path.exists(filename): return # ���� ����� �� ����, �����
        f=codecs.open(filename,'r','utf-8-sig') # ������ ���� ��� �������
        oldlines=f.readlines() # ������ ������ �����
        newlines=[] # ����� ������ �����
        for line in oldlines: # ��� ��� ����� � ������
            if '=' in line: # ���� � ����� � ������ "="
                pair=line.split('=') # �������� �����
                pair=pair[0].strip()[1:-1],pair[1].strip() # �������� ������ � �����
                if pair[0] in d.keys(): # ���� ��� ������� �� "=" � ����� ������ ��������
                    line='"'+pair[0]+'" = '+str(d[pair[0]])+"\n" # ��������� ����� �����
            newlines.append(line) # �������� ����� � ����� ������ �����
        f.close() # ������� ����
        
        f=codecs.open(filename,'w','utf-8-sig') # ������� ���� ��� ������
        f.writelines(newlines) # �������� ������ ����� �����
        f.close() # ������� ����
        
    def assignParam(self, profil, name, nameList):
        """�������� �������� ������� ��������� �������� ���������.
        profil - ����������� ���������� ������� (���. ����� �������),
        name - ����� �������,
        nameList - ������ ���� �������� SolidWorks.
        """
        p=profil()
        p.create(name=name)
        self.d=p.setSWParam(self.d, nameList)
            
class SWmodelPRT(SWmodel):
    """���� ����� ����� SolidWorks"""
    fileType=".SLDPRT"

class SWmodelASM(SWmodel):
    """���� ����� ������ SolidWorks"""
    fileType=".SLDASM"      

class PumpingUnit(SWmodelASM):
    """���� ��������-��������"""
    fileName="�������"
    d={
    "���" : "���8-3-4000",
    "����������� ��������� ������������" : 80.0,
    "������ ������ ���� ����������� �����" : [1200.0, 1600.0, 2000.0, 2500.0, 3000.0],
    "������� ���� ����������� �����" : 2000.0,
    "������ ������� �������" : [5.0,12.0],
    "ʳ������ �������" : 5.0,
    "������������ ������� ������" : 40.0,
    "������� ���������� ����� ���������" : 2290.0,
    "������� �������� ����� ���������" : 2000.0,
    "������� ������" : 3000.0,
    "��������� ����� ���������" : 1290.0,
    "����� ���������" : None,
    "������������� ������� �� ����� ����� ��������� � ����������� ���� ���������" : 1345.0,
    "����������� ������� �� ����� ����� ��������� � ����������� ���� ���������" : None,
    "�������" : 6900.0,
    "������" : 2250.0,
    "������" : 4910.0,
    "����" : 11780.0,
    "������� ��������������" : "����������",
    "���� ����������� ��������" : 750.0,
    "����������� ������� ����������� ��������" : 6,
    "��������� ��������� ��������������" : 18.5,
    "��������" : "�2��-750 �"}
    
    def create(self):
        """��������� ��������� �����"""
        sList=self.d["������ ������ ���� ����������� �����"]
        s0=self.d["������� ���� ����������� �����"]
        # ���������� ����������� ������ �� ����� ����� ��������� � ����������� ���� ���������
        smax=max(sList) # �������� ������� ����
        rmax=self.d["��������� ����� ���������"]
        l=self.d["������� ������"]
        k=self.d["������� �������� ����� ���������"]
        k1=self.d["������� ���������� ����� ���������"]
        l1=self.d["������������� ������� �� ����� ����� ��������� � ����������� ���� ���������"]
        h=smax*k/k1 # ������������ ��� ����� ��������
        b=sqrt(k**2-(h/2)**2) # ������� �� ����� ��������� �� �������� h
        H=sqrt((l-rmax)**2-(b-l1)**2)+h/2
        self.d["����������� ������� �� ����� ����� ��������� � ����������� ���� ���������"]=H
        
        # ���������� ������ ���������
        h=s0*k/k1 # ������������ ��� ����� ��������
        b=sqrt(k**2-(h/2)**2) # ������� �� ����� ��������� �� �������� h
        r=l-sqrt((H-h/2)**2+(b-l1)**2)
        self.d["����� ���������"]=r
        
        # ��������� �� ���������� �������
        self.GolovBalansir=GolovBalansir()
        self.GolovBalansir.create(paramPU=self.d)       
        self.Balansir=Balansir()
        self.Balansir.create(paramPU=self.d, paramGolovBalansir=self.GolovBalansir.d)
        self.GolovBalansir.d["������ �� �������@������� �������"]=self.Balansir.d["������@������� �������"]
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
        # ������:
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
        # ������:
        self.BalansirVZbori.rebuildModel()
        self.Krivoshypy.rebuildModel()
        self.TraversaVZbori.rebuildModel()
        self.Karkas.rebuildModel()

class Profil(object):
    """���� �������"""
    fileName=None
    d={} # ������� ���������
    
    def create(self, name):
        self.getParamFromCSV(name)
        
    def setSWParam(self,d,nameList):
        """������� ������� ��������� ��� SolidWorks.
        d - ���������� ������� ��������� SolidWorks,
        nameList - ������ ���� �������� SolidWorks.
        �������:
        setSWParam({"������@�������1":0, "������@�������2":0},
            ["�������1","�������2"])
        """
        for k in d: # ��� ��� ������ �������� ��������� ��� SolidWorks
            if k.count('@')==1: # ���� � ����� � ����� 1 ������ '@'
                pair=k.split('@') # �������� �� �� �������
                if pair[1] in nameList: # ���� ����� �������� � ������
                    d[k]=self.d[pair[0]] # ������ ��������
        return d
    
    def getParamFromCSV(self, name):
        """� ����� CSV ������ � ������� ��������� ��������� ������� � ������ name"""
        import csv
        csv_file=open(self.fileName+".csv", "rb")
        reader=csv.DictReader(csv_file, delimiter = ';')
        for row in reader:
            if row["�����"]==name:
                self.d=row
        csv_file.close()
        
        # ���������� �������� �������� � ���� �����
        for k in self.d:
            if k!="�����": # ���� �����
                if "," in self.d[k]: # ���� � ������ ","
                    self.d[k]=float(self.d[k].replace(",", "."))
                else:
                    self.d[k]=float(self.d[k])

class DvotavrProfil(Profil):
    """���� ��������"""
    fileName="�������� ���� 8239-89"
    
class ShvelerProfil(Profil):
    """���� �������"""
    fileName="������� ���� � ���� 3436-96"

class KutnykProfil(Profil):
    """���� �������"""
    fileName="������ ���� 2251-93"
                               
class Balansir(SWmodelPRT):
    """���� ���������"""
    fileName="��������"
    #������� ���������
    d={
    "������� ����� �������" : None}
    
    def create(self, paramPU, paramGolovBalansir):
        self.read_dict_from_SW_equations() # ������ ��������� � ����� ������
        
        # ������� ��������
        self.assignParam(DvotavrProfil, self.d["������� ����� �������"], ["������� �������"])
            
        # ������� ����
        lpp=paramPU["������� ���������� ����� ���������"]
        lzp=paramPU["������� �������� ����� ���������"]
        # ������� ������� ���������
        lgb=paramGolovBalansir["������@������� �������"]-paramGolovBalansir["������� �����@������� �������"]
        self.d["������� ���������� ����� ���������"]=lpp-lgb  
        self.d["�������@�������"]=lpp-lgb+self.d["����������@������� ����� ���������"]+self.d["������@������� ����� ���������"]/2
        self.d["����������@������� ����� ���������"]=lzp-self.d["������@������� ����� ���������"]/2+self.d["������@������� �����"]/2+self.d["����������@������� �����"]
        # ���������� ����� ���������
        self.d["����������@����� �������"]=100.0
        self.d["����������@����� �����1"]=(lzp-100)/2
        self.d["����������@����� �����2"]=(lzp+100)
        self.d["����������@����� �����3"]=lzp+lpp/2
        
class GolovBalansir(SWmodelPRT):
    """���� ������� ���������"""
    fileName="������� ���������"
    #������� ���������
    d={
    "������� ����� �������" : None,        
    "�����@������� �������" : 2290.0,
    "������ �� �������@������� �������" : 640.0,
    "������@������� �������" : 1000.0,
    "������� �����@������� �������" : 300.0,
    "������� �����@������� �������" : 6.0,
    "�������� �����@������� �������" : 14.0,
    "������� �������� ������@������� �������" : 12.3,
    "������� �����@������� �������" : 7.5,
    "������@������� �������" : 145.0,
    "������@������� �������" : 360.0,
    "������@���������� �����1" : 70.0,
    "������@���������� �����2" : 70.0,
    "������ �� �����@������� �������" : 120.0,
    "���@������� �������" : 75.0}
    
    def create(self, paramPU):
        # ������� ��������
        self.assignParam(DvotavrProfil, self.d["������� ����� �������"], ["������� �������"])
        
        r=paramPU["������� ���������� ����� ���������"]
        self.d["�����@������� �������"]=r
        s=max(paramPU["������ ������ ���� ����������� �����"]) # ��������
        self.d["���@������� �������"]=degrees(s/r)
        
class Shatun(SWmodelPRT):
    """���� ������"""
    fileName="�����"
    #������� ���������
    d={"�������@ҳ��" : 3000.0,
       "������@�����" : 70.0}
    def create(self, paramPU=None):
        dsh=paramPU["������� ������"]
        self.d["�������@ҳ��"]=dsh

class Krivoshyp(SWmodelPRT):
    """���� ���������"""
    fileName="��������"
    #������� ���������
    d={"�������@����" : 2000.0,
       "�����@����" : 1290.0,
       "������@�������" : 150.0,
       "������@����" : 400.0,
       "���������� ������@����" : 300.0}
    
    def create(self, paramPU=None):
        r=paramPU["����� ���������"]
        rmax=paramPU["��������� ����� ���������"]
        self.d["�����@����"]=r
        self.d["�������@����"]=rmax+600.0          

class Traversa(SWmodelPRT):
    """���� ��������"""
    fileName="��������"
    #������� ���������
    d={"������� �����@������ �������" : 6.0,
       "�������� �����@������ �������" : 15.0,
       "������� �������� ������@������ �������" : 13.5,
       "������� �����@������ �������" : 8.0,
       "��� ������@������ �������" : 5.73,
       "������@������ �������" : 115.0,
       "������@������ �������" : 400.0,
       "ϳ��������@������" : 1125.0,
       "�������@������� �����" : 10.0,
       "����������@������� �����" : 100.0,
       "����������@����� �����" : 562.5,
       "ʳ������@����� �����" : 2,
       "ĳ�����@���� ������ �� �����" : 50.0,
       "����������@���� ������ �� �����" : 100.0,
       "������@���� �����" : 240.0,
       "ĳ����� ������@���� �����" : 100.0,
       "������ ��@���� �����" : 120.0,
       "�������@�����" : 30.0,
       "����������@�����" : 170.0,
       "�������@������� ������ �����" : 10.0,
       "ϳ��������@������ �����": 240.0}
    
    def create(self, paramPU=None):
        # ������� �������
        self.assignParam(ShvelerProfil, self.d["������ ����� �������"], ["������ �������"])
            
        ln=paramPU["������"]
        self.d["ϳ��������@������"]=ln/2
        self.d["����������@����� �����"]=self.d["ϳ��������@������"]/2

class Val(SWmodelPRT):
    """���� ����"""
    fileName="���"
    #������� ���������
    d={"�������@���" : 2400.0}
    def create(self, paramPU, paramShatun, paramTraversa):
        l1=paramShatun["������@�����"]
        ln=paramTraversa["ϳ��������@������"]
        l2=paramTraversa["����������@���� ������ �� �����"]
        self.d["�������@���"]=2*(ln-l2-l1/2)

class Reduktor(SWmodelPRT):
    """���� ���������"""
    fileName="��������"
    #������� ���������
    d={"������ ��@������� ������" : 350.0,
       "���������� ��@������� ������" : 300.0,
       "�������@������� �������" : 1100.0}
           
class Stiyka(SWmodelPRT):
    """���� �����"""
    fileName="�����"
    #������� ���������
    d={"������@���� ������" : 2000.0,
       "�������@���� ������" : 2000.0,
       "������@������" : 4500.0,
       "������@������1" : 500.0,
       "������@������2" : 1500.0,
       "������@������3" : 2500.0,
       "������@������4" : 3500.0,
       "������@���� ����� ���������" : 400.0,
       "������@���� ����� ���������" : 400.0,
       "������ ��@���� ����� ���������" : 200.0,
       "�������@����� ���������" : 100.0,
       "����������@����� ���������" : 170.0,
       "�������@������ �����" : 20.0,
       "���" : 8}
    
    def create(self, paramPU, paramReduktor, paramVal, paramKrivoshyp):
        s=paramVal["�������@���"]-2*paramKrivoshyp["������@�������"]-100
        self.d["������@���� ������"]=s
        self.d["�������@���� ������"]=s
        h=paramPU["����������� ������� �� ����� ����� ��������� � ����������� ���� ���������"]
        hr=paramReduktor["������ ��@������� ������"]
        self.d["������@������"]=h+hr-self.d["������ ��@���� ����� ���������"]-self.d["�������@������ �����"]
        hsect1=self.d["������@������"]/9
        self.d["������@������1"]=hsect1
        hsect=(self.d["������@������"]-hsect1)/4
        self.d["������@������2"]=hsect1+hsect
        self.d["������@������3"]=hsect1+2*hsect
        self.d["������@������4"]=hsect1+3*hsect
        
class Rama(SWmodelPRT):
    """���� ����"""
    fileName="����"
    #������� ���������
    d={"�������@����" : 5000.0,
       "������@����" : 2000.0,
       "L1@����" : 200.0,
       "L2@����" : 2000.0,
       "L3@����" : 400.0,
       "L4@����" : 1000.0,
       "L5@����" : 300.0}
    
    def create(self, paramPU, paramStiyka, paramReduktor):
        l0=paramPU["������������� ������� �� ����� ����� ��������� � ����������� ���� ���������"]
        s=paramStiyka["������@���� ������"]
        l2=paramStiyka["�������@���� ������"]
        lr=paramReduktor["�������@������� �������"]
        ko=paramReduktor["���������� ��@������� ������"]
        
        self.d["������@����"]=s
        self.d["L2@����"]=l2
        self.d["L3@����"]=l0-l2/2-ko
        self.d["L4@����"]=lr-100
        self.d["�������@����"]=self.d["L1@����"]+self.d["L2@����"]+self.d["L3@����"]+self.d["L4@����"]+self.d["L5@����"]+700     

class Protyvaga(SWmodelPRT):
    """���� ���������"""
    fileName="���������"
    #������� ���������
    d={"�����@�����1" : 1625.0,
       "ϳ������� ���������@�����1" : 200.0,
       "������@�����1" : 800.0,
       "�������@�������-��������1" : 150.0}
    
    def create(self, paramKrivoshyp):
        self.d["�����@�����1"]=paramKrivoshyp["�������@����"]-paramKrivoshyp["���������� ������@����"]
        self.d["ϳ������� ���������@�����1"]=paramKrivoshyp["������@����"]/2
        self.d["�������@�������-��������1"]=paramKrivoshyp["������@�������"]
        # ������ �������� �� ����
        
class BalansirVZbori(SWmodelASM):
    """���� ��������� � ����"""
    fileName="�������� � ����"

class Krivoshypy(SWmodelASM):
    """���� ��������� � ����� � ����"""
    fileName="���������"    

class TraversaVZbori(SWmodelASM):
    """���� �������� � �������� � ����"""
    fileName="�������� � ��������"

class Karkas(SWmodelASM):
    """���� ������� � ����"""
    fileName="������"     

# pu=PumpingUnit()            
# pu.create()
# pu.rebuildModel()

# k=KutnykProfil()
# k.create("150�150�12")
# print k.d["������� �����"]
