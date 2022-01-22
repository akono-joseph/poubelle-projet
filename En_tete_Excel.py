# -*- coding: utf-8 -*-
"""

@author: Orie-san
"""

import pandas as pd
import openpyxl
import  xlsxwriter
import numpy as np
import datetime as dt
import openpyxl as op
import random


book = op.load_workbook('Gestion.xlsx')

B = pd.ExcelWriter('Gestion.xlsx', engine='openpyxl') 
B.book = book

B.sheets = dict((ws.title, ws) for ws in book.worksheets)


###Préparation des en-têtes:

#Employés    
z1 = ['NOM','PRENOM', 'MAIL', 'TEL', 'POSTE', 'STATION',  'DISPONIBILITE']
z1 = np.array(z1)

z1 = z1.reshape(1, 7)

z1 = pd.DataFrame(z1)

z1.to_excel(B, "Employés", header=False, index=False, startrow= 0  , startcol= 0 ) 
B.save()




#Poubelles
z2 = ['ID','LOCALISATION', 'TYPE', 'CONT ACTU']
z2 = np.array(z2)

z2 = z2.reshape(1, 4)

z2 = pd.DataFrame(z2)


z2.to_excel(B, "Poubelles", header= False, index=False, startrow= 0, startcol= 0 ) #(B, 'nom de la feuille')
B.save()


#Camion
z3 = ['IMMATRI', 'TYPE', 'DISPONIBILITE']
z3 = np.array(z3)

z3 = z3.reshape(1, 3)

z3 = pd.DataFrame(z3)

z3.to_excel(B, "Camions", header= False, index=False, startrow= 0, startcol= 0 ) #(B, 'nom de la feuille')
B.save()



#Ramassage
z4 = ['DATE ET HEURE', 'POUBELLE ID', 'LOCALISATION', 'TYPE POUBELLE', 'CAMION', 'CHAUFFEUR','R1','R2','R3']
z4 = np.array(z4)

z4 = z4.reshape(1, 9)

z4 = pd.DataFrame(z4)

z4.to_excel(B, "Ramassage", header= False, index=False, startrow= 0, startcol= 0 ) #(B, 'nom de la feuille')
B.save()
