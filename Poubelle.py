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
from Fonctions_utiles import *
from En_tete_Excel import *



class Poubelle():
    def __init__(self, ID, loc, genre, cont=0, p = False, pp= False ): #p = plein; pp= presque plein; cont = contenance actuelle de la poubelle
        if (genre in ['P', 'M', 'G'] ) == False:
            return 'Modifiez le genre'
        
        
        
        self.ID = ID
        self.loc = loc
        self.genre = genre
        self.cont = cont
        self.p = p
        self.pp = pp
        
        if genre == 'P':
            self.total = 300
        elif genre == 'M':
            self.total = 1000
        else:
            self.total = 2000
     
    def verif_pou(self):
        X = recup_excel("Gestion.xlsx", "Poubelles")
        
        if (self.ID in list(X['ID'])) == True:
            
            f = True
        else:
            f = False
        
        return f
    
    def enreg_pou(self):
        
        if Poubelle.verif_pou(self)== True:
            return 'Poubelle déjà enregistrée'
        
        X = recup_excel("Gestion.xlsx", "Poubelles")
        
        a = [self.ID, self.loc, self.genre, self.cont]
        
        
            
        a = np.array(a)
        a = a.reshape(1, 4)
        a = pd.DataFrame(a)
        a.to_excel(B, "Poubelles", header= False, index=False, startrow= X.shape[0] + 1, startcol= 0 )
        B.save()
        print("Enregistrement de la poubelle d'ID ",self.ID," effectué")
        
        
    def suppr_pou(self):
        if Poubelle.verif_pou(self) == True:
            X = recup_excel("Gestion.xlsx", "Poubelles")
            
            v = list(X['ID']).index(self.ID)
            
            a = [self.ID, self.loc, 'Fermée', ""]
            a = np.array(a)
            a = a.reshape(1, 4)
            a = pd.DataFrame(a)
            a.to_excel(B, "Poubelles", header= False, index=False, startrow= v+1 , startcol= 0 )
            B.save()
            print("Suppression de la poubelle ",self.ID," effectuée")
        
        else:
            print(" Poubelle non enregistrée, suppression impossible")
        
        
    def modif_poubelle(self, ID, loc, genre, cont):
         if Poubelle.verif_pou(self) == True:
            X = recup_excel("Gestion.xlsx", "Poubelles")            
            
            v = list(X['ID']).index(self.ID)
            
            self.ID = ID
            self.loc = loc
            self.genre = genre
            self.cont = cont    
             
            
            a = [self.ID, self.loc, self.genre, self.cont]
                             
            a = np.array(a)
            a = a.reshape(1, 4)
            a = pd.DataFrame(a)
            a.to_excel(B, "Poubelles", header= False, index=False, startrow= v + 1, startcol= 0 )
            B.save()    
            
            
         else:
            self.ID = ID
            self.loc = loc
            self.genre = genre
            self.cont = cont
            
        
    
    def recevoir(self, quantite):
         
        if Poubelle.verif_pou(self) == True:
            self.cont = self.cont + quantite
            
            X = recup_excel("Gestion.xlsx", "Poubelles")            
            
            v = list(X['ID']).index(self.ID)
            a = [self.ID, self.loc, self.genre, self.cont]
                             
            a = np.array(a)
            a = a.reshape(1, 4)
            a = pd.DataFrame(a)
            a.to_excel(B, "Poubelles", header= False, index=False, startrow= v + 1, startcol= 0 )
            B.save()
            
            if self.cont > 0.9 * self.total:
                self.pp = True 
                
            if self.cont >= self.total:
                self.p = True
                
        else:
            return 'Poubelle non enregistrée; elle ne peut être remplie'
        
        
            

    def vider(self):
        if Poubelle.verif_pou(self) == True:
            self.cont = 0
            
            Poubelle.modif_poubelle(self, self.ID, self.loc, self.genre, self.cont)
            print("Poubelle ",self.ID," du secteur ",self.loc," vidée avec succès")
        
        else:
            return 'Poubelle non enregistrée; Impossible à vider'
        
