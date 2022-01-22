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

class Camion():
    def __init__(self, imm, genre):
        if (genre in ['P', 'M', 'G'])== False:
            return "Modifiez le genre du camion"
            
        self.imm = imm
        self.genre = genre
        self.dispo = True
        
    
    def verif_cam(self):
        X = recup_excel("Gestion.xlsx", "Camions")
        
        if (self.imm in list(X['IMMATRI'])) == True:
            
            f = True
        else:
            f = False
        
        return f
    
    def enreg_cam(self):
        if Camion.verif_cam(self)== True:
            return "Camion déjà enregistrée"
        
        X = recup_excel("Gestion.xlsx", "Camions")
        
        a = [self.imm, self.genre, 'D']
        
        a = np.array(a)
        a = a.reshape(1, 3)
        a = pd.DataFrame(a)
        a.to_excel(B, "Camions", header= False, index=False, startrow= X.shape[0] + 1, startcol= 0 )
        B.save()
        print("Enregistrement du camion d'immatriculation ",self.imm," effectué avec succès")
        
    
    def suppr_cam(self):
         if Camion.verif_cam(self) == True:
            X = recup_excel("Gestion.xlsx", "Camions")
            
            v = list(X['IMMATRI']).index(self.imm)
            
            a = [self.imm,  'Usé', ""]
            a = np.array(a)
            a = a.reshape(1, 3)
            a = pd.DataFrame(a)
            a.to_excel(B, "Camions", header= False, index=False, startrow= v+1 , startcol= 0 )
            B.save()
            print("Suppression du camion d'immatriculation",self.imm," effectuée")
        
         else:
            print(" Camion non enregistré, suppression impossible")
        
  
    def modif_cam(self, imm, genre):
         if Camion.verif_cam(self) == True:
            X = recup_excel("Gestion.xlsx", "Camions")            
            
            v = list(X['IMMATRI']).index(self.imm)
            
            self.imm = imm
            self.genre = genre
             
            a = [self.imm,  self.genre, X['DISPONIBILITE'][v]]
            
            
                
            a = np.array(a)
            a = a.reshape(1, 3)
            a = pd.DataFrame(a)
            a.to_excel(B, "Camions", header= False, index=False, startrow= v + 1, startcol= 0 )
            B.save()    
            print("Modification du camion terminée")
            
            
         else:
            self.imm = imm
            self.genre = genre
