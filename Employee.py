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
from Personne import *
from Fonctions_utiles import *
from En_tete_Excel import *

class Employe(Personne):
    def __init__(self,nom, prenom , mail, tel, poste, station):
        Personne.__init__(self, nom, prenom, mail, tel)
        
        if (poste in ['chauffeur', 'ramasseur'])== False:
            return 'Modifiez le poste'
        
        self.poste= poste
        self.station = station
        self.dispo= True
        
        
        
        
    def verif_ins(self):
        X = recup_excel("Gestion.xlsx", "Employés")
        
        if (self.mail in list(X['MAIL'])) == True:
            
            f = True
        else:
            f = False
        
        return f
    
    
    def enreg_empl(self):
        
        if Employe.verif_ins(self)== True:
            return "Employé déjà enregistré"
        
        X = recup_excel("Gestion.xlsx", "Employés")
        
        a = [self.nom, self.prenom, self.mail, self.tel, self.poste, self.station, 'D']
        
        
         
        a = np.array(a)
        a = a.reshape(1, 7)
        a = pd.DataFrame(a)
        a.to_excel(B, "Employés", header= False, index=False, startrow= X.shape[0] + 1, startcol= 0 )
        B.save()
        print("L'employé dont l'email est le suivant : ",self.mail," a été enregistré")
        
    def suppr_empl(self):
        if Employe.verif_ins(self) == True:
            X = recup_excel("Gestion.xlsx", "Employés")
            
            v = list(X['MAIL']).index(self.mail)
            
            a = [self.nom, self.prenom, self.mail, self.tel, self.poste,self.station, 'Viré']
            a = np.array(a)
            a = a.reshape(1, 7)
            a = pd.DataFrame(a)
            a.to_excel(B, "Employés", header= False, index=False, startrow= v + 1 , startcol= 0 )
            B.save()
            print("Suppression de l'employé d'email ",self.mail," terminéé")
        else:
            print(" Employé non enregistré, suppression impossible")
            
    
    def infos(self):
        a = [self.nom, self.prenom, self.mail, self.tel, self.poste, self.station]
        return a
    
        
    def modif_empl(self, nom, prenom, mail, tel, poste, station):
        
        
        if Employe.verif_ins(self) == True:
            X = recup_excel("Gestion.xlsx", "Employés")            
            
            v = list(X['MAIL']).index(self.mail)
            
            self.nom = nom
            self.prenom = prenom
            self.mail = mail
            self.tel = tel    
            self.poste =  poste 
            self.station = station
            
            a = [self.nom, self.prenom, self.mail, self.tel, self.poste, self.station, X['DISPONIBILITE'][v]]
            
            
            a = np.array(a)
            a = a.reshape(1, 7)
            a = pd.DataFrame(a)
            a.to_excel(B, "Employés", header= False, index=False, startrow= v + 1, startcol= 0 )
            B.save()
            print("Modification de l'employé terminée")
            
            
        else:
            self.nom = nom
            self.prenom = prenom
            self.mail = mail
            self.tel = tel    
            self.poste =  poste 
            self.station = station
            
