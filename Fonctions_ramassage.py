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




def signal(P):
    
    genre = P.genre
    F = recup_excel("Gestion.xlsx", "Camions")
    
    if P.pp == True and P.p == False:
        print("Poubelle bientôt pleine")

    if P.p == True:
        d = dt.datetime.today()
        A = [str(d.date())+ " "+ str(d.hour)+ " " +str(d.minute) ,  P.ID, P.loc, P.genre]
        
        
        F1 = (F[['IMMATRI', 'DISPONIBILITE']])[ F['TYPE']== P.genre]
            
        F1 = (F1[['IMMATRI']])[ F1['DISPONIBILITE']== 'D' ]
            
        F1 = list(F1['IMMATRI'])

        if len(F1) != 0:
            veh = random.choice(F1)
            A.append(veh)
            
            v = list(F['IMMATRI']).index(veh)
            
            a = [veh, genre, 'ND']
            a = np.array(a)
            a = a.reshape(1, 3)
            a = pd.DataFrame(a)
            a.to_excel(B, "Camions", header= False, index=False, startrow= v + 1, startcol= 0 )
            B.save()
        
            E = recup_excel("Gestion.xlsx", "Employés")
            E1 = (E[['NOM', 'DISPONIBILITE', 'POSTE']])[ E['STATION']== P.loc]
            E1 = (E1[['NOM', 'DISPONIBILITE']])[ E1['POSTE']== 'chauffeur']
            E1 = (E1[['NOM']])[ E1['DISPONIBILITE']== 'D']
            E1 = list(E1['NOM'])

            if len(E1) == 0:
                return "Chauffeur indisponible pour cette opération de ramassage"
            else:
                k = random.choice(E1)
                A.append(k)
        
                u = list(E['NOM']).index(k)
            
                a = [k, E['PRENOM'][u] , E['MAIL'][u], E['TEL'][u],  E['POSTE'][u],   E['STATION'][u]   , 'ND']
                a = np.array(a)
                a = a.reshape(1, 7)
                a = pd.DataFrame(a)
                a.to_excel(B, "Employés", header= False, index=False, startrow= u + 1, startcol= 0 )
                B.save()
        
        
                L = ['P', 'M', 'G']
        
                r = L.index(P.genre) + 1
        
        
        
                for i in range(r):
                    E = recup_excel("Gestion.xlsx", "Employés")
                    E1 = (E[['NOM', 'DISPONIBILITE', 'POSTE']])[ E['STATION']== P.loc]
                    E1 = (E1[['NOM', 'DISPONIBILITE']])[ E1['POSTE']== 'ramasseur']
                    E1 = (E1[['NOM']])[ E1['DISPONIBILITE']== 'D']
                    E1 = list(E1['NOM'])
                    if len(E1) == 0:
                        return "Ramasseur indisponible pour cette mission"
                    else:    
                        k = random.choice(E1)
                        A.append(k)
            
                        v = list(E['NOM']).index(k)
            
                        a = [k, E['PRENOM'][v] , E['MAIL'][v], E['TEL'][v],  E['POSTE'][v],   E['STATION'][v]   , 'ND']
                        a = np.array(a)
                        a = a.reshape(1, 7)
                        a = pd.DataFrame(a)
                        a.to_excel(B, "Employés", header= False, index=False, startrow= v + 1, startcol= 0 )
                        B.save()
            
                while len(A)<9:
                    A.append("")
        
                A= np.array(A)
                A = A.reshape(1,9)
                A = pd.DataFrame(A)
                X = recup_excel("Gestion.xlsx", "Ramassage") 
        
                A.to_excel(B, "Ramassage", header= False, index=False, startrow= X.shape[0] + 1, startcol= 0 )
                B.save()
        
                P.vider()
        else:
            return "Aucun camion disponible pour cette mission de ramassage"

        
        


def rendre_dispo():
    X = recup_excel("Gestion.xlsx", "Ramassage") 
    s = dt.datetime.today()
    
    A = X[['DATE ET HEURE', "CAMION", "CHAUFFEUR", "R1", "R2", "R3"]]
    
    S = A['DATE ET HEURE']
    S = list(S)
    for i in S:
        j = i.split()
        j = j[2]
        
        if abs(s.minute - float(j) ) >=1:
            k = S.index(i)
            C0 = A['CAMION'][k]
            C1 = A['CHAUFFEUR'][k]
            C2 = A['R1'][k]
            C3 = A['R2'][k]
            C4 = A['R3'][k]
            
            E =  recup_excel("Gestion.xlsx", "Employés")
            H = list(E['NOM'])
            
            
            for l in [C1, C2, C3, C4]:
                try:
                    v = H.index(l)
                    a= [l, E['PRENOM'][v] , E['MAIL'][v], E['TEL'][v],  E['POSTE'][v], E['STATION'][v]    , 'D']
                    a = np.array(a)
                    a = a.reshape(1, 7)
                    a = pd.DataFrame(a)
                    a.to_excel(B, "Employés", header= False, index=False, startrow= v + 1, startcol= 0 )
                    B.save()
                except:
                    pass

        F =  recup_excel("Gestion.xlsx", "Camions")
        G = list(F['IMMATRI'])
            
        try:
            v = G.index(C0)
            a= [C0, E['TYPE'][v], 'D']
            a = np.array(a)
            a = a.reshape(1, 3)
            a = pd.DataFrame(a)
            a.to_excel(B, "Camions", header= False, index=False, startrow= v + 1, startcol= 0 )
            B.save()
        except:
            pass
                        
            
    
    