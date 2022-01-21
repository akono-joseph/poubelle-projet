# -*- coding: utf-8 -*-
"""
@author: Orie-san
"""

import pandas as pd
import openpyxl
import xlsxwriter
import numpy as np
import datetime as dt
import openpyxl as op
import random

#Création du fichier
#A = pd.ExcelWriter("Gestion.xlsx", engine='xlsxwriter' )

#z = ["", ""]
#z = np.array(z)
#z = pd.DataFrame(z)
#z.to_excel(A, "Employés", header=False, index=False, startrow= 0  , startcol= 0 ) 
#A.save()
##



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

#


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


##Fonctions utiles
def recup_excel(document, feuille):
    import pandas as pd
    X = pd.read_excel(document, feuille)
    return X



    

#Classes
class Personne():
    def __init__(self, nom, prenom, mail, tel):
        self.nom = nom
        self.prenom = prenom
        self.mail = mail
        self.tel = tel
        


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
            return 'Employé déjà enregistré'
        
        X = recup_excel("Gestion.xlsx", "Employés")
        
        a = [self.nom, self.prenom, self.mail, self.tel, self.poste, self.station, 'D']
        
        
         
        a = np.array(a)
        a = a.reshape(1, 7)
        a = pd.DataFrame(a)
        a.to_excel(B, "Employés", header= False, index=False, startrow= X.shape[0] + 1, startcol= 0 )
        B.save()
        
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
            
            
        else:
            self.nom = nom
            self.prenom = prenom
            self.mail = mail
            self.tel = tel    
            self.poste =  poste 
            self.station = station
            
            
    '''
    def travail(self):
        if self.dispo == False:
            return 'employé non disponible'
        else:
            self.dispo = False
        
        if Employe.verif_ins(self)== True:
            X = recup_excel("Gestion.xlsx", "Employés")
            v = list(X['MAIL']).index(self.mail)
            
            a = [self.nom, self.prenom, self.mail, self.tel, self.poste, 'ND']
            a = np.array(a)
            a = a.reshape(1, 6)
            a = pd.DataFrame(a)
            a.to_excel(B, "Employés", header= False, index=False, startrow= v + 1, startcol= 0 )
            B.save()
            
        
        else:
            print("Employé non enregistré; il ne peut travailler")'''
        
        

class Poubelle():
    def __init__(self, ID, loc, genre, cont=0, p = False, pp= False ): #p = plein; pp= presque plein
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
        print("enregistrement effectué")
        
        
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
            print("suppression effectuée")
        
        else:
            print(" Poubelle non enregistrée, suppression impossible")
        
        
    def modif_pou(self, ID, loc, genre, cont):
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
            
            Poubelle.modif_pou(self, self.ID, self.loc, self.genre, self.cont)
            print("Poubelle vidée avec succès")
        
        else:
            return 'Poubelle non enregistrée; impossible à vider'
        
        
        


            
            
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
            return 'Poubelle déjà enregistrée'
        
        X = recup_excel("Gestion.xlsx", "Camions")
        
        a = [self.imm, self.genre, 'D']
        
        a = np.array(a)
        a = a.reshape(1, 3)
        a = pd.DataFrame(a)
        a.to_excel(B, "Camions", header= False, index=False, startrow= X.shape[0] + 1, startcol= 0 )
        B.save()
        print("enregistrement effectué")
        
    
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
            print("suppression effectuée")
        
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
            
            
         else:
            self.imm = imm
            self.genre = genre
  
    
'''
    def travail_cam(self):
        if self.dispo == False:
            return 'Camion non disponible'
        else:
            self.dispo = False
        
        if Camion.verif_cam(self)== True:
            X = recup_excel("Gestion.xlsx", "Camions")
            v = list(X['IMMATRI']).index(self.imm)
            
            a = [self.imm, self.genre, 'ND']
            a = np.array(a)
            a = a.reshape(1, 3)
            a = pd.DataFrame(a)
            a.to_excel(B, "Camions", header= False, index=False, startrow= v + 1, startcol= 0 )
            B.save()
            
        
        else:
            print("Camion non enregistré; il ne peut être pris")'''
        
  
    
  
    
  
    


##FONCTIONS RAMASSAGE

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
        k = random.choice(E1)
        A.append(k)
        
        u = list(E['NOM']).index(k)
            
        a = [k, E['PRENOM'][u] , E['MAIL'][u], E['TEL'][u],  E['POSTE'][u],   E['STATION'][u]   , 'ND']
        a = np.array(a)
        a = a.reshape(1, 7)
        a = pd.DataFrame(a)
        a.to_excel(B, "Employés", header= False, index=False, startrow= v + 1, startcol= 0 )
        B.save()
        
        
        L = ['P', 'M', 'G']
        
        r = L.index(genre) + 1
        
        
        
        for i in range(r):
            E = recup_excel("Gestion.xlsx", "Employés")
            E1 = (E[['NOM', 'DISPONIBILITE', 'POSTE']])[ E['STATION']== P.loc]
            E1 = (E1[['NOM', 'DISPONIBILITE']])[ E1['POSTE']== 'ramasseur']
            E1 = (E1[['NOM']])[ E1['DISPONIBILITE']== 'D']
            E1 = list(E1['NOM'])
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
        


def rendre_dispo():
    X = recup_excel("Gestion.xlsx", "Ramassage") 
    s = dt.datetime.today()
    
    A = X[['DATE ET HEURE', "CHAUFFEUR", "R1", "R2", "R3"]]
    
    S = A['DATE ET HEURE']
    S = list(S)
    for i in S:
        j = i.split()
        j = j[2]
        
        if abs(s.minute - float(j) ) >=1:
            k = S.index(i)
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
            
    
    
        
        
            
            
        
        
            
            
            
    








        
        
        
###TESTS

P3 = Poubelle(1659, "Nlongkak", "P")          
P3.enreg_pou()        
P3.recevoir(900)


        
#P3.verif_pou()     


E1 = Employe("KLEB", "HABI", "HABI@gmail.com", 671874512, "ramasseur", "Nlongkak")
#E2 = Employe("FANKEP", "LOUIS", "louis@gmaiL.com", 677410528, "ramasseur","Flamenco")
#E3 = Employe("KOGORO", "MOURI", "mouri@gmail.com", 678954780, "ramasseur","Flamenco")
#E4 = Employe("NYASSA", "Louis", "nyassa@gmaail.com", 694879654, "ramasseur","Flamenco")
#E5 = Employe("Nouga", "Brice", "Brice@gmail.com", 698741787, "chauffeur","Flamenco")

#C1 = Camion("CK235", 'P')
#C2 = Camion("CM458",'P')
#C3 = Camion("AS342", 'M')
#C4 = Camion("GN2833", 'G')



#for i in [E1]:#E2,E3,E4,E5]:
#    i.enreg_empl()
    

#for i in [C1, C2, C3, C4]:
#    i.enreg_cam()



#signal(P3)        

rendre_dispo()        
    
    
#E1.verif_ins()        

#P1.suppr_pou()
