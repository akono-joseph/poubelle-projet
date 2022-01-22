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
from Poubelle import *
from Camion import *
from Employee import *
from Fonctions_utiles import *
from Fonctions_ramassage import *
from En_tete_Excel import *


      
##TESTS

#Test enregistrement d'une poubelle
P1 = Poubelle(4785, "Nlongkak", "G")          
P1.enreg_pou()  

#Test mréception ordure d'une Poubelle
P1.recevoir(4000)
     

E1 = Employe("Fijls0", "Joet", "ddmddef@gmail.com", 3441052578, "ramasseur", "Nlongkak")
E2 = Employe("AROCOO", "Joer", "aaa@yahoo.com", 34414789, "ramasseur","Nlongkak")
E3 = Employe("NARAA", "Joen", "dss@yahoo.fr", 34447887, "ramasseur","Flamenco")
E4 = Employe("FOKANA", "Joep", "FOKANA@gmail.com", 614789971, "ramasseur","Mokolo")

C1 = Camion("CA234", 'P')
C2 = Camion("CA155",'P')
C3 = Camion("PS322", 'P')
C4 = Camion("L2313", 'G')



#test enregistrement Employés
for i in [E1,E2,E3,E4]:
    i.enreg_empl()
    

#Test enregistrement Camions
for i in [C1, C2, C3, C4]:
    i.enreg_cam()

#Test ramassage automatique d'une poubelle
signal(P1)        

#test remise en disponibilité des employés
rendre_dispo()
    
