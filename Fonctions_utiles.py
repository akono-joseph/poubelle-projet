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

def recup_excel(document, feuille):
    import pandas as pd
    X = pd.read_excel(document, feuille)
    return X