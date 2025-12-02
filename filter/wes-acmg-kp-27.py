# -*- coding: utf-8 -*-
"""
Created on Wed Feb 26 12:24:46 2025

@author: esra.bardakci
"""

import numpy as np
import pandas as pd

filename='MG25_0079.xlsx' #put your file name here
final_name=filename.replace('.xlsx','')+'-WES-finall'
data=pd.read_excel(filename)
data.index+=1

acmg=['APC','MYH11','ACTA2','TMEM43','DSP','PKP2','DSG2','DSC2','BTD','BRCA1','BRCA2','SCN5A','RYR2','CASQ2','CALM1','TRDN',
     'FLNC','LMNA','TNNT2','DES','MYH7','TNNC1','RBM20','BAG3','TTN','COL3A1','GLA','LDLR','APOB','MYH7','TNNT2','TPM1','MYBPC3',
     'PRKAG2','TNNI3','MYL3','MYL2','ACTC1','RET','PALB2','HFE','ENG','ACVRL1','SDHD','SDHB','TTR','PCSK9','BMPR1A','SMAD4',
     'TP53','TGFBR1','SMAD3','TRDN','KCNQ1','KCNH2','CALM2','CALM3','MSH2','MLH1','PMS2','MSH6','RYR1','CACNA1S','FBN1',
     'HNF1A','MEN1','RET','MUTYH','DES','FLNC','BAG3','NF2','OTC','SDHAF2','SDHC','STK11','MAX','TMEM127','GAA','PTEN','RB1',
     'RPE65','TSC1','TSC2','VHL','WT1','ATP7B']

filter = data[data['OMIM Gene'].str.extract(r'\b(' + '|'.join(acmg) + r')\b', expand=False).notna()]
####27 gen#####
    
kp_update=['ABRAXAS1','FAM175A','ATM','BAP1', 'BRCA1', 'BRCA2','BRIP1','CDH1','CHEK2','DICER1','ERBB2','FANCC','MLH1', 'MRE11',
          'MSH2','MSH6', 'MUTYH', 'PALB2', 'PIK3CA', 'PMS2', 'POLD1', 'POLE', 'PTEN', 'RAD50', 'RAD51', 'SMARCA4', 'TP53', 'XRCC2']

filter2 = data[data['OMIM Gene'].str.extract(r'\b(' + '|'.join(kp_update) + r')\b', expand=False).notna()]

####export data####

with pd.ExcelWriter(final_name+'.xlsx') as writer:  
    data.to_excel(writer, sheet_name='original')
    filter.to_excel(writer,sheet_name='acmg')
    filter2.to_excel(writer,sheet_name='kp-27')
   
print('Exported file: {}'.format(final_name))
print('DONE:)')