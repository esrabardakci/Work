# -*- coding: utf-8 -*-
"""
Created on Mon Jul 21 11:47:35 2025

@author: esra.bardakci
"""



import numpy as np
import pandas as pd

filename='MG25-0240.xlsx' #put your file name here
final_name=filename.replace('.xlsx','')+'-WES-finall'
data=pd.read_excel(filename)
data.index+=1

###ilaç hassasiyet####
ilachassasiyet=['CYP1A2','CYP2C9', 'CYP2C19', 'CYP3A4', 'CYP2D6']

#filter=data[:][data['OMIM Gene'].isin(list(acmg))]
#filter = data[data['OMIM Gene'].str.contains(r'\b(' + '|'.join(acmg) + r')\b', na=False)]
filter1 = data[data['OMIM Gene'].str.extract(r'\b(' + '|'.join(ilachassasiyet) + r')\b', expand=False).notna()]


###spor genetiği###
sporgenetigi=['ACE', 'ACTN3', 'VEGFR', 'IL6', 'NOS3', 'PPARG', 'PPARGC1A', 'DIO1' ]
filter2 = data[data['OMIM Gene'].str.extract(r'\b(' + '|'.join(sporgenetigi) + r')\b', expand=False).notna()]

####KP#####

with open('hereditercancer.txt','r') as file:
    kp=file.read()
    #print(kp)
    
kp_update=kp.replace("'","").strip('\n').split(', ')

#filter2=data[:][data['OMIM Gene'].isin(list(kp_update))]
#filter2 = data[data['OMIM Gene'].str.contains(r'\b(' + '|'.join(kp_update) + r')\b', na=False)]
filter3 = data[data['OMIM Gene'].str.extract(r'\b(' + '|'.join(kp_update) + r')\b', expand=False).notna()]

###nutrigenetik###
nutrigenetik=['SLC23A1', 'GSTT1', 'GSTM1', 'MTHFR','TCF7L2', 'NOS3', 'ACE', 'CYP1A2', 'GC', 'CYP2R1', 'DHCR7', 'APOA2', 'FTO', 'LCT', 'HLA-DQ']
filter4 = data[data['OMIM Gene'].str.extract(r'\b(' + '|'.join(nutrigenetik) + r')\b', expand=False).notna()]

###Kardiyomiyopati###
kardiyomiyopati=list(pd.read_excel('kardiyomiyopati.xlsx', header=None)[0].str.strip())
filter5 = data[data['OMIM Gene'].str.extract(r'\b(' + '|'.join(kardiyomiyopati) + r')\b', expand=False).notna()]

###Genel Filtre###
genelfiltre=ilachassasiyet+sporgenetigi+kp_update+nutrigenetik+kardiyomiyopati
filter6 = data[data['OMIM Gene'].str.extract(r'\b(' + '|'.join(genelfiltre) + r')\b', expand=False).notna()]



####export data####

with pd.ExcelWriter(final_name+'.xlsx') as writer:  
    data.to_excel(writer, sheet_name='original')
    filter6.to_excel(writer, sheet_name='genel filtre')
    filter3.to_excel(writer,sheet_name='kp-275')
    filter5.to_excel(writer, sheet_name='kardiyomiyopati')
    filter1.to_excel(writer,sheet_name='ilaç hassasiyeti')
    filter2.to_excel(writer,sheet_name='spor genetiği')
    filter4.to_excel(writer, sheet_name='nutrigenetik')
   
print('Exported file: {}'.format(final_name))
print('DONE:)')

