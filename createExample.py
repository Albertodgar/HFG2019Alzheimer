import os
import pandas as pd
import math

def load_data():
    cal = pd.read_csv('./calendar.csv', sep=';')
    cap = pd.read_excel('./capabilities.xlsx',
                    sheet_name='ENE-FEB', header=None)
    cap = cap.drop(cap.index[0])
    lim_col = None
    for col in list(cap):
        if isinstance(cap[col][1], float) and math.isnan(cap[col][1]):
            lim_col = col
            break
    cap.drop(cap.columns[list(range(lim_col,len(cap.columns)-1))],
                     axis=1, inplace=True)
    cap.columns = cap.ix[1,:]
    cap.drop(cap.columns[len(cap.columns)-1], axis=1, inplace=True)
    cap.drop(cap.index[0], inplace=True)
    cap.dropna(axis=0, how='any', inplace=True)
    cap.columns = [a.strip() for a in cap]
    return cal, cap

def create():    
    asignaturas = list(cap.columns)[4:]
    niveles = ['A','B','M']
    os.system('mkdir -p ./FICHAS')
    for a in asignaturas:    
        mkdir = 'mkdir -p ./FICHAS/'+a
        os.system(mkdir)
        for n in niveles:
            niveles_mkdir = mkdir + '/'+n
            os.system(niveles_mkdir)
            for i in range(1,6):
                path = niveles_mkdir[9:]
                pdf_cp = 'cp ./ejercicio.pdf ' + path 
                pdf_cp = pdf_cp + '/'+a+'_'+n+'_'+str(i)+'.pdf'
                pdf_cp = os.system(pdf_cp)
    dias = list(cal['DIA'])
    for d in dias:
        s = 'mkdir -p ./IMPRIMIR/'+d 
        os.system(s)
        for sala in ['SALA1','SALA2','SALA3']:
            s2 = s + '/'+sala
            os.system(s2)
                    
cal, cap = load_data()
create()