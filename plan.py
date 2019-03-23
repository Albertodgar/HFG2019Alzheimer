import pandas as pd
import math
from PyPDF2 import PdfFileWriter, PdfFileReader
import io
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import os
import random
import time

def write_to_PDF(usuario, input_file, output_file):
    packet = io.BytesIO()
    
    # create a new PDF with Reportlab
    can = canvas.Canvas(packet, pagesize=letter)
    can.drawString(140, 706, usuario)
    can.save()
    
    # move to the beginning of the StringIO buffer
    packet.seek(0)
    new_pdf = PdfFileReader(packet)
    # read your existing PDF
    existing_pdf = PdfFileReader(open(input_file, "rb"))
    output = PdfFileWriter()
    # add the "watermark" (which is the new pdf) on the existing page
    page = existing_pdf.getPage(0)
    page2 = new_pdf.getPage(0)
    page.mergePage(page2)
    output.addPage(page)
    # finally, write "output" to a real file
    outputStream = open(output_file, "wb")
    output.write(outputStream)
    outputStream.close()

def capabilities(day):
    d = dict()
    cols = [list(cal[i]) for i in cal]
    day_index = cols[0].index(day)
    x = [i[day_index] for i in cols if not str(i[day_index]) == 'nan' ][2::2]
    for i in x:
        d[i] = list(cap[i])
    return d

def cast_cap(cap):
    if cap == 'A' or cap == 'A-M' or cap == 'M-A':
        return 'A'
    elif cap == 'M' or cap == 'M-B' or cap == 'B-M' or 'B_M':
        return 'M'
    elif cap == 'B' or cap == 'MB':
        return 'B'
    else:
        return cap


def name_from_path(path):
    path = path[::-1]
    res = ''
    for i in range(len(path)):
        if path[i] == '/':
            return res[::-1]
        res = res + path[i]
    return None
        

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

start_time = time.time()
      
cal, cap = load_data()
alumnos = list(cap['USUARIOS'])
salas = list(cap['SALA'])
salas_dict = {i:[] for i in range(4)}
for i in range(len(salas)):
    j = salas[i]
    salas_dict[j].append(alumnos[i])
    

for d in list(cal['DIA']):
    asignaturas = capabilities(d)
    for a in asignaturas:
        habilidades = asignaturas[a]
        for i in range(1,4):
            l_alumnos = salas_dict[i]
            for alum in l_alumnos: 
                index_alumno = alumnos.index(alum)
                cap_alum = habilidades[index_alumno]
                cap_alum = cast_cap(cap_alum)
                if cap_alum == 'A' or cap_alum == 'B' or cap_alum == 'M':
                    path = './FICHAS/'+a+'/'+ cap_alum
                    l = os.listdir(path)
                    path = path + '/' + random.choice(l)
                    nombre2 = alum + '_' + name_from_path(path)
                    path2 = "./IMPRIMIR/"
                    path2 = path2 + "{}/SALA{}/{}".format(d, str(i), nombre2)
                    #print(path,path2)
                    write_to_PDF(alum, path, path2)

t = time.time() - start_time
print("Tiempo en ejecutarse para {}: {} segundos".format(len(alumnos),t))
print("Tiempo por usuario: {}".format(t/len(alumnos)))




