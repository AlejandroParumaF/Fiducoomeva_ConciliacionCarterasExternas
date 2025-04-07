from os import listdir
from os.path import join
import fitz
import pandas as pd
from pymupdf4llm.helpers.get_text_lines import get_text_lines
import re

cod_dict = {

'119656':	{'name': 'PATRIMONIO AUTONOMO LA HIPOTECARIA', 'nit':	901817824},
'102699':	{'name': 'PATRIMONIO AUTONOMO ACCIAL - AVISTA', 'nit':	901061400},
'93235':	{'name': 'PATRIMONIO AUTONOMO ACCECREDITOS', 'nit':	900951022},
'102022':	{'name': 'PA VALORENZ', 'nit':	901541567},
'71316':	{'name': 'PATRIMONIO AUTONOMO AVISTA', 'nit':	901086223},
'98080':	{'name': 'PATRIMONIO AUTONOMO VALORANZ', 'nit':	901473216},
'116047':	{'name': 'PATRIMONIO AUTONOMO FINEXUS', 'nit':	901752630},
'106295':	{'name': 'PATRIMONIO AUTONOMO CREDIALIANZA', 'nit':	901588750},
'106919':	{'name': 'PATRIMONIO AUTONOMO ACTIVOS Y FINANZAS', 'nit':	901600175}
}

def convert_to_float(num):
    if num[-3] == '.' or num[-3] == ',':
        num = num[:-3].replace('.', '').replace(',','') + '.' + num[-2:]
    else:
        num = num.replace('.', '').replace(',','')
    return float(num)

def process_period(cad):
    dates = cad.split('-')

    if 'ENE' in dates[1].upper():
        return dates[2] + '01'
    if 'FEB' in dates[1].upper():
        return dates[2] + '02'
    if 'MAR' in dates[1].upper():
        return dates[2] + '03'
    if 'ABR' in dates[1].upper():
        return dates[2] + '04'
    if 'MAY' in dates[1].upper():
        return dates[2] + '05'
    if 'JUN' in dates[1].upper():
        return dates[2] + '06'
    if 'JUL' in dates[1].upper():
        return dates[2] + '07'
    if 'AGO' in dates[1].upper():
        return dates[2] + '08'
    if 'SEP' in dates[1].upper():
        return dates[2] + '09'
    if 'OCT' in dates[1].upper():
        return dates[2] + '10'
    if 'NOV' in dates[1].upper():
        return dates[2] + '11'
    if 'DIC' in dates[1].upper():
        return dates[2] + '12'

    return 'ERROR'

def process_numeric_line(line):
    try:
        rex = r"(\d+.*\d+)"
        match = re.search(rex, line)
        if match:
            return convert_to_float(match[0])
    except:
        return 'ERROR'

def read_model1(page, cod):
    data = [cod_dict[cod]['nit'], cod, cod_dict[cod]['name']]

    lines = get_text_lines(page).split('\n')
    search_list = {'ADICIONES': 2, 'RETIROS': 3, 'REND. NETOS': 0, 'RETENCIÓN': 1, 'NUEVO SALDO': 2}
    for x, line in enumerate(lines):

        if 'DESDE' in line.upper() and 'HASTA' in line.upper():
            data.append(re.search(r'\d\d\d\d\d\d', line)[0])
        else:
            for w in search_list.keys():
                if w in line:
                    try:
                        num = convert_to_float(lines[x+4].split('\t')[search_list[w]])
                        search_list[w] = num
                    except:
                        search_list[w] = 'ERROR'
    
    data.append(search_list['ADICIONES'])
    data.append(search_list['REND. NETOS'])
    data.append(search_list['RETIROS'])
    data.append(search_list['RETENCIÓN'])
    data.append(search_list['NUEVO SALDO'])

    return data

def read_model2(page, cod):
    data = [cod_dict[cod]['nit'], cod, cod_dict[cod]['name']]

    lines = get_text_lines(page).split('\n')
    for x, line in enumerate(lines):
        #print(line)
        if 'PERIODO' in line.upper():
            match = re.search(r"\d\d-(.*?)-\d\d\d\d", line)
            if match:
                data.append(process_period(match[0]))
        elif 'APORTES' in line.upper():
            match = process_numeric_line(line)
            data.append(match)
        elif 'RETIROS' in line.upper():
            match = process_numeric_line(line)
            data.append(match)   
        elif 'RETENCIÓN EN LA FUENTE' in line.upper():
            match = process_numeric_line(line)
            data.append(match)
        elif 'RENDIMIENTOS NETOS' in line.upper():
            match = process_numeric_line(line)
            data.append(match)
        elif 'SALDO FINAL' in line.upper():
            match = process_numeric_line(line)
            data.append(match)       
    return data

#Leer PDFs desde una carpeta
def extract_data(page, cod):
    text = page.get_text("text")
    if 'bancolombia.com' in text:
        return read_model1(page, cod)
    else:
        return read_model2(page, cod)
    
def read_pdfs():
    df = pd.DataFrame(columns=['NIT', 'Código', 'Nombre', 'Periodo', 'Aporte', 'Rendimiento', 'Retiro', 'Retefuente', 'Saldo Final'])

    for doc in listdir('TEST'):
        if doc.upper().endswith('PDF'):
            for cod in cod_dict.keys():
                if cod in doc:
                    try:
                        #print(f'Documento: {doc} pertenece a {cod_dict[cod]['name']}')
                        pdf = fitz.Document(join(r'C:\Users\alpf0069\Documents\Proyectos\Area Contabilidad\Conciliación Cartera Externa\TEST', doc))
                        if pdf.needs_pass:
                            #print('Requiere password')
                            pdf.authenticate(str(cod_dict[cod]['nit']))
                        
                        data = extract_data(pdf[-1], cod)
                        print(data)
     
                        df.loc[len(df)] = data
                    except:
                        print(f'No se pudo leer el documento: {doc}')
                        continue   
    df.to_excel('Res.xlsx', index = False)




