import configparser
from os.path import join, exists, basename
from os import listdir, environ, makedirs, remove
from module_extract_data import *
import xlwings as xw
import fitz
from pymupdf4llm.helpers.get_text_lines import get_text_lines
import re
import shutil
import easyocr
import ssl
from pathlib import Path

class Periodo(object):
    def __init__(self, id):
        self.id = id
        self.aporte_extracto = 'NO LEIDO'
        self.retiro_extracto = 'NO LEIDO'
        self.retefuente_extracto = 'NO LEIDO'
        self.rendimientos_extracto = 'NO LEIDO'
        self.saldo_final_extracto = 'NO LEIDO'
        self.aporte_sifi = 0
        self.retiro_sifi = 0

class Negocio(object):
    def __init__(self, id, nombre, nit):
        self.id = id
        self.nombre = nombre
        self.nit = nit
        self.periodos = []
    
    def add_periodo(self, id):
        self.periodos.append(Periodo(id))
    
    def update_periodo_extracto(self, data):
        for periodo in self.periodos:
            if periodo.id == data[0]:
                
                periodo.aporte_extracto = data[1]
                periodo.retiro_extracto = data[2]
                periodo.retefuente_extracto = data[3]
                periodo.rendimientos_extracto = data[4]
                periodo.saldo_final_extracto = data[5]
    
    def update_periodo_sifi(self, data):
        for periodo in self.periodos:
            if periodo.id == data[0]:
                
                periodo.aporte_sifi = data[1]
                periodo.retiro_sifi = data[2]


class Controller(object):

    def __init__(self):
        self.folder_path = self.find_path('BOT CONCILIACION CARTERA EXTERNA')
        self.read_config()
        self.data_manager = DataManager()
        self.get_cias_info()
        self.generate_table_all_movements()
        self.save_to_excel()
        self.extract_from_pdf()
        self.save_extract_info()

    def find_path(self, target, start_path=None):
        if start_path is None:
            profile_path = environ.get('USERPROFILE')
            if not profile_path:
                print("Error: Variable de entorno 'USERPROFILE' no encontrada.")
                return None
            start_dir = Path(profile_path)
        else:
            start_dir = Path(start_path)

        if not start_dir.is_dir():
            print(f"Error: La ruta de inicio '{start_dir}' no es un directorio válido.")
            return None
        try:
            for potential_match in start_dir.rglob(target):
                if potential_match.is_file():
                    print(f"¡Archivo encontrado!: {potential_match}")
                    return potential_match
                elif potential_match.is_dir():
                    print(f"¡Directorio encontrado!: {potential_match}")
                    return potential_match
        except PermissionError:
            print(f"Se encontró un error de permisos durante la búsqueda en '{start_dir}' o subdirectorios.")
        except Exception as e:
            print(f"Ocurrió un error inesperado durante la búsqueda: {e}")

        print(f"Archivo '{target}' no encontrado.")
        return None

    def save_to_excel(self):
        with pd.ExcelWriter(join(self.folder_path, 'Conciliacion.xlsx')) as writer:  
            self.data_manager.all_movements.to_excel(writer, index = None, sheet_name = 'Mvtos')

            for cias in self.list_cias:
                df = pd.DataFrame()
                name = str(cias.id) + '-' + cias.nombre.replace('PATRIMONIO AUTONOMO', 'PA')
                df.to_excel(writer, index = None, sheet_name = name, header = None)                  

    def read_config(self):
        config = configparser.ConfigParser()
        config.read(join(self.folder_path, 'conf.ini'),  encoding= "utf-8")

        dict = {
        'date_from' : config['Configuraciones']['Periodo Inicial'],
        'date_to' : config['Configuraciones']['Periodo Final'],
        'cias_list' : config['Configuraciones']['Códigos PA']
        }

        dict['cias_list'] = dict['cias_list'].replace('[','').replace(']','').replace(',','').split(' ')
        self.conf_variables = dict

    def save_config(self):
        config = configparser.ConfigParser()
        config['Configuraciones'] = {'Códigos PA' : [119656,102699,93235,102022,71316,98080,116047,106295,106919],
                                     'Periodo Inicial': '202407',
                                     'Periodo Final': '202411'}

        with open(join(self.folder_path, 'conf.ini'), 'w', encoding= "utf-8") as configfile:
            config.write(configfile)

    def get_cias_info(self):
        self.list_cias = []

        df = self.data_manager.get_cias_info(self.conf_variables['cias_list'])
        for id, d in enumerate(df['CODIGO'].values()):

            cias = Negocio(d, df['NOMBRE'][id], df['NIT'][id])

            inicial = self.conf_variables['date_from']
            final = self.conf_variables['date_to']

            anho = inicial[:4]
            mes = inicial[4:]

            while (anho + mes) != final:
                cias.add_periodo(anho+mes)
                mes = int(mes) + 1

                if mes > 12:
                    mes = 1
                    anho = int(anho) + 1
                
                if mes < 10:
                    mes = '0' + str(mes)
                else:
                    mes = str(mes)
                
                anho = str(anho) 

            cias.add_periodo(final)

            self.list_cias.append(cias)

    def generate_table_all_movements(self):
        self.data_manager.get_all_movements(self.conf_variables['date_from'], self.conf_variables['date_to'], self.conf_variables['cias_list'])
        for cias in self.list_cias:
            for period in cias.periodos:
                ##RETIROS
                df = self.data_manager.all_movements.loc[(self.data_manager.all_movements['NEGOCIO'] == cias.id) & (self.data_manager.all_movements['PERIODO'] == int(period.id)) & (self.data_manager.all_movements['TIPO'] == 'RE')]
                total_retiros = abs(df['VALOR'].sum())

                ##APORTES
                df2 = self.data_manager.all_movements.loc[(self.data_manager.all_movements['NEGOCIO'] == cias.id) & (self.data_manager.all_movements['PERIODO'] == int(period.id)) & (self.data_manager.all_movements['TIPO'] == 'AP')]
                total_aportes = abs(df2['VALOR'].sum())

                cias.update_periodo_sifi([period.id, total_aportes, total_retiros])
                
    def save_extract_info(self):
        book = xw.Book(join(self.folder_path, 'Conciliacion.xlsx'))
        
        for cias in self.list_cias:
            i = 14
            name = str(cias.id) + '-' + cias.nombre.replace('PATRIMONIO AUTONOMO', 'PA')
            sheet = xw.sheets[name]

            for p in cias.periodos:
                try:
                    dif_aporte = p.aporte_extracto - p.aporte_sifi
                except:
                    dif_aporte = 0
                try:
                    dif_retiro = p.retiro_extracto - p.retiro_sifi
                except:
                    dif_retiro = 0

                sheet[f'A{i}'].value = p.id
                sheet[f'B{i}'].value = 'EXTRACTO'
                sheet[f'C{i}'].value = 'SIFI'
                sheet[f'D{i}'].value = 'DIFERENCIA'
                sheet[f'E{i}'].value = 'OBSERVACIÓN'
                sheet[f'A{(i+1)}'].value = 'Aporte'
                sheet[f'B{(i+1)}'].value = p.aporte_extracto
                sheet[f'C{(i+1)}'].value = p.aporte_sifi
                sheet[f'D{(i+1)}'].value = dif_aporte
                sheet[f'A{(i+2)}'].value = 'Retiro'
                sheet[f'B{(i+2)}'].value = p.retiro_extracto
                sheet[f'C{(i+2)}'].value = p.retiro_sifi
                sheet[f'D{(i+2)}'].value = dif_retiro
                sheet[f'A{(i+3)}'].value = 'Retefuente'
                sheet[f'B{(i+3)}'].value = p.retefuente_extracto
                sheet[f'A{(i+4)}'].value = 'Rendimientos'
                sheet[f'B{(i+4)}'].value = p.rendimientos_extracto
                sheet[f'A{(i+5)}'].value = 'Saldo final'
                sheet[f'B{(i+5)}'].value = p.saldo_final_extracto
                sheet[f'A{(i+6)}'].value =  'Observación'

                sheet[f'A{i}'].font.bold = True
                sheet[f'B{i}'].font.bold = True
                sheet[f'C{i}'].font.bold = True
                sheet[f'A{(i+6)}'].font.bold = True

                sheet[f'A{(i+3)}'].font.color = "#FF0000"
                sheet[f'B{(i+3)}'].font.color = "#FF0000"

                if dif_aporte != 0:
                     sheet[f'D{(i+1)}'].color = "#FFFF00"
                if dif_retiro != 0:
                     sheet[f'D{(i+2)}'].color = "#FFFF00"

                i+=8
            
            sheet[f'A{2}'].value = 'RESUMEN'
            sheet[f'A{5}'].value = 'APORTES'
            sheet[f'A{6}'].value = 'RETIROS'
            sheet[f'A{7}'].value = 'RETEFUENTE'
            sheet[f'A{8}'].value = 'RENDIMIENTOS'
            sheet[f'A{9}'].value = 'SALDO FINAL'
            sheet[f'A{10}'].value = 'APORTE PENDIENTE'
            sheet[f'A{11}'].value = 'RETIRO PENDIENTE'

            sheet[f'A{7}'].font.color = "#FF0000"
            sheet[f'A{2}'].color = "#FFFF00"

            c = 0
            
            column_list = ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L' ,'M']
            for p in cias.periodos:
                j = 4
                sheet[f'{column_list[c]}{j}'].value = p.id

                try:
                    dif_aporte = p.aporte_extracto - p.aporte_sifi
                except:
                    dif_aporte = 0
                try:
                    dif_retiro = p.retiro_extracto - p.retiro_sifi
                except:
                    dif_retiro = 0

                sheet[f'{column_list[c]}{(j+1)}'].value = p.aporte_extracto
                sheet[f'{column_list[c]}{(j+2)}'].value = p.retiro_extracto
                sheet[f'{column_list[c]}{(j+3)}'].value = p.retefuente_extracto
                sheet[f'{column_list[c]}{(j+4)}'].value = p.rendimientos_extracto
                sheet[f'{column_list[c]}{(j+5)}'].value = p.saldo_final_extracto
                sheet[f'{column_list[c]}{(j+6)}'].value = dif_aporte
                sheet[f'{column_list[c]}{(j+7)}'].value = dif_retiro

                sheet[f'{column_list[c]}{(j+3)}'].font.color = "#FF0000"
                if dif_aporte != 0:
                    sheet[f'{column_list[c]}{(j+6)}'].color = "#FFFF00"
                if dif_retiro != 0:
                    sheet[f'{column_list[c]}{(j+7)}'].color = "#FFFF00"

                c+=1

        book.save()
        book.close()

    def find_cias_by_id(self, id):
        for cias in self.list_cias:
            if id == cias.id:
                return cias
        return -1

    def extract_data(self, pdf):
        text = pdf[-1].get_text("text")
        if 'Detalle Movimientos' in text:
            return self.read_model3(pdf[0], pdf[-1])      
        elif 'bancolombia.com' in text:
            return self.read_model1(pdf[-1])
        else:
            return self.read_model2(pdf[-1])
        
    def convert_to_float(self, num):
        if num[-3] == '.' or num[-3] == ',':
            num = num[:-3].replace('.', '').replace(',','') + '.' + num[-2:]
        else:
            num = num.replace('.', '').replace(',','')
        return float(num)

    def process_period(self, cad):
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
        if 'AGO' in dates[1].upper() or 'AG0'in dates[1].upper():
            return dates[2] + '08'
        if 'SEP' in dates[1].upper():
            return dates[2] + '09'
        if 'OCT' in dates[1].upper() or '0CT'in dates[1].upper():
            return dates[2] + '10'
        if 'NOV' in dates[1].upper() or 'N0V'in dates[1].upper():
            return dates[2] + '11'
        if 'DIC' in dates[1].upper() or 'D1C'in dates[1].upper():
            return dates[2] + '12'

        return 'ERROR'

    def process_numeric_line(self, line):
        try:
            rex = r"(\d+.*\d+)"
            match = re.search(rex, line)
            if match:
                return self.convert_to_float(match[0].replace(' ',''))
        except:
            return 'ERROR'

    def read_model1(self, page):
        data = []

        lines = get_text_lines(page).split('\n')
        search_list = {'ADICIONES': 2, 'RETIROS': 3, 'REND. NETOS': 0, 'RETENCIÓN': 1, 'NUEVO SALDO': 2}
        for x, line in enumerate(lines):

            if 'DESDE' in line.upper() and 'HASTA' in line.upper():
                data.append(re.search(r'\d\d\d\d\d\d', line)[0])
            else:
                for w in search_list.keys():
                    if w in line:
                        try:
                            num = self.convert_to_float(lines[x+4].split('\t')[search_list[w]])
                            search_list[w] = num
                        except:
                            search_list[w] = 'ERROR'
        
        data.append(search_list['ADICIONES'])
        data.append(search_list['RETIROS'])
        data.append(search_list['RETENCIÓN'])
        data.append(search_list['REND. NETOS'])
        data.append(search_list['NUEVO SALDO'])

        return data

    def read_model2(self, page):
        periodo = 'ERROR'
        aporte = 'ERROR'
        retiro = 'ERROR'
        rend = 'ERROR'
        retefuente = 'ERROR'
        saldo_final = 'ERROR'

        lines = get_text_lines(page).split('\n')
        for x, line in enumerate(lines):
            if 'PERIODO' in line.upper() and periodo == 'ERROR':
                match = re.search(r"\d\d-(.*?)-\d\d\d\d", line)
                if match:
                    periodo = self.process_period(match[0])
            elif 'APORTES' in line.upper() and aporte == 'ERROR':
                aporte = self.process_numeric_line(line)

            elif 'RETIROS' in line.upper() and retiro == 'ERROR':
                retiro = self.process_numeric_line(line)

            elif 'RETENCIÓN EN LA FUENTE' in line.upper() and retefuente == 'ERROR':
                retefuente = self.process_numeric_line(line)

            elif 'RENDIMIENTOS NETOS' in line.upper() and rend == 'ERROR':
                rend = self.process_numeric_line(line)

            elif 'SALDO FINAL' in line.upper() and saldo_final == 'ERROR':
                saldo_final = self.process_numeric_line(line)

               
        return [periodo, aporte, retiro, retefuente, rend, saldo_final]

    def read_model3(self, page1, page2):
        periodo = 'ERROR'
        aporte = 'ERROR'
        retiro = 'ERROR'
        rend = 'ERROR'
        retefuente = 'ERROR'
        saldo_final = 'ERROR'

        try:
            tabs = page2.find_tables()
            if tabs.tables:
                date = tabs[1].extract()[-2][1].split('/')
                periodo = str(date[2] + date[1])
                aporte = self.process_numeric_line(tabs[1].extract()[-1][3])
                retiro = self.process_numeric_line(tabs[1].extract()[-1][4])
                retefuente = self.process_numeric_line(tabs[1].extract()[-1][6])
                saldo_final = self.process_numeric_line(tabs[1].extract()[-1][8])
            
            tabs = page1.find_tables()
            if tabs.tables:
                rend = self.process_numeric_line(tabs[0].extract()[-1][5])
        except:
            ['ERROR', 'ERROR', 'ERROR', 'ERROR', 'ERROR' 'ERROR']

        return [periodo, aporte, retiro, retefuente, rend, saldo_final]

    def read_model_ocr(self, page):
        periodo = 'ERROR'
        aporte = 'ERROR'
        retiro = 'ERROR'
        rend = 'ERROR'
        retefuente = 'ERROR'
        saldo_final = 'ERROR'

        for x, line in enumerate(page):
            if 'PERIODO' in line.upper() and periodo == 'ERROR':
                periodo = page[x+1]
                periodo = self.process_period(periodo)
            elif 'APORTES' in line.upper() and aporte == 'ERROR':
                aporte = page[x+1][1:]
                aporte = self.process_numeric_line(aporte)

            elif 'RETIROS' in line.upper() and retiro == 'ERROR':
                retiro = page[x+1][1:]
                retiro = self.process_numeric_line(retiro)

            elif 'RETENCIÓN EN LA FUENTE' in line.upper() and retefuente == 'ERROR':
                retefuente = page[x+1][1:]
                retefuente = self.process_numeric_line(retefuente)

            elif 'RENDIMIENTOS NETOS' in line.upper() and rend == 'ERROR':
                rend = page[x+1][1:]
                rend = self.process_numeric_line(rend)

            elif 'SALDO FINAL' in line.upper() and saldo_final == 'ERROR':
                saldo_final = page[x+1][1:]
                saldo_final = self.process_numeric_line(saldo_final)

               
        return [periodo, aporte, retiro, retefuente, rend, saldo_final]

    def extract_from_pdf(self):
            base_path = join(self.folder_path, 'Lista Extractos')
            pdf_not_native = []
            for doc in listdir(base_path):
                if doc.upper().endswith('PDF'):
                    for cod in self.conf_variables['cias_list']:
                        if cod in doc:
                            try:
                                cias = self.find_cias_by_id(int(cod))
                                pdf = fitz.Document(join(join(self.folder_path, 'Lista Extractos'), doc))
                                if pdf.needs_pass:
                                    pdf.authenticate(str(cias.nit))
                                
                                data = self.extract_data(pdf)
                                if not 'ERROR' in data:
                                    cias.update_periodo_extracto(data)
                                    print(f'Se leyó correctamente el archivo: {doc}')
                                else:
                                    pdf_not_native.append(cias)
                                    pdf_not_native.append(doc)
                            except:
                                pdf_not_native.append(cias)
                                pdf_not_native.append(doc)
                            finally:
                                break
            if len(pdf_not_native) > 0:
                self.read_with_ocr(pdf_not_native)

    def read_with_ocr(self, pdf_list):
        base_path = join(self.folder_path, 'Lista Extractos')
        pos = 0
        ssl._create_default_https_context = ssl._create_stdlib_context
        reader = easyocr.Reader(['es'])

        while pos < len(pdf_list):
            try:
                path = self.pdf_to_images(pdf_list[pos+1], pdf_list[pos])
                result = reader.readtext(path, detail=0)
                print(result)
                data = self.read_model_ocr(result)
                if 'ERROR' in data:
                    self.file_not_read(base_path, pdf_list[pos+1])
                    continue
                else:
                    print(data)
                    pdf_list[pos].update_periodo_extracto(data)
                    print(f'Se leyó correctamente el archivo: {pdf_list[pos+1]}')
                
            except:
                self.file_not_read(base_path, pdf_list[pos+1])
                continue
            finally:
                pos+=2
                remove(path)
            
    def pdf_to_images(self, doc, cias):
        pdf = fitz.Document(join(join(self.folder_path, 'Lista Extractos'), doc))
        if pdf.needs_pass:
            pdf.authenticate(str(cias.nit))                     
    
        val = join(join(self.folder_path, 'Lista Extractos'), f"image_{doc}.png")
        page = pdf[-1]
        zoom = 4
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat)
        pix.save(val)
        pdf.close()
        return val

    def file_not_read(self, base_path, doc):
        path = join(self.folder_path, 'NO LEIDOS')
        if not exists(path):
            makedirs(path)                     
        shutil.copyfile(join(base_path, doc), join(path, doc))
        print(f'No se pudo leer el documento: {doc}')

controller = Controller()
#controller.save_config()
