import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import oracledb
from pathlib import Path
import pymupdf as fitz
import re
from pymupdf4llm.helpers.get_text_lines import get_text_lines
import easyocr

class BotConciliacion:
    def __init__(self, root):
        self.root = root
        self.root.title("Bot Conciliación Carteras Externas")
        self.root.geometry("550x550")
        self.root.resizable(False, False)
        self.codigos = []  # Lista para almacenar los códigos

        # Configuración de estilos
        self.configurar_estilos()

        # Frame principal
        self.main_frame = ttk.Frame(root, padding="15")
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        # Título
        self.crear_titulo()

        # Sección de gestión de códigos
        self.crear_seccion_codigos()

        # Sección de parámetros de conciliación
        self.crear_seccion_parametros()

        # Botón Iniciar
        self.crear_boton_iniciar()

    def configurar_estilos(self):
        estilo = ttk.Style()
        estilo.theme_use('clam')
        
        COLOR_PRIMARIO = '#4a6fa5'
        COLOR_SECUNDARIO = '#6c757d'
        COLOR_TEXTO_CLARO = '#ffffff'
        
        estilo.configure('Titulo.TLabel', font=('Segoe UI', 16, 'bold'), foreground='#2c3e50')
        estilo.configure('BotonPrincipal.TButton', font=('Segoe UI', 10, 'bold'), 
                        background=COLOR_PRIMARIO, foreground=COLOR_TEXTO_CLARO, padding=10)
        estilo.configure('BotonSecundario.TButton', font=('Segoe UI', 9), 
                        background=COLOR_SECUNDARIO, foreground=COLOR_TEXTO_CLARO, padding=5)
        
        estilo.map('BotonPrincipal.TButton', background=[('active', '#3a5a80')])
        estilo.map('BotonSecundario.TButton', background=[('active', '#5a6268')])

    def crear_titulo(self):
        titulo = ttk.Label(self.main_frame, 
                          text="BOT CONCILIACIÓN CARTERAS EXTERNAS", 
                          style='Titulo.TLabel')
        titulo.grid(row=0, column=0, columnspan=3, pady=(0, 15))

    def crear_seccion_codigos(self):
        # Frame para la sección de códigos
        frame_codigos = ttk.LabelFrame(self.main_frame, text="Listado de PAs", padding=10)
        frame_codigos.grid(row=1, column=0, columnspan=3, sticky='ew', pady=10, padx=5)

        # Entrada para nuevo código
        ttk.Label(frame_codigos, text="Nuevo Código:").grid(row=0, column=0, sticky='w')
        self.entry_codigo = ttk.Entry(frame_codigos, width=25)
        self.entry_codigo.grid(row=0, column=1, padx=5, sticky='we')

        # Botón para agregar
        btn_agregar = ttk.Button(frame_codigos, text="Agregar", 
                                style='BotonSecundario.TButton',
                                command=self.agregar_codigo)
        btn_agregar.grid(row=0, column=2, padx=5)

        # ListBox para mostrar códigos
        self.listbox = tk.Listbox(frame_codigos, height=6, selectmode=tk.SINGLE)
        self.listbox.grid(row=1, column=0, columnspan=2, pady=10, sticky='nsew')

        # Scrollbar
        scrollbar = ttk.Scrollbar(frame_codigos, orient=tk.VERTICAL, command=self.listbox.yview)
        scrollbar.grid(row=1, column=2, sticky='ns')
        self.listbox.config(yscrollcommand=scrollbar.set)

        # Botón para eliminar
        btn_eliminar = ttk.Button(frame_codigos, text="Eliminar Seleccionado", 
                                style='BotonSecundario.TButton',
                                command=self.eliminar_codigo)
        btn_eliminar.grid(row=2, column=0, columnspan=3, pady=5, sticky='we')

        # Configurar expansión
        frame_codigos.columnconfigure(1, weight=1)

    def crear_seccion_parametros(self):
        # Frame para parámetros
        frame_parametros = ttk.LabelFrame(self.main_frame, text="Información Necesaria", padding=10)
        frame_parametros.grid(row=2, column=0, columnspan=3, sticky='ew', pady=10, padx=5)

        # Campo Carpeta Extractos
        ttk.Label(frame_parametros, text="Carpeta Extractos:").grid(row=0, column=0, sticky='w', pady=5)
        self.entrada_carpeta = ttk.Entry(frame_parametros, width=40)
        self.entrada_carpeta.grid(row=0, column=1, padx=5, sticky='ew')
        boton_carpeta = ttk.Button(frame_parametros, text="Seleccionar", 
                                 style='BotonSecundario.TButton',
                                 command=self.seleccionar_carpeta)
        boton_carpeta.grid(row=0, column=2, padx=(5, 0))

        # Campo Archivo Datos Dev Retefte
        ttk.Label(frame_parametros, text="Archivo Datos Dev Retefte:").grid(row=1, column=0, sticky='w', pady=5)
        self.entrada_archivo = ttk.Entry(frame_parametros, width=40)
        self.entrada_archivo.grid(row=1, column=1, padx=5, sticky='ew')
        boton_archivo = ttk.Button(frame_parametros, text="Seleccionar", 
                                 style='BotonSecundario.TButton',
                                 command=self.seleccionar_archivo)
        boton_archivo.grid(row=1, column=2, padx=(5, 0))

        # Campo Periodo
        ttk.Label(frame_parametros, text="Periodo:").grid(row=2, column=0, sticky='w', pady=5)
        self.entrada_periodo = ttk.Entry(frame_parametros, width=15)
        self.entrada_periodo.grid(row=2, column=1, sticky='w', padx=5)

        # Configurar expansión
        frame_parametros.columnconfigure(1, weight=1)

    def crear_boton_iniciar(self):
        boton_iniciar = ttk.Button(self.main_frame, 
                                 text="INICIAR PROCESO", 
                                 style='BotonPrincipal.TButton', 
                                 command=self.iniciar_proceso)
        boton_iniciar.grid(row=3, column=0, columnspan=3, pady=(15, 0), sticky='ew')

    def agregar_codigo(self):
        nuevo_codigo = self.entry_codigo.get().strip()
        if nuevo_codigo:
            if nuevo_codigo not in self.codigos:
                self.codigos.append(nuevo_codigo)
                self.actualizar_listbox()
                self.entry_codigo.delete(0, tk.END)
            else:
                messagebox.showwarning("Advertencia", "El código ya existe en la lista")
        else:
            messagebox.showwarning("Advertencia", "Ingrese un código válido")

    def eliminar_codigo(self):
        seleccionado = self.listbox.curselection()
        if seleccionado:
            index = seleccionado[0]
            self.codigos.pop(index)
            self.actualizar_listbox()
        else:
            messagebox.showwarning("Advertencia", "Seleccione un código para eliminar")

    def actualizar_listbox(self):
        self.listbox.delete(0, tk.END)
        for codigo in self.codigos:
            self.listbox.insert(tk.END, codigo)

    def seleccionar_carpeta(self):
        carpeta = filedialog.askdirectory()
        if carpeta:
            self.entrada_carpeta.delete(0, tk.END)
            self.entrada_carpeta.insert(0, carpeta)

    def seleccionar_archivo(self):
        archivo = filedialog.askopenfilename()
        if archivo:
            self.entrada_archivo.delete(0, tk.END)
            self.entrada_archivo.insert(0, archivo)

    def validar_campos(self):
        errores = []
        if not self.entrada_carpeta.get():
            errores.append("Debe seleccionar la Carpeta Extractos")
        if not self.entrada_archivo.get():
            errores.append("Debe seleccionar el Archivo Datos Dev Retefte")
        if not self.entrada_periodo.get():
            errores.append("Debe ingresar el Periodo")
        if not self.codigos:
            errores.append("Debe agregar al menos un código")
        return errores

    def iniciar_proceso(self):
        errores = self.validar_campos()
        
        if errores:
            messagebox.showerror("Error", "\n".join(errores))
        else:
            # Mostrar resumen con los códigos
            resumen = (f"Proceso iniciado correctamente\n"
                      f"Periodo: {self.entrada_periodo.get()}\n"
                      f"Carpeta: {self.entrada_carpeta.get()}\n"
                      f"Archivo: {self.entrada_archivo.get()}\n"
                      f"\nCódigos a procesar ({len(self.codigos)}):\n" + 
                      "\n".join(f"- {codigo}" for codigo in self.codigos))
            
            self.df_retfte = self.leer_datos_rtef()
            self.df_database = self.leer_datos_database()
            self.df_extractos = self.leer_datos_extractos()

            self.generar_conciliacion(self.df_retfte, self.df_database, self.df_extractos)

    def consultar_nit(self, codigo):
        nit = self.df_database.loc[self.df_database['CÓDIGO'] == codigo, 'NIT PA']
        if not nit.empty:
            return nit.values[0]
        else:
            messagebox.showerror("Error", f"NIT no encontrado para el código {codigo}")
            return None

    def generar_conciliacion(self, df_retfte, df_database, df_extractos):

        df_retfte['CÓDIGO'] = df_retfte['CÓDIGO'].astype(float)
        df_database['CÓDIGO'] = df_database['CÓDIGO'].astype(float)
        df_extractos['CÓDIGO'] = df_extractos['CÓDIGO'].astype(float)
        print(df_retfte)
        print(df_database)
        print(df_extractos)

        df1 = pd.merge(df_database, df_retfte, on='CÓDIGO', how='left')

        df2 = pd.merge(df1, df_extractos, on='CÓDIGO', how='left')

        df2['APORTES SIFI'] = df2['APORTES SIFI'].fillna(0)
        df2['RETIROS SIFI'] = df2['RETIROS SIFI'].fillna(0)

        df2['NIT PA'] = df2['NIT PA'].astype(int)
        df2['PERIODO'] = df2['PERIODO'].astype(int)
        df2['TERCERO'] = df2['TERCERO'].astype(int)

        df2['DIFERENCIA'] = df2['SALDO EXTRACTO'] - df2['SALDO CONTABLE']
        df2['APORTE PENDIENTE'] = df2['APORTES'] - df2['APORTES SIFI']
        df2['RETIROS PENDIENTES'] = df2['RETIROS'] - df2['RETIROS SIFI']
        df2['RETENCIONES NO PROCEDENTES'] = df2['RETENCIÓN EN LA FUENTE'] + df2['RETENCIONES NO PROCEDENTES']
        df2['AJUSTE EN RENDIMIENTOS'] = df2['DIFERENCIA'] - (df2['APORTE PENDIENTE'] - df2['RETIROS PENDIENTES'] -df2['RETENCIONES NO PROCEDENTES'])

        df2['RENDIMIENTO EN CONTRA'] = df2['AJUSTE EN RENDIMIENTOS'].apply(lambda valor: valor if valor < 0 else 0)
        df2['RENDIMIENTO A FAVOR'] = df2['AJUSTE EN RENDIMIENTOS'].apply(lambda valor: valor if valor > 0 else 0)

        df2.drop(columns=['APORTES SIFI', 'RETIROS SIFI', 'PERIODO', 'APORTES', 'RETIROS', 'RETENCIÓN EN LA FUENTE', 'RENDIMIENTOS NETOS'], inplace=True)

        df2 = df2[['NIT PA', 'CÓDIGO', 'NOMBRE ENCARGO', 'PATRIMONIO AUTÓNOMO', 'TERCERO', 'SALDO CONTABLE', 'SALDO EXTRACTO', 'DIFERENCIA', 'APORTE PENDIENTE', 'RETIROS PENDIENTES', 'RETENCIONES NO PROCEDENTES', 'AJUSTE EN RENDIMIENTOS', 'RENDIMIENTO EN CONTRA', 'RENDIMIENTO A FAVOR']]
        df2.to_excel('Conciliacion.xlsx', index=False)
    
    def leer_datos_rtef(self):
        df_retfte = pd.read_excel(self.entrada_archivo.get())

        try:
            df_retfte = df_retfte[['Codigo PA', 'TOTAL']]
            df_retfte['TOTAL'] = df_retfte['TOTAL'].astype(float)
            df_retfte = df_retfte.rename(columns={'Codigo PA': 'CÓDIGO', 'TOTAL': 'RETENCIONES NO PROCEDENTES'})
            
            return df_retfte
        except KeyError as e:
            messagebox.showerror("Error", f"El archivo Datos Dev Retefte no contiene las columnas Código PA y/o TOTAL {e}")
            return None
        except Exception as e:
            messagebox.showerror("Error", f"Error al leer el archivo: {e}")
            return None

    def leer_datos_database(self):
        config_sifi = {
                'user'           : "SIFI_RPT",
                'pwd'            : 'C00m3v4123*',
                'host'           : 'BDCDPORA11.INTRACOOMEVA.COM.CO',
                'port'           :  1554,
                'service_name'   : 'CDPORA11'
                }
        
        codigos = [codigo.strip() for codigo in self.codigos]
        cuenta = 130205010105
        periodo = self.entrada_periodo.get()

        df = pd.DataFrame()
      
        in_clause_placeholders = ", ".join([f":{i+3}" for i in range(len(codigos))])
        sql = f'''
                SELECT
                    CIAS_NRORIF AS "NIT PA",
                    CIAS_CIAS AS "CÓDIGO",
                    AUXI_DESCRI AS "NOMBRE ENCARGO",
                    CIAS_DESCRI AS "PATRIMONIO AUTÓNOMO",
                    AUXI_NIT AS "TERCERO",
                    SALD_SALACT AS "SALDO CONTABLE"
                FROM
                    SC_TSALD SC_TSALD
                    LEFT JOIN GE_TAUXIL ON SC_TSALD.SALD_AUXI = GE_TAUXIL.AUXI_AUXI
                    LEFT JOIN GE_TCIAS ON SALD_CIAS = CIAS_CIAS
                WHERE
                    SALD_MAYO = :1
                    AND AUXI_DESCRI NOT LIKE 'ENCARGO%'
                    AND SALD_FECMOV = :2
                    AND SALD_CIAS IN ({in_clause_placeholders}) -- Aquí se inserta :3, :4, ...
            '''    
        
        sql2 = f'''
            SELECT CÓDIGO, COALESCE(APORTES, 0) AS "APORTES SIFI", COALESCE(RETIROS, 0) AS "RETIROS SIFI" FROM(
                SELECT NEGOCIO AS "CÓDIGO", SUM(VALOR) AS VALOR, TIPO from(
                    SELECT
                        ge_tcias.cias_cias NEGOCIO,
                        sc_tmvco.mvco_mtoren VALOR,
                        CASE
                            WHEN mvco_descri LIKE '% 241' THEN 'AP'
                            WHEN mvco_descri LIKE '% 240' THEN 'AP'
                            WHEN mvco_descri LIKE '%ORDEN DE INVERSIÓN %' THEN 'AP'
                            WHEN mvco_descri LIKE '% 242' THEN 'RE'
                            WHEN mvco_te_tpmv = 34 THEN 'RE'
                        ELSE 'NA'
                        END AS TIPO
                    FROM
                        vu_sfi.sc_tmvco sc_tmvco
                        LEFT JOIN ge_tauxil ON sc_tmvco.mvco_auxi = ge_tauxil.auxi_auxi
                        LEFT JOIN ge_tcias ON sc_tmvco.mvco_cias = ge_tcias.cias_cias     
                        JOIN sc_tctco 
                        ON sc_tmvco.mvco_fecmov = sc_tctco.ctco_fecmov 
                        AND sc_tmvco.mvco_tpco = sc_tctco.ctco_tpco 
                        AND sc_tmvco.mvco_nrocom = sc_tctco.ctco_nrocom
                        AND ge_tcias.cias_cias = sc_tctco.CTCO_CIAS
                    WHERE
                        sc_tmvco.mvco_mayo = :1
                        AND sc_tmvco.mvco_fecmov = :2
                        AND mvco_cias IN ({in_clause_placeholders})
                        
            ) where TIPO IN ('AP','RE')

                GROUP BY
                NEGOCIO, TIPO)
                PIVOT (
                    MAX(ABS(VALOR))
                    FOR TIPO
                    IN ('AP' AS APORTES, 'RE' AS RETIROS))

    '''
        
        params = [cuenta, periodo] + codigos
 
        
        try:
            con_params = oracledb.ConnectParams(host = config_sifi['host'], port = config_sifi['port'], service_name = config_sifi['service_name'])

            with oracledb.connect(user = config_sifi['user'],  password = config_sifi['pwd'], params = con_params) as connection:
                df1 = pd.read_sql(sql, connection, params = params)
                df2 = pd.read_sql(sql2, connection, params = params)
                df = pd.merge(df1, df2, on='CÓDIGO', how='left')    
            return df
        except oracledb.DatabaseError as e:
            error, = e.args
            messagebox.showerror("Error", f"Error de base de datos: {error.message}")
            print(e)
            return None
        except Exception as e:
            messagebox.showerror("Error", f"Error al conectar a la base de datos: {e}")
            print(e)
            return None

    def leer_datos_extractos(self):
        df_extractos = pd.DataFrame(columns=['CÓDIGO', 'PERIODO', 'APORTES', 'RETIROS', 'RETENCIÓN EN LA FUENTE', 'RENDIMIENTOS NETOS', 'SALDO EXTRACTO'])
        lista_codigos = [str(codigo).strip() for codigo in self.codigos]
        carpeta = Path(self.entrada_carpeta.get())

        if not carpeta.is_dir():
            messagebox.showerror("Error", f"La ruta '{self.entrada_carpeta.get()}' no existe o no es un directorio.")
            return None
        else:
            try:
                for item in carpeta.iterdir():
                    if item.is_file() and item.suffix.lower() == ".pdf":                
                        for codigo in lista_codigos:
                            if codigo in item.stem:
                                nit = self.consultar_nit(int(codigo))
                                pdf = fitz.Document(item)
                                if pdf.needs_pass:
                                    pdf.authenticate(str(nit))
                                
                                text = pdf[-1].get_text("text")
                                if len(text.strip()) >= 100:
                                    print(f"Archivo: {item.name} - Código: {codigo} - Naturaleza: Texto")
                                    datos = self.extraer_datos_pdf_texto(text, pdf)
                                    datos = [codigo, *datos]
                                else:
                                    print(f"Archivo: {item.name} - Código: {codigo} - Naturaleza: Escaneado")
                                    datos = self.extraer_datos_pdf_escaneado(pdf, 2000)
                                    datos = [codigo, *datos]
                                df_extractos.loc[len(df_extractos)] = datos
                return df_extractos
            except PermissionError:
                messagebox.showerror("Error", f"No tienes permisos para leer la carpeta '{self.entrada_carpeta.get()}'.")
                return None
            except Exception as e:
                messagebox.showerror("Error", f"Ocurrió un error inesperado: {e}")
                return None

    def extraer_datos_pdf_texto(self, texto, pdf):
        if 'Detalle Movimientos' in texto:
            return self.leer_pdf_modelo3(pdf[0], pdf[-1])      
        elif 'bancolombia.com' in texto:
            return self.leer_pdf_modelo1(pdf[-1])
        else:
            return self.leer_pdf_modelo2(pdf[-1])

    def leer_pdf_modelo1(self, pagina):
        datos = []

        lineas = get_text_lines(pagina).split('\n')
        search_list = {'ADICIONES': 2, 'RETIROS': 3, 'REND. NETOS': 0, 'RETENCIÓN': 1, 'NUEVO SALDO': 2}
        for x, linea in enumerate(lineas):

            if 'DESDE' in linea.upper() and 'HASTA' in linea.upper():
                datos.append(re.search(r'\d\d\d\d\d\d', linea)[0])
            else:
                for w in search_list.keys():
                    if w in linea:
                        try:
                            num = self.convert_to_float(lineas[x+4].split('\t')[search_list[w]])
                            search_list[w] = num
                        except:
                            search_list[w] = 'ERROR'
        
        datos.append(search_list['ADICIONES'])
        datos.append(search_list['RETIROS'])
        datos.append(search_list['RETENCIÓN'])
        datos.append(search_list['REND. NETOS'])
        datos.append(search_list['NUEVO SALDO'])

        return datos

    def leer_pdf_modelo2(self, pagina):
        periodo = 'ERROR'
        aporte = 'ERROR'
        retiro = 'ERROR'
        rend = 'ERROR'
        retefuente = 'ERROR'
        saldo_final = 'ERROR'

        lineas = get_text_lines(pagina).split('\n')
        for x, linea in enumerate(lineas):
            if 'PERIODO' in linea.upper() and periodo == 'ERROR':
                match = re.search(r"\d\d-(.*?)-\d\d\d\d", linea)
                if match:
                    periodo = self.process_period(match[0])
            elif 'APORTES' in linea.upper() and aporte == 'ERROR':
                aporte = self.process_numeric_line(linea)

            elif 'RETIROS' in linea.upper() and retiro == 'ERROR':
                retiro = self.process_numeric_line(linea)

            elif 'RETENCIÓN EN LA FUENTE' in linea.upper() and retefuente == 'ERROR':
                retefuente = self.process_numeric_line(linea)

            elif 'RENDIMIENTOS NETOS' in linea.upper() and rend == 'ERROR':
                rend = self.process_numeric_line(linea)

            elif 'SALDO FINAL' in linea.upper() and saldo_final == 'ERROR':
                saldo_final = self.process_numeric_line(linea)

               
        return [periodo, aporte, retiro, retefuente, rend, saldo_final]

    def leer_pdf_modelo3(self, pagina_inicial, pagina_final):
        periodo = 'ERROR'
        aporte = 'ERROR'
        retiro = 'ERROR'
        rend = 'ERROR'
        retefuente = 'ERROR'
        saldo_final = 'ERROR'

        try:
            tabs = pagina_final.find_tables()
            if tabs.tables:
                date = tabs[1].extract()[-2][1].split('/')
                periodo = str(date[2] + date[1])
                aporte = self.process_numeric_line(tabs[1].extract()[-1][3])
                retiro = self.process_numeric_line(tabs[1].extract()[-1][4])
                retefuente = self.process_numeric_line(tabs[1].extract()[-1][6])
                saldo_final = self.process_numeric_line(tabs[1].extract()[-1][8])
            
            tabs = pagina_inicial.find_tables()
            if tabs.tables:
                rend = self.process_numeric_line(tabs[0].extract()[-1][5])
        except:
            ['ERROR', 'ERROR', 'ERROR', 'ERROR', 'ERROR' 'ERROR']

        return [periodo, aporte, retiro, retefuente, rend, saldo_final]
    
    def extraer_datos_pdf_escaneado(self, pdf, dpi=300, idiomas=['es'], usar_gpu=False):
        try:
            reader = easyocr.Reader(idiomas, gpu=usar_gpu)

            zoom = dpi / 72  # El DPI base de PDF suele ser 72
            matriz = fitz.Matrix(zoom, zoom)
            pixmap = pdf[-1].get_pixmap(matrix=matriz, alpha=False) 
            img_bytes = pixmap.tobytes("png")

            resultados = reader.readtext(img_bytes, detail=0, paragraph=False)

            if resultados:
                datos = self.leer_pdf_ocr(resultados)
                return datos
            else:
                print(f"No se detectó texto en la página.")

        except Exception as e:
            print(f"Error al procesar el pdf")
            return None
        finally:
            pdf.close()

    def leer_pdf_ocr(self, resultado):
        periodo = 'ERROR'
        aporte = 'ERROR'
        retiro = 'ERROR'
        rend = 'ERROR'
        retefuente = 'ERROR'
        saldo_final = 'ERROR'

        for x, line in enumerate(resultado):
            if 'PERIODO' in line.upper() and periodo == 'ERROR':
                periodo = resultado[x+1]
                periodo = self.process_period(periodo)
            elif 'APORTES' in line.upper() and aporte == 'ERROR':
                aporte = resultado[x+1].split(' ')[-1]
                aporte = self.process_numeric_line_ocr(aporte)

            elif 'RETIROS' in line.upper() and retiro == 'ERROR':
                retiro = resultado[x+1].split(' ')[-1]
                retiro = self.process_numeric_line_ocr(retiro)

            elif retefuente == 'ERROR' and any(phrase in line.upper() for phrase in ["RETENCIÓN EN LA FUENTE", "RETENCION EN LA FUENTE", "RELENCION EN LA FUENTE"]):
                retefuente = resultado[x+1].split(' ')[-1]
                retefuente = self.process_numeric_line_ocr(retefuente)

            elif rend == 'ERROR' and any(phrase in line.upper() for phrase in ["RENDIMIENTOS NETOS", "RENDIMIENLOS NETOS"]):
                rend = resultado[x+1].split(' ')[-1]
                rend = self.process_numeric_line_ocr(rend)

            elif 'SALDO FINAL' in line.upper() and saldo_final == 'ERROR':
                saldo_final = resultado[x+1].split(' ')[-1]
                saldo_final = self.process_numeric_line_ocr(saldo_final)
               
        return [periodo, aporte, retiro, retefuente, rend, saldo_final]

    def process_numeric_line(self, linea):
        try:
            rex = r"(\d+.*\d+)"
            match = re.search(rex, linea)
            if match:
                return self.convert_to_float(match[0].replace(' ',''))
        except:
            return 'ERROR'

    def process_numeric_line_ocr(self, line):
        try:
            line = line.replace("o", "0").replace("O", "0").replace("l", "1").replace("L", "1")
            line = line.replace("I", "1").replace("i", "1").replace("S", "5").replace("s", "5")
            rex = r"(\d+.*\d+)"
            match = re.search(rex, line)
            if match:
                return self.convert_to_float(match[0].replace(' ',''))
        except:
            return 'ERROR'

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

    def convert_to_float(self, num):
        if num[-3] == '.' or num[-3] == ',':
            num = num[:-3].replace('.', '').replace(',','') + '.' + num[-2:]
        else:
            num = num.replace('.', '').replace(',','')
        return float(num)

root = tk.Tk()
app = BotConciliacion(root)
root.mainloop()