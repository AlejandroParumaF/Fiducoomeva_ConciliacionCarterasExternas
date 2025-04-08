import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import oracledb
from pathlib import Path
import pymupdf as fitz
import re
from pymupdf4llm.helpers.get_text_lines import get_text_lines

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
            self.leer_datos_extractos()

            print(self.df_retfte)
            print(self.df_database)

    def consultar_nit(self, codigo):
        nit = self.df_database.loc[self.df_database['CÓDIGO'] == codigo, 'NIT PA']
        if not nit.empty:
            return nit.values[0]
        else:
            messagebox.showerror("Error", f"NIT no encontrado para el código {codigo}")
            return None

    def leer_datos_rtef(self):
        df_retfte = pd.read_excel(self.entrada_archivo.get())

        try:
            df_retfte = df_retfte[['Codigo PA', 'TOTAL']]
            df_retfte = df_retfte.rename(columns={'Codigo PA': 'CÓDIGO', 'TOTAL': 'RETENCIONES NO PROCEDENTES'})
            return df_retfte
        except KeyError:
            messagebox.showerror("Error", "El archivo Datos Dev Retefte no contiene las columnas Código PA y/o TOTAL")
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
        params = [cuenta, periodo] + codigos
 
        try:
            con_params = oracledb.ConnectParams(host = config_sifi['host'], port = config_sifi['port'], service_name = config_sifi['service_name'])

            with oracledb.connect(user = config_sifi['user'],  password = config_sifi['pwd'], params = con_params) as connection:
                df = pd.read_sql(sql, connection, params = params)     
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
                                    print(datos)
                                else:
                                    print(f"Archivo: {item.name} - Código: {codigo} - Naturaleza: Escaneado")
                                    self.extraer_datos_pdf_escaneado(codigo, text)
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
    
    def extraer_datos_pdf_escaneado(self, codigo, texto):
        print(f"Extrayendo datos del PDF escaneado para el código {codigo}")

    def process_numeric_line(self, linea):
        try:
            rex = r"(\d+.*\d+)"
            match = re.search(rex, linea)
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