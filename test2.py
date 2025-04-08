import oracledb, pandas as pd

config_sifi = {
                'user'           : "SIFI_RPT",
                'pwd'            : 'C00m3v4123*',
                'host'           : 'BDCDPORA11.INTRACOOMEVA.COM.CO',
                'port'           :  1554,
                'service_name'   : 'CDPORA11'
                }

codigos = [119656, 102699, 93235, 102022, 71316, 98080, 116047, 106295, 106919]
cuenta = 130205010105
periodo = '202412'

df = pd.DataFrame()
sql_final = ""

try:
    if not codigos:
        print("La lista 'pa_list' está vacía. No se ejecutará la consulta.")
    else:
        
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
        sql_final = sql        
        params = [cuenta, periodo] + codigos
 
        con_params = oracledb.ConnectParams(host=config_sifi['host'],
                                            port=config_sifi['port'],
                                            service_name=config_sifi['service_name'])

        with oracledb.connect(user=config_sifi['user'], password=config_sifi['pwd'], params=con_params) as connection:
            df = pd.read_sql(sql_final, connection, params=params)


except oracledb.DatabaseError as e:
    error, = e.args
    print("Error de Oracle al ejecutar la consulta:")
    print(f"SQL Intentado:\n{sql_final}")
    # print(f"Parámetros intentados: {params_combined}") # Descomentar si ayuda
    print(f"Código de Error: {error.code}")
    print(f"Mensaje: {error.message}")
    print(f"Contexto: {error.context}")
except Exception as ex:
    print(f"Ocurrió un error inesperado: {ex}")
    print(f"SQL durante el error:\n{sql_final}")

# --- Mostrar Resultados ---
if not df.empty:
    print("\n--- Resultados de la Consulta (DataFrame) ---")
    print(df)
elif not codigos:
     print("\nNo se ejecutó la consulta porque la lista 'codigos' estaba vacía.")
else:
    print("\nLa consulta no devolvió resultados o ocurrió un error.")