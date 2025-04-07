import oracledb
import pandas as pd

class DataManager(object):
    def __init__(self):
            self.config_sifi = {
                'user'           : "SIFI_RPT",
                'pwd'            : 'C00m3v4123*',
                'host'           : 'BDCDPORA11.INTRACOOMEVA.COM.CO',
                'port'           :  1554,
                'service_name'   : 'CDPORA11'
                }
            

    
    def _extract_data_sifi(self, sql, params):
        con_params = oracledb.ConnectParams(host = self.config_sifi['host'], port = self.config_sifi['port'], service_name = self.config_sifi['service_name'])

        with oracledb.connect(user = self.config_sifi['user'],  password = self.config_sifi['pwd'], params = con_params) as connection:
            df = pd.read_sql(sql, connection, params = params)
            connection.close()
        return df
    
    def get_all_movements(self, date_from, date_to, cias_list):


        sql = '''
        select * from(
   SELECT
            mvco_te_tpmv TIPO_BD,
            ge_tcias.cias_cias NEGOCIO,
            ge_tcias.cias_descri NOMBRE,
            sc_tmvco.mvco_mtoren VALOR,
            sc_tmvco.mvco_tporan NATURALEZA,
            sc_tmvco.mvco_fecmov PERIODO,
            sc_tmvco.mvco_mayo CUENTA,
            ge_tauxil.auxi_nit NIT,
            CASE
                WHEN mvco_descri LIKE '% 241' THEN 'AP'
                WHEN mvco_descri LIKE '% 240' THEN 'AP'
                WHEN mvco_descri LIKE '%ORDEN DE INVERSIÓN %' THEN 'AP'
                WHEN mvco_descri LIKE '% 242' THEN 'RE'
                WHEN mvco_te_tpmv = 34 THEN 'RE'
            ELSE 'NA'
            END AS TIPO,
            mvco_descri DESCRI,
            sc_tmvco.mvco_fecdoc FECHA_DOCUMENTO,
            sc_tmvco.mvco_etct ESTRUCTURA,
            sc_tmvco.mvco_feccre FECHA_CREACION,
            sc_tmvco.mvco_usua USUARIO,
            sc_tmvco.mvco_fcmod FECHA_MODIFICACION
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
            sc_tmvco.mvco_cias = :cias
            AND sc_tmvco.mvco_fecmov between :periodo1 and :periodo2
            AND sc_tmvco.mvco_mayo between 130205010105 and 130205010105
           
        ORDER BY
            ge_tcias.cias_cias
) where TIPO IN ('AP','RE')
        '''
   
        final_df = pd.DataFrame()
        for c in cias_list:
            params = [str(c), date_from, date_to]
            df = self._extract_data_sifi(sql, params)
            final_df = pd.concat([final_df, df] , ignore_index=True)

        self.all_movements = final_df

    def get_cias_info(self, cias_list):
        sql = '''
            Select cias_cias CODIGO, cias_descri NOMBRE, cias_nrorif NIT from ge_tcias where cias_cias = :cias       
            '''
        final_df = pd.DataFrame()
        for c in cias_list:
            params = [str(c)]
            df = self._extract_data_sifi(sql, params)
            final_df = pd.concat([final_df, df] , ignore_index=True)

        return final_df.to_dict()