import pandas as pd
from datetime import datetime
import win32com.client as win32
import xlsxwriter
import os



#Dependencies: Engineer name in sharepoints (i.e. everything is assigned), ai and str groups
#Revise M2 code if engineer names changes
#To do: Open each file for update or check if automatic when accesing via python

today = datetime.today().strftime('%d-%b-%y')
today_week = datetime.today().isocalendar()[1]
opp_placards = 0

#Setting Sharpeoint excel lists paths
path_sp = r'C:/Users/ggalina/Engineering Metrics/Sharepoint Data.xlsx'

#Refreshing Data Sources (PON TODO EN UN EXCEL EN VARIOS SHEETS, MEJOR ASI!)
# xlapp = win32com.client.DispatchEx("Excel.Application")
# wb = xlapp.Workbooks.Open(path_sdocs)
# wb.RefreshAll()
# xlapp.CalculateUntilAsyncQueriesDone()
# wb.Save()
# xlapp.Quit()

#Setting up dataframes
df_pma = pd.read_excel(path_sp,sheet_name= 'PMA') #Dataframe for PMAs
df_sdocs = pd.read_excel(path_sp,sheet_name= 'Source Docs') #Dataframe for open source docs
df_er = pd.read_excel(path_sp,sheet_name= 'Engineering Requests') #Dataframe for engineering requests
df_alerts = pd.read_excel(path_sp,sheet_name= 'Alerts') #Dataframe for alerts
df_aog = pd.read_excel(path_sp,sheet_name= 'AOG') #Dataframe for AOGs
df_ual = pd.read_excel(path_sp,sheet_name= 'UAL Docs') #Dataframe for UAL Build Standards
df_edocs = pd.read_excel(path_sp,sheet_name= 'EDocs') #Dataframe for UAL Build Standards
df_cck = pd.read_excel(path_sp,sheet_name= 'CCK Data') #Dataframe for CCK


#Setting up AI and STR groups
ai_group = ['Carlos Martin Gonzalez (CM)','DIANNETTE ALVARADO LIZONDRO (CM)' ,'GIANCARLO GALINA QUINTERO (CM)', 'KARLA DEL CID QUINTANA (CM)','MONIQUE PEREGRINA VILLALAZ (CM)','JOSE APARICIO CASTILLO (CM)','ANGEL VALDES MARIN (CM)', 'JUAN BEROY  (CM)']
str_group = ['ARTURO SUCRE MELFI (CM)','JOSE VALDES CABALLERO (CM)','OSVALDO GALLARDO O (CM)','JOSE HENRY ICAZA (CM)','CESAR BARROSO RODRIGUEZ (CM)','JUAN DE LAS CASAS CORDERO (CM)']

#PMA Info
filt_total = (df_pma['Status'] == 'PENDING EVALUATION') | (df_pma['Status'] == 'CHECKLIST ON SIGNATURE')
filt_str = ((df_pma['Status'] == 'PENDING EVALUATION') | (df_pma['Status'] == 'CHECKLIST ON SIGNATURE')) & (df_pma['Engineer'].isin(str_group))
filt_ai = ((df_pma['Status'] == 'PENDING EVALUATION') | (df_pma['Status'] == 'CHECKLIST ON SIGNATURE')) & (df_pma['Engineer'].isin(ai_group))
pma_open = (df_pma.loc[filt_total].shape)[0]
pma_str = (df_pma.loc[filt_str].shape)[0]
pma_ai = (df_pma.loc[filt_ai].shape)[0]

#Open Source Docs Info
#Embraer source docs info
filt_emb_str = (df_sdocs['FLEET / TYPE'] == 'EMB') & (df_sdocs['EVALUATION'].isin(['PENDING','PENDING EA'])) & (df_sdocs['ASSIGNED TO'].isin(str_group))
filt_emb_ai = (df_sdocs['FLEET / TYPE'] == 'EMB') & (df_sdocs['EVALUATION'].isin(['PENDING','PENDING EA'])) & (df_sdocs['ASSIGNED TO'].isin(ai_group))
sdocs_emb_str = (df_sdocs[filt_emb_str].shape)[0]
sdocs_emb_ai = (df_sdocs[filt_emb_ai].shape)[0]

#737NG source doc info
filt_ng_str = (df_sdocs['FLEET / TYPE'] == 'NG') & (df_sdocs['EVALUATION'].isin(['PENDING','PENDING EA'])) & (df_sdocs['ASSIGNED TO'].isin(str_group))
filt_ng_ai = (df_sdocs['FLEET / TYPE'] == 'NG') & (df_sdocs['EVALUATION'].isin(['PENDING','PENDING EA'])) & (df_sdocs['ASSIGNED TO'].isin(ai_group))
sdocs_ng_str = (df_sdocs[filt_ng_str].shape)[0]
sdocs_ng_ai = (df_sdocs[filt_ng_ai].shape)[0]

#737MAX source doc info
filt_max_str = (df_sdocs['FLEET / TYPE'] == 'MAX') & (df_sdocs['EVALUATION'].isin(['PENDING','PENDING EA'])) & (df_sdocs['ASSIGNED TO'].isin(str_group))
filt_max_ai = (df_sdocs['FLEET / TYPE'] == 'MAX') & (df_sdocs['EVALUATION'].isin(['PENDING','PENDING EA'])) & (df_sdocs['ASSIGNED TO'].isin(ai_group))
sdocs_max_str = (df_sdocs[filt_max_str].shape)[0]
sdocs_max_ai = (df_sdocs[filt_max_ai].shape)[0]

#737SL source doc info
filt_sl_str = (df_sdocs['FLEET / TYPE'] == 'SL737') & (df_sdocs['EVALUATION'].isin(['PENDING','PENDING EA'])) & (df_sdocs['ASSIGNED TO'].isin(str_group))
filt_sl_ai = (df_sdocs['FLEET / TYPE'] == 'SL737') & (df_sdocs['EVALUATION'].isin(['PENDING','PENDING EA'])) & (df_sdocs['ASSIGNED TO'].isin(ai_group))
sdocs_sl_str = (df_sdocs[filt_sl_str].shape)[0]
sdocs_sl_ai = (df_sdocs[filt_sl_ai].shape)[0]

#Components source doc info
filt_comp_str = (df_sdocs['FLEET / TYPE'] == 'COMPONENTS SB/SL') & (df_sdocs['EVALUATION'].isin(['PENDING','PENDING EA'])) & (df_sdocs['ASSIGNED TO'].isin(str_group))
filt_comp_ai = (df_sdocs['FLEET / TYPE'] == 'COMPONENTS SB/SL') & (df_sdocs['EVALUATION'].isin(['PENDING','PENDING EA'])) & (df_sdocs['ASSIGNED TO'].isin(ai_group))
sdocs_comp_str = (df_sdocs[filt_comp_str].shape)[0]
sdocs_comp_ai = (df_sdocs[filt_comp_ai].shape)[0]

#Others (SAIB & AMOCS) source doc info
filt_others_str = (df_sdocs['FLEET / TYPE'].isin(['AMOC 737','SAIB'])) & (df_sdocs['EVALUATION'].isin(['PENDING','PENDING EA'])) & (df_sdocs['ASSIGNED TO'].isin(str_group))
filt_others_ai = (df_sdocs['FLEET / TYPE'].isin(['AMOC 737','SAIB'])) & (df_sdocs['EVALUATION'].isin(['PENDING','PENDING EA'])) & (df_sdocs['ASSIGNED TO'].isin(ai_group))
sdocs_others_str = (df_sdocs[filt_others_str].shape)[0]
sdocs_others_ai = (df_sdocs[filt_others_ai].shape)[0]

#Engineering requests info. Nota: Hay que reasignar casos y limpiar sp, hay asignados a brujos
filt_er_str = (df_er['Status (Sólo ING)'] == 'Abierto') & (df_er['Asignado A'].isin(str_group))
filt_er_ai = (df_er['Status (Sólo ING)'] == 'Abierto') & (df_er['Asignado A'].isin(ai_group))
er_str = (df_er[filt_er_str].shape)[0]
er_ai = (df_er[filt_er_ai].shape)[0]

#Alerts info
filt_alerts_str = (df_alerts['STATUS'] == 'OPEN') & (df_alerts['ASSIGNED TO'].isin(str_group))
filt_alerts_ai = (df_alerts['STATUS'] == 'OPEN') & (df_alerts['ASSIGNED TO'].isin(ai_group))
alerts_str = (df_alerts[filt_alerts_str].shape)[0]
alerts_ai = (df_alerts[filt_alerts_ai].shape)[0]

#C-Check info
filt_eng = (df_cck['Aircraft'].str.contains('HP')) & (df_cck['Time Initial Request'].dt.year ==int(datetime.today().strftime('%Y')))
df_cck.dropna(axis = 0,how = 'all', subset = ['Aircraft'], inplace= True)
c_ck_cases = df_cck[filt_eng].shape[0] #contar todos los casos
c_ck_time = round(df_cck[filt_eng]['TiempoEng'].mean(),1)

#AOG info
df_aog['MOC TO ENG'] = pd.to_datetime(df_aog['MOC TO ENG'])
filt_aog = (df_aog['MOC TO ENG'].dt.year) == int(datetime.today().strftime('%Y'))
aog_cases = (df_aog[filt_aog].shape)[0]
aog_time = round(df_aog.loc[filt_aog,'ENG TIME'].mean(),2)

#UAL Build Standards info
ual_ecra_str = df_ual['COPA ECRA'].isna().sum()
ual_ecra_ai = 0

#Getting E-Docs info
filt_str = (df_edocs['Task Code (JIC)'].str.startswith('EA 5')) |(df_edocs['Task Code (JIC)'].str.startswith('EA 7'))|(df_edocs['Task Code (JIC)'].str.startswith('FCD 5'))
df_edocs_str = df_edocs[filt_str]
df_edocs_ai = df_edocs[~filt_str]
edocs_str = df_edocs_str.isnull().any(axis=1).sum()
edocs_ai = df_edocs_ai.isnull().any(axis=1).sum()
#df_edocs.isnull().any(axis=1).sum()

#Creating M2 content. Verify engineers name if they change
filt_pend_as = (df_sdocs['EVALUATION'].isin(['PENDING','PENDING EA'])) & (df_sdocs['ASSIGNED TO']=='ARTURO SUCRE MELFI (CM)')
filt_pend_cb = (df_sdocs['EVALUATION'].isin(['PENDING','PENDING EA'])) & (df_sdocs['ASSIGNED TO']=='CESAR BARROSO RODRIGUEZ (CM)')
filt_pend_da = (df_sdocs['EVALUATION'].isin(['PENDING','PENDING EA'])) & (df_sdocs['ASSIGNED TO']=='DIANNETTE ALVARADO LIZONDRO (CM)')
filt_pend_gg = (df_sdocs['EVALUATION'].isin(['PENDING','PENDING EA'])) & (df_sdocs['ASSIGNED TO']=='GIANCARLO GALINA QUINTERO (CM)')
# filt_pend_cc = (df_sdocs['EVALUATION'].isin(['PENDING','PENDING EA'])) & (df_sdocs['ASSIGNED TO']=='GIANCARLO GALINA QUINTERO (CM)')
filt_pend_jh = (df_sdocs['EVALUATION'].isin(['PENDING','PENDING EA'])) & (df_sdocs['ASSIGNED TO']=='JOSE HENRY ICAZA (CM)')
filt_pend_jv = (df_sdocs['EVALUATION'].isin(['PENDING','PENDING EA'])) & (df_sdocs['ASSIGNED TO']=='JOSE VALDES CABALLERO (CM)')
filt_pend_kdc = (df_sdocs['EVALUATION'].isin(['PENDING','PENDING EA'])) & (df_sdocs['ASSIGNED TO']=='KARLA DEL CID QUINTANA (CM)')
filt_pend_mp = (df_sdocs['EVALUATION'].isin(['PENDING','PENDING EA'])) & (df_sdocs['ASSIGNED TO']=='MONIQUE PEREGRINA VILLALAZ (CM)')

pend_as = df_sdocs.loc[filt_pend_as,'FLEET / TYPE'].value_counts().to_dict()
pend_cb = df_sdocs.loc[filt_pend_cb,'FLEET / TYPE'].value_counts().to_dict()
pend_da = df_sdocs.loc[filt_pend_da,'FLEET / TYPE'].value_counts().to_dict()
pend_gg = df_sdocs.loc[filt_pend_gg,'FLEET / TYPE'].value_counts().to_dict()
pend_jh = df_sdocs.loc[filt_pend_jh,'FLEET / TYPE'].value_counts().to_dict()
pend_jv = df_sdocs.loc[filt_pend_jv,'FLEET / TYPE'].value_counts().to_dict()
pend_kdc = df_sdocs.loc[filt_pend_kdc,'FLEET / TYPE'].value_counts().to_dict()
pend_mp = df_sdocs.loc[filt_pend_mp,'FLEET / TYPE'].value_counts().to_dict()

#Arturo Sucre Variables
try:
    as_amoc = pend_as['AMOC 737']
except:
    as_amoc = 0 
try:
    as_emb = pend_as['EMB']
except:
    as_emb = 0
try:
    as_max = pend_as['MAX']
except:
    as_max = 0
try:
    as_ng = pend_as['NG']
except:
    as_ng = 0
try:
    as_saib = pend_as['SAIB']
except:
    as_saib = 0
try:
    as_sl = pend_as['SL737']
except:
    as_sl = 0
try:
    as_comp = pend_as['COMPONENTS SB/SL']
except:
    as_comp = 0


#Cesar Barroso Variables
try:
    cb_amoc = pend_cb['AMOC 737']
except:
    cb_amoc = 0 
try:
    cb_emb = pend_cb['EMB']
except:
    cb_emb = 0
try:
    cb_max = pend_cb['MAX']
except:
    cb_max = 0
try:
    cb_ng = pend_cb['NG']
except:
    cb_ng = 0
try:
    cb_saib = pend_cb['SAIB']
except:
    cb_saib = 0
try:
    cb_sl = pend_cb['SL737']
except:
    cb_sl = 0
try:
    cb_comp = pend_cb['COMPONENTS SB/SL']
except:
    cb_comp = 0


#Diannette Alvarado Variables
try:
    da_amoc = pend_da['AMOC 737']
except:
    da_amoc = 0 
try:
    da_emb = pend_da['EMB']
except:
    da_emb = 0
try:
    da_max = pend_ad['MAX']
except:
    da_max = 0
try:
    da_ng = pend_da['NG']
except:
    da_ng = 0
try:
    da_saib = pend_da['SAIB']
except:
    da_saib = 0
try:
    da_sl = pend_da['SL737']
except:
    da_sl = 0
try:
    da_comp = pend_da['COMPONENTS SB/SL']
except:
    da_comp = 0


#Giancarlo Variabes
try:
    gg_amoc = pend_gg['AMOC 737']
except:
    gg_amoc = 0 
try:
    gg_emb = pend_gg['EMB']
except:
    gg_emb = 0
try:
    gg_max = pend_gg['MAX']
except:
    gg_max = 0
try:
    gg_ng = pend_gg['NG']
except:
    gg_ng = 0
try:
    gg_saib = pend_gg['SAIB']
except:
    gg_saib = 0
try:
    gg_sl = pend_gg['SL737']
except:
    gg_sl = 0
try:
     gg_comp = pend_gg['COMPONENTS SB/SL']
except:
    gg_comp = 0
    
    
#Jose Henry Variables
try:
    jh_amoc = pend_jh['AMOC 737']
except:
    jh_amoc = 0 
try:
    jh_emb = pend_jh['EMB']
except:
    jh_emb = 0
try:
    jh_max = pend_jh['MAX']
except:
    jh_max = 0
try:
    jh_ng = pend_jh['NG']
except:
    jh_ng = 0
try:
    jh_saib = pend_jh['SAIB']
except:
    jh_saib = 0
try:
    jh_sl = pend_jh['SL737']
except:
    jh_sl = 0
try:
    jh_comp = pend_jh['COMPONENTS SB/SL']
except:
    jh_comp = 0


#Jose Valdes Variables
try:
    jv_amoc = pend_jv['AMOC 737']
except:
    jv_amoc = 0 
try:
    jv_emb = pend_jv['EMB']
except:
    jv_emb = 0
try:
    jv_max = pend_jv['MAX']
except:
    jv_max = 0
try:
    jv_ng = pend_jv['NG']
except:
    jv_ng = 0
try:
    jv_saib = pend_jv['SAIB']
except:
    jv_saib = 0
try:
    jv_sl = pend_jv['SL737']
except:
    jv_sl = 0
try:
    jv_comp = pend_jv['COMPONENTS SB/SL']
except:
    jv_comp = 0

#Karla Del Cid Variables
try:
    kdc_amoc = pend_kdc['AMOC 737']
except:
    kdc_amoc = 0 
try:
    kdc_emb = pend_kdc['EMB']
except:
    kdc_emb = 0
try:
    kdc_max = pend_kdc['MAX']
except:
    kdc_max = 0
try:
    kdc_ng = pend_kdc['NG']
except:
    kdc_ng = 0
try:
    kdc_saib = pend_kdc['SAIB']
except:
    kdc_saib = 0
try:
    kdc_sl = pend_kdc['SL737']
except:
    kdc_sl = 0
try:
    kdc_comp = pend_kdc['COMPONENTS SB/SL']
except:
    kdc_comp = 0


#Monique Peregrina Variables
try:
    mp_amoc = pend_mp['AMOC 737']
except:
    mp_amoc = 0 
try:
    mp_emb = pend_mp['EMB']
except:
    mp_emb = 0
try:
    mp_max = pend_mp['MAX']
except:
    mp_max = 0
try:
    mp_ng = pend_mp['NG']
except:
    mp_ng = 0
try:
    mp_saib = pend_mp['SAIB']
except:
    mp_saib = 0
try:
    mp_sl = pend_mp['SL737']
except:
    mp_sl = 0
try:
    mp_comp = pend_mp['COMPONENTS SB/SL']
except:
    mp_comp = 0

#Setting row/column values
amoc = [as_amoc, cb_amoc, da_amoc, gg_amoc, 1, jh_amoc, jv_amoc, kdc_amoc, mp_amoc]
emb = [as_emb, cb_emb, da_emb, gg_emb, 1, jh_emb, jv_emb, kdc_emb, mp_emb]
maxl = [as_max, cb_max, da_max, gg_max, 1, jh_max, jv_max, kdc_max, mp_max]
ng = [as_ng, cb_ng, da_max, gg_ng, 1, jh_ng, jv_ng, kdc_ng, mp_ng]
saib = [as_saib, cb_saib, da_saib, gg_saib, 1, jh_saib, jv_saib, kdc_saib, mp_saib]
sl = [as_sl, cb_sl, da_sl, gg_sl, 1, jh_sl, jv_sl, kdc_sl, mp_sl]
comp = [as_comp, cb_comp, da_comp, gg_comp, 1, jh_comp, jv_comp, kdc_comp, mp_comp]

#Calculating totals per Engineer
as_total = as_amoc + as_emb + as_max + as_ng + as_saib + as_sl + as_comp
cb_total = cb_amoc + cb_emb + cb_max + cb_ng + cb_saib + cb_sl + cb_comp
da_total = da_amoc + da_emb + da_max + da_ng + da_saib + da_sl + da_comp
gg_total = gg_amoc + gg_emb + gg_max + gg_ng + gg_saib + gg_sl + gg_comp
cc_total = 0
jh_total = jh_amoc + jh_emb + jh_max + jh_ng + jh_saib + jh_sl + jh_comp
jv_total = jv_amoc + jv_emb + jv_max + jv_ng + jv_saib + jv_sl + jv_comp
kdc_total = kdc_amoc + kdc_emb + kdc_max + kdc_ng + kdc_saib + kdc_sl + kdc_comp
mp_total = mp_amoc + mp_emb + mp_max + mp_ng + mp_saib + mp_sl + mp_comp

total = [as_total, cb_total, da_total, gg_total, cc_total, jh_total, jv_total, kdc_total, mp_total]

#Setting indices for M2 dataframe
m2_index = {0:'Arturo Sucre', 1: 'Cesar Barroso', 2: 'Diannette Alvarado', 3: 'Giancarlo Galina',4:'Carlos Castillo', 5: 'Jose Henry Icaza', 6:'Jose Valdes', 7: 'Karla Del Cid', 8:'Monique Peregrina'}

#Setting up current week series
metrics = [pma_ai, pma_str, sdocs_emb_ai, sdocs_emb_str, sdocs_ng_ai, sdocs_ng_str, sdocs_max_ai, sdocs_max_str, sdocs_others_ai, sdocs_others_str, sdocs_comp_ai, sdocs_comp_str, ual_ecra_ai, ual_ecra_str, er_ai, er_str, alerts_ai, alerts_str, c_ck_cases, c_ck_time, aog_cases, aog_time, opp_placards, edocs_ai, edocs_str]
week_data = {f'Week {today_week}': metrics}
df_week = pd.DataFrame(week_data) #Dataframe for current week metrics


#Leo las metricas iniciales
df_init = pd.read_excel(path_init) #Dataframe for Initial Metrics

#Creo nuevo dataframe con nueva columna de current week data
df_metrics = ((df_init.join(df_week)).set_index('ITEM')).astype(object) #astype(object) te convierte al que se parece el objeto
#Si el join no sirve quita el index_col arriba

#Mando el dataframe al archivo original para añadir la nueva columna
#df_metrics.to_excel(path_init) #Te manda la columna del week a las metricas acumuladas

df_metrics_email = df_metrics[df_metrics.columns[-5:]]


#Ahora con ese nuevo archivo actualizado puedo manipular la data como quiera

#Setting up M2 for html
m2 = {'AMOC 737': amoc, 'EMB': emb, 'MAX': maxl, 'NG':ng, 'SAIB':saib, 'SL737': sl, 'COMP': comp, 'TOTAL':total}
df_m2 = pd.DataFrame.from_dict(m2)
df_m2.rename(index = m2_index, inplace= True)
html_table_m2 = df_m2.to_html(col_space = 70,table_id = 'M2')

# Setting up email send
html_table = df_metrics_email.to_html(col_space = 70,table_id = 'Metrics')
css = '''<style>
table {text-align: center;}
table thead th {text-align: center;}
table, th, td { border: 1px solid black;}
table {border-collapse: collapse}
th, td {padding: 9px}
table {font-family: verdana}
table {font-size: 12}
</style>'''
email_table = css + html_table
outlook = win32.gencache.EnsureDispatch('Outlook.Application')
mail_item = outlook.CreateItem(0)
mail_item.To = 'ggalina@copaair.com'
mail_item.Subject = 'Engineering Metrics - Open Items'
body = 'A continuación data compilada de los pendientes del departamento de ingeniería de ATA por área (AI: Aviónica, Sistemas, Interiores y STR: Estructuras y Motores):'+ '<br>' + '<br>'  + email_table + '<br>' + '<br>' + html_table_m2
mail_item.HTMLBody = (body)
mail_item.Send()




#Source doc sale beby orque usa los de JA, pendiente que se 
#reasigne
#Preguntar cuantas n columns hay que agarrar, cómo cambió la logica?


