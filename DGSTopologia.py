# Ingnore warnings
import warnings
warnings.filterwarnings('ignore')
warnings.simplefilter('ignore')

# Import libraries
import pandas as pd
from pathlib import Path
import datetime as dt
import numpy as np


# Leer archivo de excel
s_filepath=r'C:\AnalisisElectrico\Scripts_Py_Dig\Salida_Dig.xlsx'
df_data=pd.read_excel(s_filepath, sheet_name='Data')

s_Fecha=df_data.loc[0,'Valor']

sMes=df_data.loc[1,'Valor']

Version=df_data.loc[2,'Valor']

iPini=df_data.loc[3,'Valor']
iPfin=df_data.loc[4,'Valor']

ano=s_Fecha.year
mes=s_Fecha.month
dia=s_Fecha.day

# Get main path and other folders
s_mainpath=Path.cwd()
s_parentpath=s_mainpath.parent
s_folderpath='AnalisisMensual'
s_file='Mmtosdia_' + Version + '.xlsx'
s_filepath=s_parentpath.joinpath(s_folderpath,str(ano) + '-' + "{:02d}".format(sMes),Version,s_file)

#print(s_filepath)

#Read file
s_Sheet=str(mes) + '_' + str(dia) + 'A'

df_data=pd.read_excel(s_filepath, sheet_name=s_Sheet)

def EncabezadoDGS(FileName):
    file = open(FileName,'w')
    file.write('$$General;ID(a:40);Descr(a:40);Val(a:40)\n'\
                '1;Version;5.0\n')
    file.close()

def ElementosElmSym(FileName):
    file = open(FileName,'a+')
    file.write('$$ElmSym;ID(a:40);outserv(i)\n')
    file.close()

def ElementosElmGenstat(FileName):
    file = open(FileName,'a+')
    file.write('$$ElmGenstat;ID(a:40);outserv(i)\n')
    file.close()

def ElementosElmCoup(FileName):
    file = open(FileName,'a+')
    file.write('$$ElmCoup;ID(a:40);on_off(i)\n')
    file.close() 

def ElementosStaSwitch(FileName):
    file = open(FileName,'a+')
    file.write('$$StaSwitch;ID(a:40);on_off(i)\n')
    file.close() 

def ElementosElmTr2(FileName):
    file = open(FileName,'a+')
    file.write('$$ElmTr2;ID(a:40);outserv(i)\n')
    file.close() 

def ElementosElmTr3(FileName):
    file = open(FileName,'a+')
    file.write('$$ElmTr3;ID(a:40);outserv(i)\n')
    file.close() 

def ElementosElmLne(FileName):
    file = open(FileName,'a+')
    file.write('$$ElmLne;ID(a:40);outserv(i)\n')
    file.close() 

def ElementosElmShnt(FileName):
    file = open(FileName,'a+')
    file.write('$$ElmShnt;ID(a:40);outserv(i)\n')
    file.close() 

def ElementosElmTerm(FileName):
    file = open(FileName,'a+')
    file.write('$$ElmTerm;ID(a:40);outserv(i)\n')
    file.close() 


# Funciones para encontrar los periodos iniciales y finales de cada día
def New_ColumnDgs(Fkey,Estado):
    # Set hour at ini
    if Estado=='on_off':
        dgsvalue='##' + str(Fkey) + ';' + '0'
    else:
        dgsvalue='##' + str(Fkey) + ';' + '1'
    
    return dgsvalue

def WriteElm(FileName,df_data, Tipo):
    df_Elm=df_data[(df_data.Pini<=t) & (df_data.Pfin>=t) & (df_data.TipoElm==Tipo)]
    df_Elm['DGS']=df_Elm.apply(lambda row: New_ColumnDgs(row['Fkey'],row['Estado']),axis=1)
    df_Elm.drop_duplicates(subset='DGS',inplace=True)
    outPutFile=df_Elm['DGS']
    outPutFile.to_csv(FileName, index=False, header=None, mode='a+') 


# Filtrar por los elementos que se deben incluir

#Read file
s_Sheet=str(mes) + '_' + str(dia) + 'A'

df_data=pd.read_excel(s_filepath, sheet_name=s_Sheet)

df_data=df_data[df_data.Incluir==1]


# Incluir la apertura de las barras a través de los elementos
# Mapeo SIO Digsilent
s_mainpath=Path.cwd()
s_parentpath=s_mainpath.parent
s_file='Mapeo_SIO_Dig.xlsx'
s_filepath=s_parentpath.joinpath(s_folderpath,s_file)
df_Dig_Bus=pd.read_excel(s_filepath, sheet_name='Barras')

df_Bus=df_data[df_data.TipoElm=='ElmTerm']
colname=df_Bus.columns
df_Bus.drop(['Fkey','Estado','TipoElm'],axis=1,inplace=True)
df_findElm=pd.merge(df_Bus,df_Dig_Bus, left_on=['Elemento'],right_on=['Barra_SIO'], how='inner')[colname]
df_data=pd.concat([df_data,df_findElm])

# Incluir las ramas
df_Dig_Br=pd.read_excel(s_filepath, sheet_name='Branch')
df_Br=df_data[df_data.TipoElm=='ElmLne']
colname=df_Br.columns
df_Br.drop(['Fkey','Estado','TipoElm'],axis=1,inplace=True)
df_findElm=pd.merge(df_Br,df_Dig_Br, left_on=['Elemento'],right_on=['Elm_SIO'], how='inner')[colname]
df_data=pd.concat([df_data,df_findElm])

# Cargar información de topología según la fecha
df_Dig_Top=pd.read_excel(s_filepath, sheet_name='Topologia')

for ind in df_Dig_Top.index:

    if (df_Dig_Top.loc[ind,'FechaIni']<=s_Fecha) and (df_Dig_Top.loc[ind,'FechaFin']>=s_Fecha) and (df_Dig_Top.loc[ind,'Incluir']==1):
        data=pd.DataFrame(df_Dig_Top.loc[ind]).T
        data.drop(['FechaIni','FechaFin'],axis=1,inplace=True)
        data['Consecutivo']='Topología'
        data.rename(columns={'Elemento_Dig':'Elemento'},inplace=True)
        data['TipoAfe']='D'
        data['Detalle']='-'
        col_order=['Consecutivo','Elemento','TipoAfe','Detalle','Pini','Pfin','Fkey','Estado','TipoElm','Incluir']
        data=data[col_order]
        df_data=pd.concat([df_data,data])


l_periodos=[i for i in range(iPini,iPfin+1)]

RutaDGS=r"C:\AnalisisElectrico\DGS_Topologia"

for t in l_periodos:

    FileName = RutaDGS + "\P{:02}".format(int(t)) + '.dgs'
    EncabezadoDGS(FileName)
    
    ElementosElmSym(FileName)
    WriteElm(FileName,df_data, 'ElmSym')

    ElementosElmGenstat(FileName)
    WriteElm(FileName,df_data, 'ElmGenstat')

    ElementosElmCoup(FileName)
    WriteElm(FileName,df_data, 'ElmCoup')

    ElementosStaSwitch(FileName)
    WriteElm(FileName,df_data, 'StaSwitch')

    ElementosElmTr2(FileName)
    WriteElm(FileName,df_data, 'ElmTr2')

    ElementosElmTr3(FileName)
    WriteElm(FileName,df_data, 'ElmTr3')

    ElementosElmLne(FileName)
    WriteElm(FileName,df_data, 'ElmLne')

    ElementosElmShnt(FileName)
    WriteElm(FileName,df_data, 'ElmShnt')

    ElementosElmTerm(FileName)
    WriteElm(FileName,df_data, 'ElmTerm')