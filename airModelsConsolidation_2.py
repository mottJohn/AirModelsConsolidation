from __future__ import division

##############################
# USER INPUTS

# reference
#https://www.epd.gov.hk/epd/english/environmentinhk/air/guide_ref/guide_aqa_model_g5.html
#https://www.epd.gov.hk/epd/english/environmentinhk/air/guide_ref/guide_aqa_model_g1.html
##############################

#leave path empty if no use
path_cmaq = r'C:\Users\CHA82870\Mott MacDonald\AERMOD Modelling Services - Do\06 Working\PATH\cmaq'
path_chimney = r'D:\Hirams Highway\AirModelsConsolidation\mTSPt2\Chimney' #Chimney
path_construction = r'D:\Hirams Highway\AirModelsConsolidation\mTSPt2\Construction_ASRs'#construction
path_caline = r'' #caline
path_marine = r'' #marine
path_all = r'D:\Hirams Highway\AirModelsConsolidation\mTSPt2\All_model_sum'
excel_name = r'D:\Hirams Highway\AirModelsConsolidation\mTSPt2\mTSPt2_breakdown.xlsx'
ASR_list = r'C:\Users\CHA82870\Mott MacDonald\AERMOD Modelling Services - Do\06 Working\Program\REAL\ASR_list.xlsx'

grids = ['48_38']

#comment out if not use
pollutants = 1 #TSP
#pollutants = 2 #RSP (PM10) daily
#pollutants = 3 #RSP (PM10) annual
#pollutants = 4 #FSP (PM2.5) daily
#pollutants = 5 #FSP (PM2.5) annual

##############################
#CODES DO NOT MODIFY
##############################

if pollutants == 1: #TSP
    factor_annual = 1
    factor_daily = 1
    factor_aermod = 1
    
    RSP_an_adj = 0
    RSP_10_adj = 0
    
    hourlyExceedance = 500
    dailyExceedance = 100 #not used

elif pollutants == 2: #RSP daily
    factor_annual = 1
    factor_daily = 1 
    factor_aermod = 1
    
    RSP_an_adj = 15.6
    RSP_10_adj = 26.5
    
    hourlyExceedance = 500 #not used
    dailyExceedance = 100

elif pollutants == 3: #RSP annual
    factor_annual = 1
    factor_daily = 1 
    factor_aermod = 1
    
    RSP_an_adj = 15.6
    RSP_10_adj = 26.5
    
    hourlyExceedance = 500 #not used
    dailyExceedance = 100 #not used

elif pollutants == 4: #FSP daily
    factor_annual = 0.71
    factor_daily = 0.75
    factor_aermod = 1
    
    RSP_an_adj = 15.6*factor_annual
    RSP_10_adj = 26.5*factor_daily
    
    hourlyExceedance = 500 #not used
    dailyExceedance = 75

elif pollutants == 5: #FSP annaul
    factor_annual = 0.71
    factor_daily = 0.75
    factor_aermod = 1
    
    RSP_an_adj = 15.6*factor_annual
    RSP_10_adj = 26.5*factor_daily
    
    hourlyExceedance = 500 #not used
    dailyExceedance = 100 #not used

import pandas as pd
from pandas import ExcelWriter
import math
import numpy as np

import glob

def getFiles (path, type):
    filteredFiles = []
    allFiles = glob.glob(path + "/*.{}".format(type))
    
    return allFiles

def readAermod(aermod):
    xls = pd.ExcelFile(aermod)
    aermod = xls.parse('Sheet1')
    aermod = aermod.drop(aermod.columns[[0,1]], axis = 1)
    aermod.columns = aermod.loc[0] #set columns equal to row 1
    aermod = aermod.drop([0,1])
    aermod = aermod[:-24] #drop last 24 rows
    aermod = aermod.dropna(axis = 1, how='all') #drop nan columns
    aermod = aermod.reset_index(drop=True)
    lstCols = aermod.columns.tolist()
    lstCols[0]='Time'
    aermod.columns = lstCols
    aermod = aermod.apply(pd.to_numeric)

    return aermod

df_list = []
sheet_name = []

#sum all models
chimney = pd.DataFrame()
construction = pd.DataFrame()
caline = pd.DataFrame()
marine = pd.DataFrame()

for grid in grids:
    if path_chimney != '':
        for file in getFiles(path_chimney, 'xlsx'):
            if file.find(grid) != -1:
                chimney_data = readAermod(file)
                try:
                    chimney = pd.merge(chimney, chimney_data, on=['Time'])
                except:
                    chimney = chimney_data

                summation = chimney_data.copy() #added copy to remove the impact of changing underlying chimney_data

    if path_construction != '':   
        for file in getFiles(path_construction, 'xlsx'):
            if file.find(grid) != -1:
                construction_data = readAermod(file)
                try:
                    construction = pd.merge(construction, construction_data, on=['Time'])
                except:
                    construction = construction_data

                summation.iloc[:,1:] = summation.iloc[:,1:].add(construction_data, fill_value = 0)

# for file in getFiles(path_caline, 'xlsx'):
#     if file.find(grid) != -1:
#         xls = pd.ExcelFile(file)
#         model1_data = xls.parse('Sheet1')
#         model1_data = model1_data.drop(model1_data.columns[0], axis=1)
#         model1_data.insert(0, column = 'Time', value = chimney_data['Time'].tolist())
#         model1_data = model1_data.apply(pd.to_numeric)
#         print(model1_data)
#         summation = summation.add(model1_data, fill_value = 0)

#summation['Time'] = summation['Time']/2 #problematic when not all grids have been run

    if path_marine != '':   
        for file in getFiles(path_marine, 'xlsx'):
            if file.find(grid) != -1:
                marine_data = readAermod(file)
                try:
                    marine = pd.merge(marine, marine_data, on=['Time'])
                except:
                    marine = marine_data

                summation.iloc[:,1:] = summation.iloc[:,1:].add(marine_data, fill_value = 0)

    save_path = path_all + '\\'+ "{}.xlsx".format(grid)
    summation.to_excel(save_path)
 
files_cmaq = getFiles(path_cmaq, 'txt')
files_aermod = getFiles(path_all, 'xlsx')
files_chimney = getFiles(path_chimney, 'xlsx')
files_construction = getFiles(path_construction, 'xlsx')
files_caline = getFiles(path_caline, 'xlsx')
files_marine = getFiles(path_marine, 'xlsx')

def matrix(cmaq, aermod, factor_daily, factor_annual, factor_aermod, parse):
    data = pd.read_csv(cmaq, sep='\s+')
    data = data.drop([0,1], axis = 0)
    data = data.apply(pd.to_numeric)
    
    index = data.index.tolist()
    re_index = index[-7:] + index[:-7]
    data = data.reindex(re_index) #move the last 8 rows to the top
    
    for i in list(range(7)): #set YYYY to be the year in the 9th row
        data.iloc[i,0] = data.iloc[7,0]
    
    data = data.reset_index(drop=True)
    
    if parse == False:
        xls = pd.ExcelFile(aermod)
        aermod = xls.parse('Sheet1')
    elif parse == True:
        aermod = readAermod(aermod)

    factor_daily = factor_daily #factor for cmeq
    factor_aermod = factor_aermod #factor for aermod

    aermodPath = pd.DataFrame()
    aermodPath['Time'] = aermod[aermod.columns[0]]

    for cols in aermod.columns[1:]: #skip index
        aermodPath[cols] = aermod[cols]*factor_aermod + data['RSP']*factor_daily
    
    #for annual average (path data included)

    aermod_an = pd.DataFrame()
    aermod_an['Time'] = aermod[aermod.columns[0]]

    for cols in aermod.columns[1:]: #skip index
        aermod_an[cols] = aermod[cols]*factor_aermod + data['RSP']*factor_annual
    
    return data[['Year', 'mm','dd','hh','RSP']], aermod, aermodPath, aermod_an

def populateDataFrame(files_aermod, files_cmaq, parse):
    PATH = pd.DataFrame()
    AERMOD =  pd.DataFrame()
    AERMODPATH =  pd.DataFrame()
    AERMODPATH_24 = pd.DataFrame(dtype=float)
    AERMODAN = pd.DataFrame()

    for file_aermod in files_aermod: #loop through each aermod files
        for grid in grids: #loop through each grid
            if file_aermod.find(grid) != -1: #check if grid match the name of aermod
                for file_cmaq in files_cmaq: # loop through the cmaq files
                    if file_cmaq.find(grid) != -1: #find the one that match with aermod
                        PATH_temp, AERMOD_temp, AERMODPATH_temp, AERMODAN_temp = matrix(file_cmaq, file_aermod, factor_daily, factor_annual, factor_aermod, parse)
                        if len(PATH) == 0:
                            PATH = PATH_temp
                            AERMOD = AERMOD_temp
                            AERMODPATH = AERMODPATH_temp
                            AERMODAN = AERMODAN_temp
                            
                        else:
                            PATH = pd.merge(PATH, PATH_temp, on=['Year', 'mm','dd','hh'])
                            AERMOD = pd.merge(AERMOD, AERMOD_temp, on=['Time'])
                            AERMODPATH = pd.merge(AERMODPATH, AERMODPATH_temp, on=['Time'])
                            AERMODAN = pd.merge(AERMODAN, AERMODAN_temp, on=['Time'])

    AERMODPATH_24 = AERMODPATH.groupby(np.arange(len(AERMODPATH))//24).mean()
    AERMODPATH_24['Time'] = AERMODPATH.iloc[::24, 0].tolist() #set begining time as col time

    AERMOD_24 = AERMOD.groupby(np.arange(len(AERMOD))//24).mean()

    return PATH, AERMOD, AERMOD_24, AERMODPATH, AERMODPATH_24, AERMODAN

if path_all != "":
    PATH, AERMOD, AERMOD_24, AERMODPATH, AERMODPATH_24, AERMODAN = populateDataFrame(files_aermod, files_cmaq, False)

    #create excel tab
    df_list.append(PATH)
    sheet_name.append('PATH')

    cols = AERMOD.columns.tolist()
    cols = cols[:1] + sorted(cols[1:])

    AERMOD = AERMOD[cols]
    AERMOD_24 = AERMOD_24[cols]
    AERMODPATH = AERMODPATH[cols]
    AERMODPATH_24 = AERMODPATH_24[cols]
    AERMODAN = AERMODAN[cols]

    AERMODPATH_24.iloc[:,1:] = AERMODPATH_24.iloc[:,1:] + RSP_10_adj

if path_chimney != "":
    PATH_chimney, AERMOD_chimney, AERMOD_24_chimney, AERMODPATH_chimney, AERMODPATH_24_chimney, AERMODAN_chimney = populateDataFrame(files_chimney, files_cmaq, True)

    cols_chinmey = list(set(cols) & set(chimney.columns.tolist())) #sometime, tier 2 would not have the full list
    cols_chinmey.remove('Time')
    cols_chinmey = sorted(cols_chinmey)
    cols_chinmey = ['Time'] + cols_chinmey
    chimney = chimney[cols]

    df_list.append(chimney)
    sheet_name.append('Chimney')

if path_construction != "":
    PATH_construction, AERMOD_construction, AERMOD_24_construction, AERMODPATH_construction, AERMODPATH_24_construction, AERMODAN_construction = populateDataFrame(files_construction, files_cmaq, True)

    cols_construction = list(set(cols) & set(construction.columns.tolist())) #sometime, tier 2 would not have the full list
    cols_construction.remove('Time')
    cols_construction = sorted(cols_construction)
    cols_construction = ['Time'] + cols_construction
    construction = construction[cols_construction]

    df_list.append(construction)
    sheet_name.append('Construction')

if path_caline != "":
    PATH_caline, AERMOD_caline, AERMOD_24_caline, AERMODPATH_caline, AERMODPATH_24_caline, AERMODAN_caline = populateDataFrame(files_caline, files_cmaq, True)

    cols_caline = list(set(cols) & set(caline.columns.tolist())) #sometime, tier 2 would not have the full list
    cols_caline.remove('Time')
    cols_caline = sorted(cols_caline)
    cols_caline = ['Time'] + cols_caline
    caline = caline[cols_caline]

    df_list.append(caline)
    sheet_name.append('Caline')

if path_marine != "":
    PATH_marine, AERMOD_marine, AERMOD_24_marine, AERMODPATH_marine, AERMODPATH_24_marine, AERMODAN_marine = populateDataFrame(files_marine, files_cmaq, True)  

    cols_marine = list(set(cols) & set(marine.columns.tolist())) #sometime, tier 2 would not have the full list
    cols_marine.remove('Time')
    cols_marine = sorted(cols_marine)
    cols_marine = ['Time'] + cols_marine 
    marine = marine[cols_marine]

    df_list.append(marine)
    sheet_name.append('Marine')

df_list.append(AERMOD)
df_list.append(AERMOD_24)
df_list.append(AERMODPATH)
df_list.append(AERMODPATH_24)
sheet_name.extend(['Sum of Models', 'Sum of Models_24','Sum of Models_PATH', 'Sum of Models_PATH_24'])

def get_nlargest(df, n, adj):
    result = {}
    for cols in df.columns[1:]: #skip index
        tem = df[cols].nlargest(n).tolist()[-1] + adj
        result[cols] = tem
    return result

def nthProjectContribution(df_project, df_total, n, factor_aermod, adj):
    result = {}
    for cols in df_project.columns[1:]: #skip index
        value_total = df_total[cols].nlargest(n).tolist()[-1] + adj
        index = df_total[cols].nlargest(n).index.tolist()[-1]
        value_project = df_project.loc[index, cols]        
        result[cols] = (value_project*factor_aermod)/value_total
    
    return result

def nthbgContribution(df_project, df_total, n, factor_aermod, adj):
    result = {}
    for cols in df_project.columns[1:]: #skip index
        value_total = df_total[cols].nlargest(n).tolist()[-1] + adj
        index = df_total[cols].nlargest(n).index.tolist()[-1]
        value_project = df_project.loc[index, cols]        
        result[cols] = 1- ((value_project*factor_aermod)/value_total)
    
    return result

def nthProjectContribution_breakdown(df_index, df_project, df_total, n, factor_aermod, adj):
    result = {}
    for cols in df_project.columns[1:]: #skip index
        value_total = df_total[cols].nlargest(n).tolist()[-1] + adj
        index = df_total[cols].nlargest(n).index.tolist()[-1]
        value_breakdown = df_index.loc[index, cols]        
        result[cols] = (value_breakdown*factor_aermod)/value_total
    
    return result

lst = []
lst.append(get_nlargest(AERMODPATH, 1, 0)) #Max Hourly
lst.append(get_nlargest(AERMODPATH_24, 10,0)) #10th Max Daily
lst.append(get_nlargest(AERMODPATH, 19, 0)) #19th Max Hourly
lst.append((AERMODAN.iloc[:,1:].mean() + RSP_an_adj).to_dict()) #annual average
lst.append(AERMODPATH.iloc[:,1:][AERMODPATH>hourlyExceedance].count().to_dict()) #Exceedance of hourly
lst.append(AERMODPATH_24.iloc[:,1:][AERMODPATH_24>dailyExceedance].count().to_dict()) #Exceedance of daily
lst.append(((AERMOD.iloc[:,1:].mean()/factor_aermod)/((AERMODAN.iloc[:,1:].mean() + RSP_an_adj))).to_dict()) #annual project contribution
lst.append((1-(AERMOD.iloc[:,1:].mean()/factor_aermod)/((AERMODAN.iloc[:,1:].mean() + RSP_an_adj))).to_dict()) #annaul background contribution

lst.append(nthProjectContribution(AERMOD_24, AERMODPATH_24, 10, factor_aermod, 0)) #10th daily max -project contribution (zero becoz added already)
lst.append(nthbgContribution(AERMOD_24, AERMODPATH_24, 10, factor_aermod, 0)) #10th daily max -background contribution


lst.append(nthProjectContribution(AERMOD, AERMODPATH, 1, factor_aermod, 0)) #max hourly -project contribution
lst.append(nthbgContribution(AERMOD, AERMODPATH, 1, factor_aermod, 0)) #Max hourly max -background contribution

lst.append(nthProjectContribution(AERMOD, AERMODPATH, 19, factor_aermod, 0)) #19th hourly -project contribution
lst.append(nthbgContribution(AERMOD, AERMODPATH, 19, factor_aermod, 0)) #19th hourly max -background contribution

#model breakdown
if path_chimney != "":
    lst.append(((AERMOD_chimney.iloc[:,1:].mean()/factor_aermod)/((AERMODAN.iloc[:,1:].mean() + RSP_an_adj))).to_dict()) #annual project contribution
    lst.append(nthProjectContribution_breakdown(AERMOD_24_chimney, AERMOD_24, AERMODPATH_24, 10, factor_aermod, 0)) #10th daily max -project contribution (zero becoz added already)
    lst.append(nthProjectContribution_breakdown(AERMOD_chimney, AERMOD, AERMODPATH, 1, factor_aermod, 0)) 
    lst.append(nthProjectContribution_breakdown(AERMOD_chimney, AERMOD, AERMODPATH, 19, factor_aermod, 0)) 
if path_construction != "":
    lst.append(((AERMOD_construction.iloc[:,1:].mean()/factor_aermod)/((AERMODAN.iloc[:,1:].mean() + RSP_an_adj))).to_dict()) #annual project contribution
    lst.append(nthProjectContribution_breakdown(AERMOD_24_construction, AERMOD_24, AERMODPATH_24, 10, factor_aermod, 0)) #10th daily max -project contribution (zero becoz added already)
    lst.append(nthProjectContribution_breakdown(AERMOD_construction, AERMOD, AERMODPATH, 1, factor_aermod, 0))
    lst.append(nthProjectContribution_breakdown(AERMOD_construction, AERMOD, AERMODPATH, 19, factor_aermod, 0)) 
if path_caline != "":
    lst.append(((AERMOD_caline.iloc[:,1:].mean()/factor_aermod)/((AERMODAN.iloc[:,1:].mean() + RSP_an_adj))).to_dict()) #annual project contribution
    lst.append(nthProjectContribution_breakdown(AERMOD_24_caline, AERMOD_24, AERMODPATH_24, 10, factor_aermod, 0)) #10th daily max -project contribution (zero becoz added already)
    lst.append(nthProjectContribution_breakdown(AERMOD_caline, AERMOD, AERMODPATH, 1, factor_aermod, 0))
    lst.append(nthProjectContribution_breakdown(AERMOD_caline, AERMOD, AERMODPATH, 19, factor_aermod, 0))
if path_marine != "":
    lst.append(((AERMOD_marine.iloc[:,1:].mean()/factor_aermod)/((AERMODAN.iloc[:,1:].mean() + RSP_an_adj))).to_dict()) #annual project contribution
    lst.append(nthProjectContribution_breakdown(AERMOD_24_marine, AERMOD_24, AERMODPATH_24, 10, factor_aermod, 0)) #10th daily max -project contribution (zero becoz added already)
    lst.append(nthProjectContribution_breakdown(AERMOD_marine, AERMOD, AERMODPATH, 1, factor_aermod, 0))
    lst.append(nthProjectContribution_breakdown(AERMOD_marine, AERMOD, AERMODPATH, 19, factor_aermod, 0))

summary = pd.DataFrame(lst)
lst_index = ['Max hourly','10th Max Daily','19th Max hourly','Annual average','Exceedance of hourly',
                   'Exceedance of daily','Annual project contribution','Annual background contribution','10th Daily Max - project contribution',
                   '10th Daily Max - background contribution', 'Max Hourly project contribution','Max Hourly background contribution',
                   '19th Max Hourly project contribution','19th Max Hourly background contribution']
if path_chimney != "":
    lst_index.extend(['Annual project contribution (chimney)', '10th Daily Max - project contribution (chimney)', 'Max Hourly project contribution (chimney)', '19th Max Hourly project contribution (chimney)'])

if path_construction != "":
    lst_index.extend(['Annual project contribution (construction)', '10th Daily Max - project contribution (construction)', 'Max Hourly project contribution (construction)', '19th Max Hourly project contribution (construction)'])

if path_caline != "":
    lst_index.extend(['Annual project contribution (caline)', '10th Daily Max - project contribution (caline)', 'Max Hourly project contribution (caline)', '19th Max Hourly project contribution (caline)'])

if path_marine != "":
    lst_index.extend(['Annual project contribution (marine)', '10th Daily Max - project contribution (marine)', 'Max Hourly project contribution (marine)', '19th Max Hourly project contribution (marine)'])

summary['Index'] = lst_index
cols = summary.columns.tolist()
cols = cols[-1:]+cols[:-1] #rearrange cols
summary= summary[cols]

if pollutants == 1:
    lst_rows = ['Max hourly', 'Exceedance of hourly', 'Max Hourly project contribution','Max Hourly background contribution']
    if path_chimney != "":
        lst_rows.extend(['Max Hourly project contribution (chimney)'])
    if path_construction != "":
        lst_rows.extend(['Max Hourly project contribution (construction)'])
    if path_caline != "":
        lst_rows.extend(['Max Hourly project contribution (caline)'])
    if path_marine != "":
        lst_rows.extend(['Max Hourly project contribution (marine)'])
    
    summary = summary[summary['Index'].isin(lst_rows)]

elif pollutants == 2:
    lst_rows = ['10th Max Daily', 'Exceedance of daily','10th Daily Max - project contribution','10th Daily Max - background contribution']
    if path_chimney != "":
        lst_rows.extend(['10th Daily Max - project contribution (chimney)'])
    if path_construction != "":
        lst_rows.extend(['10th Daily Max - project contribution (construction)'])
    if path_caline != "":
        lst_rows.extend(['10th Daily Max - project contribution (caline)'])
    if path_marine != "":
        lst_rows.extend(['10th Daily Max - project contribution (marine)'])

    summary = summary[summary['Index'].isin(lst_rows)]

elif pollutants == 3:
    lst_rows = ['Annual average', 'Annual project contribution','Annual background contribution']
    if path_chimney != "":
        lst_rows.extend(['Annual project contribution (chimney)'])
    if path_construction != "":
        lst_rows.extend(['Annual project contribution (construction)'])
    if path_caline != "":
        lst_rows.extend(['Annual project contribution (caline)'])
    if path_marine != "":
        lst_rows.extend(['Annual project contribution (marine)'])

    summary = summary[summary['Index'].isin(lst_rows)]
    
elif pollutants == 4:
    lst_rows = ['10th Max Daily', 'Exceedance of daily','10th Daily Max - project contribution','10th Daily Max - background contribution']
    if path_chimney != "":
        lst_rows.extend(['10th Daily Max - project contribution (chimney)'])
    if path_construction != "":
        lst_rows.extend(['10th Daily Max - project contribution (construction)'])
    if path_caline != "":
        lst_rows.extend(['10th Daily Max - project contribution (caline)'])
    if path_marine != "":
        lst_rows.extend(['10th Daily Max - project contribution (marine)'])

    summary = summary[summary['Index'].isin(lst_rows)]

elif pollutants == 5:
    lst_rows = ['Annual average', 'Annual project contribution','Annual background contribution']
    if path_chimney != "":
        lst_rows.extend(['Annual project contribution (chimney)'])
    if path_construction != "":
        lst_rows.extend(['Annual project contribution (construction)'])
    if path_caline != "":
        lst_rows.extend(['Annual project contribution (caline)'])
    if path_marine != "":
        lst_rows.extend(['Annual project contribution (marine)'])

    summary = summary[summary['Index'].isin(lst_rows)]

lst = pd.read_excel(ASR_list)
lst = lst['ASRS'][lst['ASRS'].isin(summary.columns.tolist())].tolist() #filter ASRs that is in the sumamry table columns
lst.insert(0, "Index")
summary = summary[lst]
df_list.append(summary)

summary_T = summary.T
summary_T.columns = summary_T.iloc[0]
summary_T = summary_T.drop('Index', axis = 0)
summary_T = summary_T.reset_index()

df_list.append(summary_T)

sheet_name.extend(['Summary', 'Summary Trasposed'])

def save_xls(list_dfs, xls_path, sheet_name):
    writer = ExcelWriter(xls_path)
    for n, df in zip(sheet_name, list_dfs):
        df.to_excel(writer,'%s' % n, index = False)
    writer.save()

save_xls(df_list, excel_name, sheet_name)