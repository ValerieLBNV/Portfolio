# -*- coding: utf-8 -*-
"""
Created on Fri Oct 28 16:09:27 2022

@author: valerie.van
"""

import pandas as pd
from pathlib import Path
import numpy as np
import re


SOURCE_FOLDER = "C:\\Users\\valerie.van\\Perfetti Van Melle\\Supply Chain - Global Plan & Deliver - Demand & Supply Review (1)\\"

OUTPUT_CONSOLIDATED = "C:\\Users\\valerie.van\\Perfetti Van Melle\\Supply Chain - Global Plan & Deliver - Documents (1)\\Sales and Operations Planning\\Database\\S&OP_Consolidated_Volume.csv"

OUTPUT_FOLDER_BYTYPE = "C:\\Users\\valerie.van\\Perfetti Van Melle\\Supply Chain - Plan & Deliver (Core team) - Documents\\Plan and Deliver - Data\\S&OP\\"
ACTUAL = "S&OP_Actual_Volume.csv"
FORECAST = "S&OP_Forecast_Volume.csv"

LOOKUP_FOLDER = "C:\\Users\\valerie.van\\Perfetti Van Melle\\Supply Chain - Plan & Deliver (Core team) - Documents\\Plan and Deliver - Data\\"
OC_LIST = "OC list - Plan Deliver.csv"
FACTORIES_LIST = "Factories list - Plan Deliver.csv"

Type_Dict = {True: 'Forecast', False: 'Actual'}
Cat_Dict = {'demand': 'Demand', 'production': 'Production'}

# =============================================================================
# READ LOOKUP FILES
# =============================================================================

oc_list = pd.read_csv(LOOKUP_FOLDER+OC_LIST)
factories_list = pd.read_csv(LOOKUP_FOLDER+FACTORIES_LIST)

# =============================================================================
# SOURCE DEMAND AND PRODUCTION DATAFRAME
# =============================================================================

p= Path(SOURCE_FOLDER)

filelist_demand = []
for f in [x for x in p.glob("*demand*.xlsx")]:
    df= pd.read_excel(f)
    df['filename'] = Path(f).stem
    df.drop('Item ID', axis=1, inplace=True)    
    df_unpivot = df.melt(id_vars=['filename','Version', 'BU', 'Sales OC', 'Markets', 'SKU', 'SKU Description'], 
                          var_name="Month", value_name= "Value")
    df_unpivot = df_unpivot.loc[df_unpivot["Value"] != 0]
    df_unpivot = df_unpivot[df_unpivot.Value.notnull()]
    df_unpivot = df_unpivot[df_unpivot["Version"].str.contains("Total") == False] 
    df_unpivot = df_unpivot[df_unpivot["Version"].str.contains("Budget") == False] 
    filelist_demand.append(df_unpivot)
demand = pd.concat(filelist_demand)
demand = demand.rename(columns={'Markets':'Channel','Sales OC':'OC'})


demand['Version'] = demand['Version'].str.lstrip('0')
demand['Version'] = pd.to_datetime(demand['Version'],format='%m-%Y',errors='coerce').dt.date
demand['Month'] = demand['Month'].str.lstrip('0')
demand['Month'] = pd.to_datetime(demand['Month'],format='%m-%Y',errors='coerce').dt.date

demand = pd.merge(demand,oc_list[['source_system','operating_company']],how='left',left_on='OC',right_on='operating_company')
demand['SKU'] = demand['SKU'].astype(str).replace('\.0', '', regex=True)
demand[['Item', 'XX']] = demand['SKU'].str.split('-X', 1, expand=True)
demand['Item'] = demand['Item'].astype(str).str.pad(18, side= 'left', fillchar= '0')
demand['ItemID'] = demand['source_system'] + demand['Item']
demand.drop(['operating_company','Item', 'XX'],axis=1, inplace=True)
demand['Value'] = pd.to_numeric(demand['Value'], errors='coerce')

demand_cat = demand['filename'].str.split("_", n = 2, expand = True)
demand['Category'] = demand_cat[2]
demand['Category'] = demand['Category'].replace(Cat_Dict)
demand['Type'] = demand.eval('Month >= Version')
demand['Type'] = demand['Type'].replace(Type_Dict)



filelist_production = []
for f in [x for x in p.glob("*production*.xlsx")]:
    df2 = pd.read_excel(f)
    df2['filename'] = Path(f).stem
    df2.drop('Item ID', axis=1, inplace=True)
    df2_unpivot = df2.melt(id_vars=['filename','Version', 'BU', 'Plant', 'Channel', 'SKU', 'SKU Description'], 
                         var_name="Month", value_name= "Value")
    df2_unpivot = df2_unpivot.loc[df2_unpivot["Value"] != 0]
    df2_unpivot = df2_unpivot[df2_unpivot.Value.notnull()]
    df2_unpivot = df2_unpivot[df2_unpivot["Version"].str.contains("Total") == False]
    df2_unpivot = df2_unpivot[df2_unpivot["Version"].str.contains("Budget") == False]
    filelist_production.append(df2_unpivot)
production = pd.concat(filelist_production)


production['Version'] = production['Version'].str.lstrip('0')   
production['Version'] = pd.to_datetime(production['Version'],format='%m-%Y',errors='coerce').dt.date
production['Month'] = production['Month'].str.lstrip('0')
production['Month'] = pd.to_datetime(production['Month'],format='%m-%Y',errors='coerce').dt.date

production['Plant'] = production['Plant'].str.replace('CN Shenzhen','CN Shenzen')
production = pd.merge(production, factories_list[['ocg_id','factory_name']], how='left', left_on='Plant', right_on='factory_name')
production = pd.merge(production, oc_list[['source_system', 'operating_company','ocg_id']], how='left',on='ocg_id')
production = production.rename(columns={'operating_company':'OC'})
production['SKU'] = production['SKU'].astype(str).replace('\.0', '', regex=True)
production[['Item', 'XX']] = production['SKU'].str.split('-X', 1, expand=True)
production['Item'] = production['Item'].astype(str).str.pad(18, side= 'left', fillchar= '0')
production['ItemID'] = production['source_system'] + production['Item']
production.drop(['ocg_id','factory_name','Item', 'XX'],axis=1, inplace=True)
production['Value'] = pd.to_numeric(production['Value'], errors='coerce')

production_cat = production['filename'].str.split("_", n = 2, expand = True)
production['Category'] = production_cat[2]
production['Category'] = production['Category'].replace(Cat_Dict)
production['Type'] = production.eval('Month >= Version')
production['Type'] = production['Type'].replace(Type_Dict)


# =============================================================================
# APPEND DATAFRAMES AND SPLIT BY TYPE 
# =============================================================================

appended_data = demand.append(production)
appended_data['ItemID'] = appended_data['ItemID'].str.replace('npd','NPD')
appended_data = appended_data[['filename', 'Version', 'BU', 'OC', 'Plant', 'Channel', 'Category','Type',
                               'SKU','source_system', 'ItemID','SKU Description','Value', 'Month']]
actual = appended_data[appended_data['Type'] == 'Actual']
forecast = appended_data[appended_data['Type'] == 'Forecast']

# =============================================================================
# EXPORT TO CSV
# =============================================================================

appended_data.to_csv(OUTPUT_CONSOLIDATED, index=False)
actual.to_csv(OUTPUT_FOLDER_BYTYPE+ACTUAL, index=False)
forecast.to_csv(OUTPUT_FOLDER_BYTYPE+FORECAST, index=False)
