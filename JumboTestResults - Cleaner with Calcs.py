# -*- coding: utf-8 -*-
"""
Created on Thu Nov  5 12:45:24 2020

@author: gbadafow
"""

# main cleaning Python program - objective open Excel Jumbo test reports, remove columns not needed, add some calculations and save to CVS file for analysis

import pandas as pd
import numpy as np

dfexcel = pd.read_excel(r'M:\Kemsley Paper Mill\Technical\Technical\PM4\PM4 Daily Analysis\Jumbo Test Results\Jumbo test results - All Data.xls') ##importing Excel file

df = dfexcel.drop(['M/C','Ash','Black Spot 200 mm2','Blister','Boat Jumbo 30min Y N','Boat QA','Boat QC','Boat QC uncured','Boat Off Winder Uncured','Boat Winder Uncured Set 2',
                   'Bright (457B) ~','Bright (457B) Lablink','Bright 457B QC','Bright UV Ex ~','Bright UV Ex Lablink','Bright UV Ex QC','Burst factor QA','Burst QA',
                   'Cobb 120m TS QA','Cobb 1m BS cured QA','Cobb 1m BS uncured QA',
                   'Cobb 1m TS uncured QA','Cobb 1m TS cured QA','Cobb 30m TS QA','Cobb 30min TS QC','Cobb 3m BS HOT QC','Cobb 5m BS QC','Cobb 5m TS QC','Density','DW VAR MD',
                   'Edge crack','Gsm QC','Hole 1000 mm2','Hot melt/wax QC','ISO Brightness QC','ISO Brightness~','Lab RH QC','Lab temp QC','Large Split','Ply Sep QC',
                   'Ply Wt B BD QC','Ply Wt M BD QC','Ply Wt T BD QC','Plybond QC','PM3 Cationic Starch D Cooker Alert','Roughness TS QC','SCT CD QA',
                   'SCT factor QA','SCT MD QC','Shade a* Lablink','Shade a QC Elrepho ','Shade b* Lablink','Shade b QC Elrepho ',
                   'Shade Delta E* Lablink','Shade Delta E* QA~','Shade Delta E QC Elrepho ','Shade L* Lablink','Shade L QC Elrepho ','Starch Spray Status','Surface Coverage QC',
                   'Tensile CD 25 QC','Tensile DRY CD 25 QC','Tensile DRY CD AL','Tensile DRY MD 25 QC','Tensile DRY MD AL','Tensile DRY MD 25 QA','Tensile DRY CD 25 QA',
                   'Tensile MD 25 QC','Tensile Ratio Dry QC','Tensile Ratio','Valsizer Alert BS','Valsizer Alert TS','Val Sizer Status','Wax Pick QA','Wet Exp QC',
                   'White Flake QC','Z Strength QA'], axis=1) # drops all the unnessecary columns

df.fillna(100,inplace=True)

df['Turned Up'] = df['Turned Up'].astype(str)

df[['Date','Time']] = df['Turned Up'].str.split(" ", expand=True)
df[['y','m','d']] = df['Date'].str.split("-", expand=True)

df['Month/Year'] = df['m']+"/"+df['y'] 

df['SCT Index'] = (df['SCT CD ~'] /df['Gsm QC ~'])*1000
df['Burst Index'] = (df['Burst~']/df['Gsm QC ~'])
df['CMT0 Index'] = (df['CMT0 QC'] /df['Gsm QC ~'])

df['SCT Retest'] = np.where(df['SCT CD QC'] == 100, 0, 1)
df['Burst Retest'] = np.where(df['Burst Average'] == 100, 0, 1)
df['Porosity Retest'] = np.where(df['Porosity QC'] == 100, 0, 1)

df.to_csv(r'M:\Kemsley Paper Mill\Technical\Technical\PM4\PM4 Daily Analysis\Jumbo Test Results\JumboTestReport-Clean.csv', index=False)