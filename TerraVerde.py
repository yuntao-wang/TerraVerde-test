# -*- coding: utf-8 -*-
"""
Created on Sat Mar 24 14:41:50 2018

@author: movin
"""

import numpy as np
import pandas as pd
    
def extract(df, date):
    """extract daily value as a separate dataframe,
    input: dataframe, date
    output: a dataframe of dimension 2x96,
    row 1 is channel 101 data, row 2 is channel 102 data"""
    
    df_out = df[df['INTRVL_DATE'] == date]
    
    return df_out
    

def convert(df):
    """input a day value dataframe, 
    output dataframe in format 96 rows x 3 columns
    101 kW, 102 kW and net kWh"""
    
    # read in daily value, transpose matrix
    df_day = df.ix[0:2, 'MIN_0015_KWH':'MIN_2400_KWH'].T
    
    # calculate kW from kWh
    df_day_kW = df_day / 0.25
    
    # calculate net kW value
    df_day_kW['net'] = df_day_kW.iloc[:,0] - df_day_kW.iloc[:,1]
    
    # calculate net kWh
    df_day_result = df_day_kW.copy()
    df_day_result['net'] = df_day_result['net'] * 0.25
    
    # rename columns
    data = {'CHNL_ID 101 (kW)':df_day_result.iloc[:,0],
       'CHNL_ID 102 (kW)':df_day_result.iloc[:,1],
       'Net Usage (kWh)':df_day_result.iloc[:,2]}
    df_day_result = pd.DataFrame(data)
    
    return df_day_result


def main():

    # read in excel file
    xlsx_file = pd.ExcelFile('Raw_Interval_Data.xlsx')
    tab_list = ['1','2','3','4','6','9','10']
    
    writer = pd.ExcelWriter('output.xlsx')
    
    for tab in tab_list:
        df = xlsx_file.parse(tab, header=0)
        
        # create an empty dataframe
        FinalDf = pd.DataFrame()
          
        # extract each day value and append to a final dataframe
        for date in pd.date_range(start='1/1/2017', end='12/31/2017'):
            DayDf = extract(df, date)
            DayResDf = convert(DayDf)
            FinalDf = pd.concat([FinalDf, DayResDf])
        
        # write final data to excel sheet
        
        FinalDf.to_excel(writer, tab)
        writer.save()

main()