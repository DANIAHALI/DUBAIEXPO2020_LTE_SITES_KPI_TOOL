import pandas
import openpyxl
from openpyxl.styles import colors
from openpyxl.styles import Font, Color
from openpyxl.styles.borders import Border, Side
from openpyxl import load_workbook
from openpyxl.chart import BarChart3D, Reference, AreaChart, AreaChart3D, Series, LineChart, LineChart3D
import os
import pdb
from openpyxl.styles import NamedStyle, Font, Border, Side
# import excel2img
import time
import datetime
import xlsxwriter
import math

def OD_IBS_SLA_STATUS(input_1, Output_path):
    print('Processing Start Please Wait!')
    start = time.time()
    print('Processing!')

    input_1_df = pandas.read_excel(input_1)

    input_1_df = input_1_df.fillna(0)
    input_1_df['L.UL.Interference.Avg(dBm)'].replace({"NIL": 0}, inplace=True)
    eNodeBName = list(set(input_1_df['eNodeB Name']))

    # input_1_df_ = input_1_df.groupby('CELLNAME')['VS.MeanRTWP(dBm)'].mean()



    eNodeBName_ = []
    CellName_ = []
    DownlinkEARFCN = []
    PhysicalcellID = []
    ulInterf_val = []
    ulInterf_dec = []


    for i in range(len(eNodeBName)):
        input_1_df__ = input_1_df[input_1_df['eNodeB Name'] == eNodeBName[i]][-3:]

        ULInterf_ = input_1_df__['L.UL.Interference.Avg(dBm)'].mean()


        ulInterf_val.append(ULInterf_)

        if ULInterf_ > -112 and ULInterf_ != 0:
            ulInterf_dec.append('High UL Interference')
        elif ULInterf_ <= -112 and ULInterf_ != 0:
            ulInterf_dec.append('UL Interference Fine')
        elif ULInterf_ == 0:
            ulInterf_dec.append('Locked')


        enodebnm = list(set(input_1_df__['eNodeB Name']))[0]
        eNodeBName_.append(enodebnm)

        celnm = list(set(input_1_df__['Cell Name']))[0]
        CellName_.append(celnm)

        dlfcn = list(set(input_1_df__['Downlink EARFCN']))[0]
        DownlinkEARFCN.append(dlfcn)

        pcelid = list(set(input_1_df__['Physical cell ID']))[0]
        PhysicalcellID.append(pcelid)

    output_df = pandas.DataFrame()

    output_df['eNodeB Name'] = eNodeBName_
    output_df['Cell Name'] = CellName_
    output_df['Downlink EARFCN'] = DownlinkEARFCN
    output_df['Physical cell ID'] = PhysicalcellID

    output_df['L.UL.Interference.Avg(dBm)'] = ulInterf_val
    output_df['UL Interference Result'] = ulInterf_dec

    output_df['L.UL.Interference.Avg(dBm)'].replace({0 : "NIL"}, inplace=True)



    print('Output Path: ', Output_path)


    output_df.to_csv(Output_path + '\\' + input_1.split('/')[-1].split('.')[0] +'.csv', index=False)


    end = time.time()
    Execute_Time = "{:.3f}".format((end - start) / 60)
    print('The Execution Time of this Tool is %s minutes.' % Execute_Time)
    time.sleep(1)
    print('Execution Completed Succcessfully!')
    time.sleep(1)
    print('')
    print('')
    print('---------------Huawei RF Middle East----------------')
    print('---------For Support: Danish Ali(dwx854280)---------')
    print('---------------Contact: 00971508552942--------------')
    time.sleep(3)


if __name__ == '__main__':
    input_1 = str(os.getcwd()) + '\\Input\\3G Cell Level Report_Query_Result_1-11Sep.csv'
    input_2 = str(os.getcwd()) + '\\Input\\3G Cell Level Report_Query_Result_1Aug-24Aug20.csv'
    input_3 = str(os.getcwd()) + '\\Input\\3G Cell Level Report_Query_Result_25-31Aug.csv'
    input_4 = str(os.getcwd()) + '\\Input\\3G Cell Level Report_Query_Result_1-22Sep.csv'

    Output_path = str(os.getcwd()) + '\\Output'
    OD_IBS_SLA_STATUS(input_1, input_2, input_3, input_4, Output_path)
