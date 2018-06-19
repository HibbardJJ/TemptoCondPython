# -*- coding: utf-8 -*-
"""
Created on Fri Jun  1 12:02:45 2018

@author: Joshua Hibbard

"""

import re
import csv
import os
import shutil
import collections
import numpy as np
from openpyxl import *


date = input('What is the date of the experiment? Uses spaces to separate the year, month, and date. Format: YYYY MM DD') 

start_time = input('What is the time stamp of the first image? Please use the format HH MM SS MSS and separate each entry with a space.')

R = input('What is the R value? (Radiation load)')


date = date.split()
year=date[0]
month=date[1]
day=date[2]

Main_Directory = 'C:\\Users\\Joshua Hibbard\\Box\\Josh data\\Red_Blue Light\\' + month + '-' + day + '-' + year[-2:]
os.chdir(Main_Directory)


start_time_list = start_time.split()
hours=start_time_list[0]
minutes=start_time_list[1]
seconds=start_time_list[2]
milliseconds=start_time_list[3]
thermal_image_time_list=[]
"Delineates the columns in the gas exchange data with the desired values."
time_stamp_column=2
lower_before_thermo_column=26
lower_after_thermo_column=27
upper_before_thermo_column=28
upper_after_thermo_column=29
xout2_column=22
"Convert the time strings into numbers."
image_true_hours=float(hours)
image_true_minutes=float(minutes)
image_true_seconds=float(seconds)
image_true_milliseconds=float(milliseconds)
'Convert the true image time into minutes elapsed since midnight.'
first_image_time = image_true_hours*60 + image_true_minutes + image_true_seconds/60 + image_true_milliseconds/60000


list1 = os.listdir('OriginalTempImages')
number_files = len(list1)

minutes_between_images = 3
image_index = 0
"Create an array of all the image times."
while image_index < number_files:
    final_image_time = minutes_between_images*image_index + first_image_time
    thermal_image_time_list.append(final_image_time)
    image_index+=1
    

thermal_image_index = 0
file_number_index = 0
list_index = 1
d={}
data_extraction = {}

"Find the time stamps of each image and place them into a dictionary."
for filename in os.listdir('OriginalTempImages'):
    regex=re.compile('([0-9]*)(?:\\.csv)')
    time_stamp1=regex.findall(filename)
    time_stamp1=time_stamp1[0]
    time_stamp1=float(time_stamp1)
    d[time_stamp1]=filename
    
'Create the FinalTempImages folder for output'
os.makedirs(Main_Directory + '\\' + 'FinalTempImages')

od = collections.OrderedDict(sorted(d.items()))

time_stamp_list = []

with open(year + '_' + month + '_' + day + '.csv', 'r') as gas_exchange_data, open('DataExtraction.csv','w') as outputfile:
    data = csv.reader(gas_exchange_data, delimiter = ',', quotechar = '\n')
    
    while True:
        get_row = next(data)
        time_stamp=float(get_row[time_stamp_column])
        if time_stamp <= thermal_image_time_list[thermal_image_index] + 0.3 and time_stamp >= thermal_image_time_list[thermal_image_index] - 0.3:
            
            lbt=float(get_row[lower_before_thermo_column])
            lat=float(get_row[lower_after_thermo_column])
            ubt=float(get_row[upper_before_thermo_column])
            uat=float(get_row[upper_after_thermo_column])
            xout2=float(get_row[xout2_column])
            
            data_extraction[time_stamp] = [ubt, uat, lbt, lat, xout2]
        
            outputfile.write(str(time_stamp))
            outputfile.write(',')
            outputfile.write('KMatrix_23C_oneamp')
            outputfile.write(',')
            outputfile.write(str(ubt))
            outputfile.write(',')
            outputfile.write(str(uat))
            outputfile.write(',')
            outputfile.write(str(lbt))
            outputfile.write(',')
            outputfile.write(str(lat))
            outputfile.write(',')
            outputfile.write(str(xout2))
            outputfile.write('\n')
        
            thermal_image_index+=1        
            
                
            if file_number_index < number_files:
                filename = od[list_index]
                shutil.copy(Main_Directory + '\\OriginalTempImages\\' + filename, Main_Directory + '\\FinalTempImages\\'+str(time_stamp)+'.csv')
                time_stamp_list.append(time_stamp)
                file_number_index+=1
                list_index+=1




'Crops the images to include only the leaf.'
for filename in os.listdir(Main_Directory + 'FinalTempImages'):
    print(filename)
    with open(filename,'r', encoding='utf-8-sig') as T_ecsvfile, open('Leaf Temperature ' + filename, 'w', newline = '') as LTFinal:
        in_file = csv.reader(T_ecsvfile)
        out_file = csv.writer(LTFinal)
        for row_number, row in enumerate(in_file):
            if row_number > 80 and row_number < 414:
                print(row_number)
                out_file.writerow(row[174:568])

workbook = load_workbook(Main_Directory + 'Graph Analysis.xlsx')
ws = workbook.create_sheet(''.join(date))






pixel_list = input('What are the Excel Coordinates you would like to analyze? Separate each coordinate with a space.')
pixel_list = pixel_list.split()



"Constants"

"Latent Heat of Vaporization for Water"
L_w = 40.68
w_0 = 657959000
T_w = 4982.85
R=float(R)

"Calculate Conductance."
def g_s(K_matrix,T_a,T_e):
    return (R + K_matrix*(T_a-T_e))/(L_w*(w_0*np.exp(-T_w/(T_e+273)) - w_a))
"Calculate Air Temperature."
def T_a(columnindex):
    return ((after_average-before_average)/(rowlength-1))*columnindex + before_average



for filename in os.listdir('FinalTempImages'):
    "Find air temperatures from the data_extraction file."
    ubt = data_extraction[filename][0]
    uat = data_extraction[filename][1]
    lbt = data_extraction[filename][2]
    lat = data_extraction[filename][3]
    "Calculate average before and average after temperatures from the thermocouples."
    before_average = (ubt + lbt)/2
    after_average = (uat+lat)/2
    "Find water mole fraction from data extraction."
    w_a = data_extraction[filename][4]
    with open('C:\\Users\\Joshua Hibbard\\Box\\Josh data\\KMatrix\\KMatrix_23C_1Amp.csv','r') as K_matrixcsvfile, open('Leaf Temperature ' + filename,'r', encoding='utf-8-sig') as T_ecsvfile, open(year + '_' + month +'_' + day + '_' + 'Conductance' + filename + '.csv','w') as outputfile:
        K_matrix = csv.reader(K_matrixcsvfile, delimiter = ',', quotechar='\n')
        T_e = csv.reader(T_ecsvfile, delimiter = ',', quotechar='\n')
        while True:
            K_matrixrow=next(K_matrix)
            T_erow=next(T_e)
            rowlength=len(T_erow)
            columnindex=0
            while columnindex < rowlength:
                K_matrixvalue = float(K_matrixrow[columnindex])
                T_avalue = float(T_a(columnindex))
                T_evalue = float(T_erow[columnindex])
                conductance = g_s(K_matrixvalue,T_avalue,T_evalue)
                outputfile.write(str(conductance))
                columnindex+=1
                if columnindex < rowlength:
                    outputfile.write(',')
            outputfile.write('\n')
    with open(year + '_' + month +'_' + day + '_' + 'Conductance' + filename + '.csv','r') as Conductance, open('Leaf Temperature ' + filename,'r', encoding='utf-8-sig') as T_ecsvfile:
        conductance_rows = list(csv.reader(Conductance))
        T_e_rows = list(csv.reader(T_ecsvfile))
        for pixel in pixel_list():
            #Pixel function that turns the Excel values into coordinates based on row and column number
            

                
            
            
            
            
            
