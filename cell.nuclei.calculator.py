import numpy as np
import xlsxwriter
import os

machine_name = input("Enter machine name: ")
tissue_type = input("Enter tissue type: ")
date = input("Enter date as MM.DD.YY: ")
weight = input("Enter tissue weight (in mg): ")
weight = float(weight)
k2_check = input("If K2 readings are available, type 'yes', else type 'no': ")

titer_list = []
titer1 = input("Enter first titer value: ")
titer1 = float(titer1)
titer_list.append(titer1)
titer2 = input("Enter second titer value: ")
titer2 = float(titer2)
titer_list.append(titer2)
titer3 = input("Enter third titer value: ")
titer3 = float(titer3)
titer_list.append(titer3)
titer4 = input("Enter fourth titer value: ")
titer4 = float(titer4)
titer_list.append(titer4)
titer_array = np.array(titer_list, dtype = 'f') #change list of titers into an array in order to more easily multiply by dilution factor

dilution = input("Enter dilution factor, if none enter 'none': ")
if dilution != 'none':
    dilution = float(dilution)
    titer1df = titer1 * dilution
    titer2df = titer2 * dilution
    titer3df = titer3 * dilution
    titer4df = titer4 * dilution
    titer_with_df_array = titer_array * dilution
    titer_sum = np.sum(titer_with_df_array)
    titer_average = (titer_sum/4)
    avg_titer_per_mg = (titer_average/weight)
    std_dev_titer = np.std(titer_with_df_array, ddof = 1)
    std_dev_per_mg = std_dev_titer/ weight
        
else:
    dilution = str(dilution)
    titer_sum = np.sum(titer_array)
    titer_average = (titer_sum/4)
    avg_titer_per_mg = (titer_average/weight)
    std_dev_titer = np.std(titer_array, ddof = 1)
    std_dev_per_mg = std_dev_titer/ weight

    
 
if k2_check == 'yes':
        k2_list = []
        k2_first_reading = input("Enter first K2 reading: ")
        k2_first_reading = float(k2_first_reading)
        k2_list.append(k2_first_reading)
        k2_second_reading = input("Enter second K2 reading: ")
        k2_second_reading = float(k2_second_reading)
        k2_list.append(k2_second_reading)
        k2_array = np.array(k2_list, dtype='f')
        k2_sum = np.sum(k2_array)
        k2_average = (k2_sum/2)
        avg_k2_per_mg = (k2_average/weight)

viability_list = []
k2_viability_list = []
viability1 = input("Enter first viability value (as a percentage): ")
viability1 = int(viability1)
viability_list.append(viability1)
viability2 = input("Enter second viability value (as a percentage): ")
viability2 = int(viability2)
viability_list.append(viability2)
viability3 = input("Enter third viability value (as a percentage): ")
viability3 = int(viability3)
viability_list.append(viability3)
viability4 = input("Enter fourth viability value (as a percentage): ")
viability4 = int(viability4)
viability_list.append(viability4)
viability_array = np.array(viability_list)
viability_sum = np.sum(viability_array)
viability_average = (viability_sum/4)
std_dev_viability = np.std(viability_array, ddof = 1) #set ddof to 1, in order to calculate sample std dev

if k2_check == 'yes':
    k2_viability1 = input("Enter first K2 viability value (as a percentage): ")
    k2_viability1 = int(k2_viability1)
    k2_viability_list.append(k2_viability1)
    k2_viability2 = input("Enter second K2 viability value (as a percentage): ")
    k2_viability2 = int(k2_viability2)
    k2_viability_list.append(k2_viability2)
    k2_viability_array = np.array(k2_viability_list)
    k2_viability_sum = np.sum(k2_array)
    k2_viability_average = (k2_viability_sum/2)
    std_dev_k2_viability = np.std(k2_viability_array, ddof =1)


workbook = xlsxwriter.Workbook(os.path.join(os.path.dirname(os.path.abspath(__file__)),machine_name+tissue_type+date+"run.xlsx"))

worksheet = workbook.add_worksheet()
worksheet.write(0, 0, tissue_type)
worksheet.write(1, 0, machine_name)
worksheet.write(0, 1, "Date")
worksheet.write(1, 1, date)
worksheet.write(0, 2, "Weight(mg)")
worksheet.write(1, 2, weight)


if dilution != 'none':
    worksheet.write(0, 3, "Titer (before df)")
    worksheet.write(1, 3, titer1)
    worksheet.write(2, 3, titer2)
    worksheet.write(3, 3, titer3)
    worksheet.write(4, 3, titer4)
    worksheet.write(0, 4, "Dilution Factor")
    worksheet.write(1,4, dilution)
    worksheet.write(0, 5, "Titer (after df)")
    worksheet.write(1, 5, titer1df)
    worksheet.write(2, 5, titer2df)
    worksheet.write(3, 5, titer3df)
    worksheet.write(4, 5, titer4df)
    worksheet.write(0, 6, "Titer Calculations")
    worksheet.write(1, 6, titer_average)
    worksheet.write(2, 6, avg_titer_per_mg)
    worksheet.write(3, 6, std_dev_titer)
    worksheet.write(4, 6, std_dev_per_mg)
    worksheet.write(0, 7, "Viability(%)")
    worksheet.write(1, 7, viability1)
    worksheet.write(2, 7, viability2)
    worksheet.write(3, 7, viability3)
    worksheet.write(4, 7, viability4)
    worksheet.write(0, 8, "Viability Calculations")
    worksheet.write(1, 8, viability_average)
    worksheet.write(3, 8, std_dev_viability)
    worksheet.write(0, 9, "Notes")
    if k2_check == 'yes':
        worksheet.write(0, 10, "K2")
        worksheet.write(0, 11, "Titer")
        worksheet.write(1, 11, k2_first_reading)
        worksheet.write(2, 11, k2_second_reading)
        worksheet.write(0, 12, "Titer Calculations")
        worksheet.write(1, 12, k2_average)
        worksheet.write(2, 12, avg_k2_per_mg)
        worksheet.write(0, 13, "Viability(%)")
        worksheet.write(1, 13, k2_viability1)
        worksheet.write(2, 13, k2_viability2)
        worksheet.write(0, 14, "Viability Calculations")
        worksheet.write(1, 14, k2_viability_average)
        worksheet.write(2, 14, std_dev_k2_viability)
        worksheet.write(0, 15, "Notes")

else:
        worksheet.write(0, 3, "Titer")
        worksheet.write(1, 3, titer1)
        worksheet.write(2, 3, titer2)
        worksheet.write(3, 3, titer3)
        worksheet.write(4, 3, titer4)
        worksheet.write(0, 4, "Titer Calculations")
        worksheet.write(1, 4, titer_average)
        worksheet.write(2, 4, avg_titer_per_mg)
        worksheet.write(3, 4, std_dev_titer)
        worksheet.write(4, 4, std_dev_per_mg)
        worksheet.write(0, 5, "Viability(%)")
        worksheet.write(1, 5, viability1)
        worksheet.write(2, 5, viability2)
        worksheet.write(3, 5, viability3)
        worksheet.write(4, 5, viability4)
        worksheet.write(0, 6, "Viability Calculations")
        worksheet.write(1, 6, viability_average)
        worksheet.write(3, 6, std_dev_viability)
        worksheet.write(0, 7, "Notes")
        if k2_check == 'yes':
            worksheet.write(0, 8, "K2")
            worksheet.write(0, 9, "Titer")
            worksheet.write(1, 9, k2_first_reading)
            worksheet.write(2, 9, k2_second_reading)
            worksheet.write(0, 10, "Titer Calculations")
            worksheet.write(1, 10, k2_average)
            worksheet.write(2, 10, avg_k2_per_mg)
            worksheet.write(0, 11, "Viability(%)")
            worksheet.write(1, 11, k2_viability1)
            worksheet.write(2, 11, k2_viability2)
            worksheet.write(0, 12, "Viability Calculations")
            worksheet.write(1, 12, k2_viability_average)
            worksheet.write(2, 12, std_dev_k2_viability)
            worksheet.write(0, 13, "Notes")

        

workbook.close()

