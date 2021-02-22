###############################
###SGW-91###_Protein__Pilot_###
###200319###Report_Extractor###
#####swilliams@cmri.org.au#####
###############################

#import the required packages
import pandas as pd
import os
import glob
import xlrd

#for a single file - set the file name as the variable "list", set the working directory and run from line 38
line = '210127_HEK-QCI_00J8W_0169V_M06_I_1__FDR'

#set the working directory to the location of the files for conversion
cd "C:\User\Data"

#for multiple file extractions, generate a list of xlsx files for the working directory - run the next 7 lines together
extension = 'xlsx'
file = glob.glob('*.{}'.format(extension))
resultFyle = open("xlsx_file_list.txt",'w')
for r in file:
	resultFyle.write(r + "\n")
resultFyle.close()

#read the xlsx file list and run the conversion - for multiple files run all lines to end togther, for a single file, run from line 38
file = open("xlsx_file_list.txt","r")
list = [line.rstrip() for line in file.readlines()]
for line in list:
	workbook = xlrd.open_workbook(line)
	worksheet = workbook.sheet_by_name("Search Summary")
	spec=("Number of spectra in search: {0}".format(worksheet.cell(9,8).value))
	prot=("Protein 1% Global FDR: {0}".format(worksheet.cell(7,4).value))
	pep=("Peptide 1% Global FDR: {0}".format(worksheet.cell(13,4).value))
	PSM=("Peptide Spectrum Match 1% Global FDR: {0}".format(worksheet.cell(19,4).value))
	
	DBS=("{0}".format(worksheet.cell(29,8).value))
	CysAlk=("Cysteine Alkylation: {0}".format(worksheet.cell(14,8).value))
	Speci=("Species: {0}".format(worksheet.cell(18,8).value))
	
	worksheet = workbook.sheet_by_name("Characterize Digestion")
	csp=("Canonical sequence peptides: {0}".format(worksheet.cell(8,14).value))
	osp=("Over-cleaved sequence peptides: {0}".format(worksheet.cell(9,14).value))
	usp=("Under-cleaved sequence peptides: {0}".format(worksheet.cell(10,14).value))
	mou=("Mixed over and under cleaved: {0}".format(worksheet.cell(11,14).value))
	
	worksheet = workbook.sheet_by_name("Most Frequent Deltas")
	MMF1=("1: {0}".format(worksheet.cell(5,9).value))
	MMF101=(": {0}".format(worksheet.cell(5,12).value))
	MMF2=("2: {0}".format(worksheet.cell(6,9).value))
	MMF201=(": {0}".format(worksheet.cell(6,12).value))
	MMF3=("3: {0}".format(worksheet.cell(7,9).value))
	MMF301=(": {0}".format(worksheet.cell(7,12).value))
	MMF4=("4: {0}".format(worksheet.cell(8,9).value))
	MMF401=(": {0}".format(worksheet.cell(8,12).value))
	MMF5=("5: {0}".format(worksheet.cell(9,9).value))
	MMF501=(": {0}".format(worksheet.cell(9,12).value))
	MMF6=("6: {0}".format(worksheet.cell(10,9).value))
	MMF601=(": {0}".format(worksheet.cell(10,12).value))
	MMF7=("7: {0}".format(worksheet.cell(11,9).value))
	MMF701=(": {0}".format(worksheet.cell(11,12).value))
	MMF8=("8: {0}".format(worksheet.cell(12,9).value))
	MMF801=(": {0}".format(worksheet.cell(12,12).value))
	MMF9=("9: {0}".format(worksheet.cell(13,9).value))
	MMF901=(": {0}".format(worksheet.cell(13,12).value))
	MMF10=("10: {0}".format(worksheet.cell(14,9).value))
	MMF1001=(": {0}".format(worksheet.cell(14,12).value))
	MMF11=("11: {0}".format(worksheet.cell(15,9).value))
	MMF111=(": {0}".format(worksheet.cell(15,12).value))
	MMF12=("12: {0}".format(worksheet.cell(16,9).value))
	MMF121=(": {0}".format(worksheet.cell(16,12).value))
	MMF13=("13: {0}".format(worksheet.cell(17,9).value))
	MMF131=(": {0}".format(worksheet.cell(17,12).value))
	MMF14=("14: {0}".format(worksheet.cell(18,9).value))
	MMF141=(": {0}".format(worksheet.cell(18,12).value))
	MMF15=("15: {0}".format(worksheet.cell(19,9).value))
	MMF151=(": {0}".format(worksheet.cell(19,12).value))
	
	worksheet = workbook.sheet_by_name("Spectrum FDR Summary")
	gFDRff=("Global FDR from Fit (%): {0}".format(worksheet.cell(5,19).value*100))
	
	worksheet = workbook.sheet_by_name("Chromatography")
	PWHH=("Chromatographic Peak Width at Half Height (Median in sec): {0}".format(worksheet.cell(6,2).value))
	
	saveFile = open('%s.txt' % line,'w')
	saveFile.write(spec)
	saveFile.write("\n")
	saveFile.write(prot)
	saveFile.write("\n")
	saveFile.write(pep)
	saveFile.write("\n")
	saveFile.write(PSM)
	
	saveFile.write("\n")
	saveFile.write("\n")
	saveFile.write("Digestion Efficiency:")
	saveFile.write("\n")
	saveFile.write(csp)
	saveFile.write("\n")
	saveFile.write(osp)
	saveFile.write("\n")
	saveFile.write(usp)
	saveFile.write("\n")
	saveFile.write(mou)
	
	saveFile.write("\n")
	saveFile.write("\n")
	saveFile.write("Most Frequent Deltas:")
	saveFile.write("\n")
	saveFile.write(MMF1)
	saveFile.write(MMF101)
	saveFile.write("\n")
	saveFile.write(MMF2)
	saveFile.write(MMF201)
	saveFile.write("\n")
	saveFile.write(MMF3)
	saveFile.write(MMF301)
	saveFile.write("\n")
	saveFile.write(MMF4)
	saveFile.write(MMF401)
	saveFile.write("\n")
	saveFile.write(MMF5)
	saveFile.write(MMF501)
	saveFile.write("\n")
	saveFile.write(MMF6)
	saveFile.write(MMF601)
	saveFile.write("\n")
	saveFile.write(MMF7)
	saveFile.write(MMF701)
	saveFile.write("\n")
	saveFile.write(MMF8)
	saveFile.write(MMF801)
	saveFile.write("\n")
	saveFile.write(MMF9)
	saveFile.write(MMF901)
	saveFile.write("\n")
	saveFile.write(MMF10)
	saveFile.write(MMF1001)
	saveFile.write("\n")
	saveFile.write(MMF11)
	saveFile.write(MMF111)
	saveFile.write("\n")
	saveFile.write(MMF12)
	saveFile.write(MMF121)
	saveFile.write("\n")
	saveFile.write(MMF13)
	saveFile.write(MMF131)
	saveFile.write("\n")
	saveFile.write(MMF14)
	saveFile.write(MMF141)
	saveFile.write("\n")
	saveFile.write(MMF15)
	saveFile.write(MMF151)
	
	saveFile.write("\n")
	saveFile.write("\n")
	saveFile.write("Correspondence between FDR Levels and ProteinPilot Reported Confidences:")
	saveFile.write("\n")
	saveFile.write(gFDRff)
	saveFile.write("\n")
	saveFile.write("\n")
	saveFile.write("QC info:")
	saveFile.write("\n")
	saveFile.write(PWHH)
	saveFile.write("\n")
	saveFile.write("\n")
	saveFile.write("Database Filename:")
	saveFile.write("\n")
	saveFile.write(DBS)
	saveFile.write("\n")
	saveFile.write("\n")
	saveFile.write(CysAlk)
	saveFile.write("\n")
	saveFile.write(Speci)

	saveFile.close()

#Export the list of peptides, proteins and modifications - to not export, put "#" before line 197
	worksheet = workbook.sheet_by_name("Flat Protein ID Summary")
	g = worksheet.col_values(2)
	worksheet = workbook.sheet_by_name("Flat Peptide Summary")
	h = worksheet.col_values(5)	
	worksheet = workbook.sheet_by_name("Mod Frequencies")
	c = worksheet.col_values(7) 
	d = worksheet.col_values(11) 
	e = worksheet.col_values(9) 
	f = worksheet.col_values(13) 

	pd.DataFrame([h,g,c, d, e, f]).T.to_csv("%s_pep-prot-mod_list.csv" % line)
	
	
file.close()