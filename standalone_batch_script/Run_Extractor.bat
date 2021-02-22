cd %root% 

dir /s /b *.xlsx | findstr /E ".xlsx" > xlsx_file_list.txt

python ./Data_Extractor_Script.bat