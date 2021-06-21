#-------------------------------------------------------------------------------
# Name:        GS_FIM_metadata 
# Purpose:     This script parses FIM related information (e.g. prov, city, year
#              volume, etc..) from the Geosearch folder inventory and exports the 
#              data by state to excel to create a master list of FIMs.
#              The script also creates a "questionable sheet" to note directories that
#              do not meet the pattern specifid in the script. Review the script 
#              and make adjustments if needed. 
#
# Author:      czhou
#
# Created:     06/18/2021
# Copyright:   (c) czhou 2021
# Licence:     <your licence>
#-------------------------------------------------------------------------------

def desc_decode(code):
    # Decode Geosearch codes to describe condition of FIM
    if " rev" in code:
        description = code.replace(" rev", "Revised")
        description = re.findall(r"\D+", description)
    elif " c" in code: 
        description = code.replace(" c", "Colour")
        description = re.findall(r"\D+", description)
    elif "rep" in code:
        description = code.replace(" rep", "Reprinted")
        description = re.findall(r"\D+", description)
    else:
        description = None
    return description


import os
import urllib
import re
import win32com.client as win32


prov = "Hawaii"    #Change prov 
gs_path = (r"\\10.6.246.73\Sanborn")
path = os.path.join(gs_path, prov)

#Loop through path to obtain list of filepaths
file_paths = [] 
for root, directories, files in os.walk(path):
    files[:] = [f for f in files if f not in ["Thumbs.db"]]
    for item in files:
        filepath = os.path.join(root, item)
        file_paths.append(filepath)


#Loop through the list of filepaths to parse metadata and append data to a dictionary
unique_list = set()          # Unique list of file paths 
resultlist = []              # List of parsed filepaths to export to excel
resultlist_questionable = [] #List of questionable paths to be manually verified

for item in file_paths:
    delimiter_count = item.count("\\")
    if delimiter_count == 9:
        file_path = item                   #e.g. \\10.6.246.73\Sanborn\Alaska\Anchorage\1916\AK_Anchorage_1916_V1_S1.jpg
        dir_path = item.rsplit("\\", 1)[0] #e.g. \\10.6.246.73\Sanborn\Alaska\Anchorage\1916
        prov = item.split("\\",5)[4]
        city = item.split("\\",6)[5]
        year = item.split("\\",7)[6]
        vol  = item.split("\\",8)[7]
        vol_year = vol  = item.split("\\",9)[8]
        description = desc_decode(year)
        map_file = item.rsplit("\\",1)[1]  #e.g     AK_Anchorage_1916_V1_S1.jpg

        if file_path not in unique_list: 
            unique_list.add(file_path)
            resultlist.append({"prov":prov, "city":city, "year":year, "dir_path": dir_path, "desc": description, 
            "file": map_file, "file_path": file_path, "vol":vol, "vol_year": vol_year}) 

    elif delimiter_count == 8:
        file_path = item                   
        dir_path = item.rsplit("\\", 1)[0] 
        prov = item.split("\\",5)[4]
        city = item.split("\\",6)[5]
        year = item.split("\\",7)[6]
        vol  = item.split("\\",8)[7]
        vol_year = ""
        description = desc_decode(year)
        map_file = item.rsplit("\\",1)[1]  

        if file_path not in unique_list: 
            unique_list.add(file_path)
            resultlist.append({"prov":prov, "city":city, "year":year, "dir_path": dir_path, "desc": description, 
            "file": map_file, "file_path": file_path, "vol":vol, "vol_year": ""}) 
    
    elif delimiter_count == 7:
        file_path = item                   
        dir_path = item.rsplit("\\", 1)[0] 
        prov = item.split("\\",5)[4]
        city = item.split("\\",6)[5]
        year = item.split("\\",7)[6]
        vol  = ""
        vol_year = ""
        description = desc_decode(year)
        map_file = item.rsplit("\\",1)[1]  

        if file_path not in unique_list: 
            unique_list.add(file_path)
            resultlist.append({"prov":prov, "city":city, "year":year, "dir_path": dir_path, "desc": description, 
            "file": map_file, "file_path": file_path, "vol":vol, "vol_year": ""})

    elif delimiter_count == 6:
        dir_path = item.rsplit("\\", 1)[0]
        prov = item.split("\\",5)[4]
        city = item.rsplit("\\",6)[5]
        year = desc_decode(city)
        vol  = ""
        vol_year = ""
        description = ""
        map_file = item.rsplit("\\",1)[1]

        if file_path not in unique_list: 
            unique_list.add(file_path)
            resultlist.append({"prov":prov, "city":city, "year":year, "dir_path": dir_path, "desc": description, 
            "file": map_file, "file_path": file_path, "vol":vol, "vol_year": ""})
    
    #Some FIMS might not have a regular pattern, set these paths to a questionable list for manual verification
    else:
        dir_path = item

        if dir_path not in unique_list: 
            unique_list.add(dir_path)
            resultlist_questionable.append({"dir_path": dir_path})
           

# Export data to excel 
# Get application object
excel = win32.Dispatch("Excel.Application")
excel.Visible = 1    
if os.path.exists(r"C:\Users\czhou\Documents\GEOSEARCH\FIM\Geosearch_Fim.xlsx"):
    excel = excel.Workbooks.Open(r"C:\Users\czhou\Documents\GEOSEARCH\FIM\Geosearch_Fim.xlsx")   
else: 
    excel.Workbooks.Add()  
    excel.ActiveWorkbook.SaveAs(Filename=r"C:\Users\czhou\Documents\GEOSEARCH\FIM\Geosearch_Fim.xlsx")      

#Create new worksheets per state

sheet1 = excel.Sheets.Add(Before = None , After = excel.Sheets(excel.Sheets.count))
sheet1.Name = prov
sheet2 = excel.Sheets.Add(Before = None , After = excel.Sheets(excel.Sheets.count))
sheet2.Name = prov + " - Questionable"
                                           
# Add table headers to Sheet1
sheet1.Cells(1,1).Value = "Province"
sheet1.Cells(1,2).Value = "City"                                                    
sheet1.Cells(1,3).Value = "Year"  
sheet1.Cells(1,4).Value = "Volume"         
sheet1.Cells(1,5).Value = "Vol Year"                                         
sheet1.Cells(1,6).Value = "Map file"                                             
sheet1.Cells(1,7).Value = "Description"
sheet1.Cells(1,8).Value = "Path"
sheet1.Cells(1,9).Value = "Missing year"
sheet1.Cells(1,10).Value = "Missing sheet"
sheet1.Cells(1,11).Value = "Colour"
sheet1.Cells(1,12).Value = "GS better quality"
sheet1.Cells(1,13).Value = "Additional Comments"
    
# Loop through each dictionary item in resultlist to put data into excel
row = 2
for item in resultlist:                                                    
    sheet1.Cells(row,1).Value = item["prov"]                                 
    sheet1.Cells(row,2).Value = item["city"]                                              
    sheet1.Cells(row,3).Value = item["year"]   
    sheet1.Cells(row,4).Value = item["vol"]     
    sheet1.Cells(row,5).Value = item["vol_year"]                                          
    sheet1.Cells(row,6).Value = item["file"]
    sheet1.Cells(row,7).Value = item["desc"]
    sheet1.Cells(row,8).Value = item["dir_path"]
    row +=1

#Format excel sheet
#Autofit field length
sheet1.Columns.AutoFit()

#Left justify fields
sheet1.Cells.HorizontalAlignment = -4131

#Bold headers
i=1
for i in range(1,14):
    sheet1.Cells(1,i).Font.Bold = True
    i += 1
   
# Add table headers to Sheet2 (Questionable)
sheet2.Cells(1,1).Value = "Path"

# Loop through each dictionary item in resultlist_questionable to put data into excel
row = 2
for item in resultlist_questionable:                                                    
    sheet2.Cells(row,1).Value = item["dir_path"]                                 
    row +=1                                                             

#Format excel sheet
#Autofit field length
sheet2.Columns.AutoFit()

# Bold header names
i=1
for i in range(1,2):
    sheet2.Cells(1,i).Font.Bold = True
    i += 1

#Save excel file
#excel.ActiveWorkbook.Save()

