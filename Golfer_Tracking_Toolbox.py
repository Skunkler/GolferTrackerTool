# Name KML_UID
# Purpose: to pull unique identifier from file name and turn into a long

# #################### TO DO #############################################
# 3. Create golf course name in spreadsheets so we can write all features to one feature class
# ########################################################################


import arcpy, os, sys, string, datetime, time, shutil, traceback, xlrd, xlwt
from arcpy import env
from datetime import time
from datetime import datetime


#user defined inputs
kmlDirectory =  arcpy.GetParameterAsText(0)          


excelgolferdata =  arcpy.GetParameterAsText(1)          

#global variables
WorkingGDB =  r"\\storage\snwa\conservation\turf_analysis\Golfer_Data\Working\WorkingGDB.gdb"      


final_fc =  r'\\storage\snwa\conservation\turf_analysis\Golfer_Data\Golf_Course_Conservation.gdb\Golfer_Tracking_Data'         


working_directory = r'\\storage\snwa\conservation\turf_analysis\Golfer_Data\Working'   


jointable = WorkingGDB + "\\golfer_jointable"

logfile = r"\\storage\snwa\conservation\turf_analysis\Golfer_Data\logfile\logfile.log"



#establishes the log file connection
outfile = open(logfile, 'a')

outfile.write('\n' + kmlDirectory + '\n' + 'Golfer_Tracking_Tool' + " ------------------------------------------" '\n')
outfile.close()

timeYearMonDay = datetime.now()
timeHour = timeYearMonDay.hour
timeMin = timeYearMonDay.minute


#create xls book object for working xls
book = xlwt.Workbook()
sheet = book.add_sheet('temp_sheet')
cols = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "K"]

sheet.write(0, 0, u'Unique ID')
sheet.write(0, 1, u'Date')
sheet.write(0, 2, u'Gender')
sheet.write(0, 3, u'Age')
sheet.write(0, 4, u'Tee Time')
sheet.write(0, 5, u'Handicap')
sheet.write(0, 6, u'Holes Played')
sheet.write(0, 7, u'Distance (mi)')
sheet.write(0, 8, u'Duration')
sheet.write(0, 9, u'Logger')
sheet.write(0, 10, u'Golf_Course_Name')


#Takes the working xls that has clean records and unique records and converts them
#to a ArcGIS table
def convertExcelToTable(tempTable):
    if arcpy.Exists(jointable):
        arcpy.Delete_management(jointable)
        arcpy.AddMessage('%s deleted!' % jointable)
        arcpy.ExcelToTable_conversion(tempTable,jointable)
    else:
        arcpy.ExcelToTable_conversion(tempTable,jointable)
    
    
#This converts the python date time to excel serial date
#That way ArcMap can read the date correctly
#
def convertSerialDate(sh_cell_val):
    try:
        dt = datetime.fromordinal(datetime(1900,1,1).toordinal() + int(sh_cell_val) - 2)
        dt_str = str(dt)
        year = dt_str.split('-')[0][-2:]
        month = dt_str.split('-')[1]
        day = dt_str.split('-')[2].split(' ')[0]
        CSV_Date = month + '/' + day + '/' + year
        return sh_cell_val
    except:
        return False

#checks the authenticity of duration time
#returns false if bad value exists in cell
def convertTime(sh_cell_val):
    try:
        x = sh_cell_val
        x = int(x*24*3600)
        csv_time = time(x//3600, (x%3600)//60, x%60)
        return sh_cell_val
    except:
        return False

#writes the output to the working excel if the line does not
#have any false values in it
def writeTempXls(line, rowVal):
    
    unique_ID = line[0]
    DateVar = line[1]
    Gender = line[2]
    Age = line[3]
    teeTime = line[5]
    handicap = line[4]
    HolesPlayed = line[6]
    Distance = line[7]
    Duration = line[8]
    Logger = line[9]
    GolfCourse = line[10]
    sheet.write(rowVal, 0, unique_ID)

    date_format = xlwt.XFStyle()
    date_format.num_format_str = 'mm/dd/yy'
    Dur_time_format = xlwt.XFStyle()
    Dur_time_format.num_format_str = 'h:mm:ss'
    timeDay_format = xlwt.XFStyle()
    timeDay_format.num_format_str = 'hh:mm AM/PM'
    
    sheet.write(rowVal, 1, DateVar, date_format)
    sheet.write(rowVal, 2, Gender)
    sheet.write(rowVal, 3, Age)
    sheet.write(rowVal, 4, teeTime, timeDay_format)
    sheet.write(rowVal, 5, handicap)
    sheet.write(rowVal, 6, HolesPlayed)
    sheet.write(rowVal, 7, Distance)
    sheet.write(rowVal, 8, Duration, Dur_time_format)
    sheet.write(rowVal, 9, Logger)
    sheet.write(rowVal, 10, GolfCourse)
    



#read the golferexcel, tests each value in each cell per row, directs
#good values to written to working excel
#skips the bad values and notifies the user of which rows have errors
def checkExcel():
    with xlrd.open_workbook(excelgolferdata) as wb:
        sh = wb.sheet_by_index(0)
    
        for i in range(1, sh.nrows):
            line = [sh.cell_value(i,0), convertSerialDate(sh.cell_value(i,1)), sh.cell_value(i,2), sh.cell_value(i,3), sh.cell_value(i,4), convertTime(sh.cell_value(i,5)), \
            sh.cell_value(i,6), sh.cell_value(i,7), convertTime(sh.cell_value(i,8)), sh.cell_value(i,9), sh.cell_value(i,10)]
            if False not in line:
                writeTempXls(line, i)
            else:
                arcpy.AddMessage("bad value with record %s" % str(line[0]))
                outfile = open(logfile, 'a')
                outfile.write("Process: Failed for: " + excelgolferdata  + " " + str(timeYearMonDay) + " " + str(timeHour) + ":" + str(timeMin) + '\n')
                outfile.write("bad value with record " + str(line[0]))
                outfile.write(str(line))
                outfile.close()


                
    book.save(r'\\storage\snwa\conservation\turf_analysis\Golfer_Data\Working\Working.xls')
    tempTable = r'\\storage\snwa\conservation\turf_analysis\Golfer_Data\Working\Working.xls'
    convertExcelToTable(tempTable)











#converts the input kmls in the kml directory to layers in seperate fgdbs if the kml has a unique name
#otherwise the kml is skipped
def GetKmlToFc(kml):
    FileGeoDB = working_directory + "\\" + kml[:-4] + ".gdb"
    if arcpy.Exists(FileGeoDB):
        try:
            arcpy.Delete_management(FileGeoDB)
            arcpy.AddMessage('%s deleted!' % FileGeoDB)
            arcpy.KMLToLayer_conversion(kml, working_directory, kml[:-4],"NO_GROUNDOVERLAY")
        except:
            arcpy.AddMessage("error deleting FDs")
            outfile = open(logfile, 'a')
            outfile.write("Process: Failed for: " + kml + " " + str(timeYearMonDay) + " " + str(timeHour) + ":" + str(timeMin) + '\n')
            outfile.write(arcpy.GetMessages(2))
            
    else:
        try:
            arcpy.AddMessage('%s does not exist therefore it was not deleted!' % FileGeoDB)
            arcpy.KMLToLayer_conversion(kml, working_directory, kml[:-4],"NO_GROUNDOVERLAY")
        except:
            arcpy.AddMessage("Error with kml to layer")
            outfile = open(logfile, 'a')
            outfile.write("Process: Failed for: " + kml + " " + str(timeYearMonDay) + " " + str(timeHour) + ":" + str(timeMin) + '\n')
            outfile.write(arcpy.GetMessages(2))
    
#this function deletes any leftover data from the working directory
#checks if the name for the input kmls are unique, if they are unique, convert kml to layer
#then run the layer to fc function at the end
def initialCheck():

    
    arcpy.env.workspace = working_directory
    wks = arcpy.ListWorkspaces('*', 'FileGDB')

    wks.remove(WorkingGDB)
    arcpy.AddMessage("Cleaning up working directory, one moment please")
    for fgdb in wks:
        if arcpy.Exists(fgdb):
            arcpy.AddMessage("deleting %s" %fgdb)
            arcpy.Delete_management(fgdb)

    files = arcpy.ListFiles("*.lyr")

    for lyr in files:
        if arcpy.Exists(lyr):
            arcpy.AddMessage("deleting %s" %lyr)
            arcpy.Delete_management(lyr)

            
    UID_List =[]
    with arcpy.da.SearchCursor(final_fc, ['UniqueID']) as cursor:
        for row in cursor:
            UID_List.append(str(row[0]) + '.kml')
    

    arcpy.env.workspace = kmlDirectory
    arcpy.env.overwriteOutput = True

    
    for kml in arcpy.ListFiles("*.kml"):
        kmlFile = str(kml)
        if kmlFile not in UID_List:
            arcpy.AddMessage('%s is a unique name' % kml[:-4])
            GetKmlToFc(kml)
        
       
        else:
             arcpy.AddMessage('%s needs a unique name, this kml file will not be processed!' % kml[:-4])
            
            
    getLayerToFC()
            
            
#taking each feature layer and combining everything as seperate feature classes in the same fgdb
#then run the final join

def getLayerToFC():
    arcpy.env.workspace = working_directory
    wks = arcpy.ListWorkspaces('*', 'FileGDB')
    
        
    
    for fgdb in wks:  
        try:
            arcpy.env.workspace = fgdb
            featureClasses = arcpy.ListFeatureClasses('*', '', 'Placemarks')
            for fc in featureClasses:
           
            
                arcpy.AddMessage('COPYING from %s' %fgdb)
                fcCopy = fgdb + os.sep + 'Placemarks' + os.sep + fc 
          
    
                fc_output = WorkingGDB + '\\' + fc + '_' + fgdb[fgdb.rfind(os.sep)+1:-4]
                if arcpy.Exists(fc_output):
                    arcpy.Delete_management(fc_output)
                    arcpy.AddMessage('%s feature class deleted!' % fc_output)
                if fc == 'Polylines':
                    arcpy.FeatureClassToFeatureClass_conversion(fcCopy, WorkingGDB, fc + '_' + fgdb[fgdb.rfind(os.sep)+1:-4])
                    arcpy.AddField_management(WorkingGDB + '\\' + fc + '_' + fgdb[fgdb.rfind(os.sep)+1:-4], 'PlaceHoldID', 'TEXT')
                    expression = fgdb.split('\\')[-1][:-4]
                    arcpy.CalculateField_management(WorkingGDB + '\\' + fc + '_' + fgdb[fgdb.rfind(os.sep)+1:-4], 'PlaceHoldID', '"' + expression + '"', 'PYTHON')
        except:
            outfile = open(logfile, 'a')
            outfile.write('error with FeatureClassToFeatureClass_conversion for ' + fgdb + " " + str(timeYearMonDay) + " " + str(timeHour) + ":" + str(timeMin) + '\n')
            outfile.write(arcpy.GetMessages(2) + '\n')



    FinalJoinData()
        
        
#takes each feature class in the working.gdb, drops unnecessary fields, creates an Id field to join with the input Table,
#adds the join to the input table, appends the result to the final fc, deletes any null value records form the final fc

def FinalJoinData():

    arcpy.env.workspace = WorkingGDB

    arcpy.env.overwriteOutput = True
    fcs = arcpy.ListFeatureClasses()

    code_block="""def getID(inVal):
      return int(inVal)"""


    for fc in fcs:
        try:
            arcpy.AddField_management(fc, "UniqueID", "LONG")
            arcpy.CalculateField_management(fc, "UniqueID", "getID(!PLaceHoldID!)", "PYTHON_9.3", code_block)
            dropfields = "FolderPath;SymbolID;AltMode;Base;Clamped;Extruded;Snippet;PopupInfo;PlaceHoldID"
            arcpy.DeleteField_management(fc, dropfields)

            arcpy.MakeFeatureLayer_management(fc,"Lyr")

        
            arcpy.AddMessage("Joining %s to golfer excel table" %fc) 
            arcpy.JoinField_management("Lyr", "UniqueID", jointable, "Unique_ID")
            arcpy.Append_management("Lyr", final_fc, "NO_TEST")
            arcpy.env.workspace = working_directory
            arcpy.env.workspace = WorkingGDB
            arcpy.Delete_management(fc)
        except:
            arcpy.AddMessage("Error with joining data " + fc + '\n')
            outfile = open(logfile, 'a')
            outfile.write("Error with joining data " + fc + " " + str(timeYearMonDay) + " " + str(timeHour) + ":" + str(timeMin) + '\n')
            outfile.write(arcpy.GetMessages(2) + '\n')

    arcpy.AddMessage("deleting null values from final feature class")
    arcpy.MakeFeatureLayer_management(final_fc, "lyr")
    arcpy.SelectLayerByAttribute_management("lyr", "NEW_SELECTION", "Unique_ID IS NULL")
    arcpy.DeleteFeatures_management("lyr")


    
checkExcel()
initialCheck()


outfile.close()

timeYearMonDay = datetime.now()
timeHour = timeYearMonDay.hour
timeMin = timeYearMonDay.minute



print "Process done! " + str(timeYearMonDay) +  " " + str(timeHour)+ ":" + str(timeMin)
outfile= open(logfile,'a')
outfile.write("Process Complete "  + str(timeYearMonDay) +  " " + str(timeHour)+ ":"   + str(timeMin) +  '\n' )
outfile.close() 

    
    

        


