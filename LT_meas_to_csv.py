## Copyright (c) 2015 Daniel O'Neill
## 
##
## Permission is hereby granted, free of charge, to any person obtaining a copy
## of this software and associated documentation files (the "Software"), to 
## deal in the Software without restriction, including without limitation the
## rights to use, copy, modify, merge, publish, distribute, sublicense, and/or
## sell copies of the Software, and to permit persons to whom the Software is
## furnished to do so, subject to the following conditions:
##
## The above copyright notice and this permission notice shall be included in
## all copies or substantial portions of the Software.
##
## THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
## IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
## FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
## AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
## LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
## FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS
## IN THE SOFTWARE.


import time, sys, traceback, os, tempfile
from xlsxwriter.workbook import Workbook
from optparse import OptionParser

#Process a logfile for step and measurement data
def processFile(inFileName):
  infile = open(inFileName, 'r')
  lines = infile.readlines()
  infile.close()

  dataList = []
  headerDone = False

  for line in lines:
    #process header line
    if ".step" in line:
      if not headerDone:
        hdrIn = line.split()[1:]
        hdrOut = []
        for d in hdrIn:
          hdrOut.append(d.split("=")[0])
        dataList.append(hdrOut)
        headerDone = True
      
      datIn = line.split()[1:]
      datOut = []
      for d in datIn:
        datOut.append(d.split("=")[1])
      dataList.append(datOut)

  #find measurement data
  measurementStartLines = []
  for i in range(len(lines)):
    if "Measurement" in lines[i]:
      if "FAIL'ed" not in lines[i]:
        measurementStartLines.append(i)
      else:
        print "WARNING: Failed measurement: %s" % lines[i].split(" ")[1].strip()
        
  #add headers header column
  for startLine in measurementStartLines:
    dataList[0].append(lines[startLine].split(" ")[1].strip())
    for i in range(len(dataList)-1):
      dataList[i+1].append(lines[startLine + 2 + i].split()[1])

  return dataList
  
  
#dump data to csv file  
def dataToCSV(dataList, outFileName):
  outfile = open(outFileName, 'w')
  for a in dataList:
    writeString = ""
    for b in a:
      outfile.write(b + ',')
    outfile.write('\n')
  outfile.close()

  
#dump data to xlsx file
def dataToXLSX(dataList, outFileName):
  workbook = Workbook(outFileName)
  worksheetData = workbook.add_worksheet('Data')
  
  #Can we make this mapping automatic?
  columnMap = {0:"A", 1:"B", 2:"C", 3:"D", 4:"E", 5:"F", 6:"G", 7:"H", 8:"I", 9:"J", 10:"K", 11:"L",
                12:"M", 13:"N", 14:"O", 15:"P"}
  bold = workbook.add_format({'bold': 1})

  # Format the worksheet1 data cells that the chart will refer to
  headings = dataList[0]
    
  # Set column width. Note that it is not possible to "auto-fit" as this only happens at run-time in Excel
  worksheetData.set_column('A:%s' % columnMap[len(dataList[0])], 15)
  
  # Write headings row to workbook
  worksheetData.write_row('A1', headings, bold) # note that alignment not applicable to the write_row method
  
  #write data
  rowCount = 2
  for data in dataList[1:]:
    columnCount = 0
    for item in data:
      worksheetData.write('%s%i' % (columnMap[columnCount], rowCount), float(item))
      columnCount += 1
    rowCount += 1

  #add an auto filter and freeze pane on the headers row
  worksheetData.autofilter('A1:%s%d' % (columnMap[len(dataList[0])-1], len(dataList)))
  worksheetData.freeze_panes(1, 0)
  workbook.close()

  
#Identify the latest log in a specific directory
def getLogFromDirectory(directoryName):
  inFileName = ""
  fileList = os.listdir(directoryName)
  timeAndFileList = []
  for i in range(len(fileList)):
    if ".log" in fileList[i]:
      timeAndFileList.append((os.path.getmtime(directoryName + '\\' + fileList[i]), directoryName + '\\' + fileList[i]))
  
  timeAndFileList.sort(key=lambda item: item[0], reverse=True)
  
  for modifiedTime, fileName in timeAndFileList:
    #print fileName
    fh = file(fileName)
    firstLine = fh.readline()
    if ".asc" in firstLine and "Circuit" in firstLine:
      inFileName = fileName
      break
  return inFileName
  
  
if __name__ == "__main__":
  usage = "usage: %prog [options] arg1 arg2"
  parser = OptionParser(usage=usage)
  parser.add_option("-t", "--type", dest="outFileType", default="xlsx",
                      help="type of output (xlsx, csv etc)")

  (options, args) = parser.parse_args()
  if len(args) == 0:
    print "No argument supplied assuming latest log file in temp directory"
    TEMP_DIR = tempfile.gettempdir()

    #get log from temp
    print "Temp Dir:", TEMP_DIR
    inFileName = getLogFromDirectory(TEMP_DIR)
    
    if inFileName == "":
      print "no log file"
      SystemExit
    print "File to process:" + inFileName
    
    
  else:
    print args
    inFileName = r"C:\Users\oneillda\AppData\Local\Temp\Soft_start_testbench.log"


  outFileType = options.outFileType
  outFileName = "{0}_{1}.{2}".format(inFileName.split('\\')[-1].split('.')[0], time.strftime('%Y-%m-%d_%H-%M-%S'), outFileType)
  
  try:
    #use input file to create dataList
    dataList = processFile(inFileName)
    
    #write data to output file
    if outFileName[-4:] == "xlsx":
      dataToXLSX(dataList, outFileName)
    else:
      dataToCSV(dataList, outFileName)
  
  #catch exceptions
  except:
    print "Fail"
    ex_type, ex, tb = sys.exc_info()
    traceback.print_tb(tb)
    print ex_type, ex
  finally:
    print "Done"
    time.sleep(2)
