import time, sys, traceback, os, tempfile
from xlsxwriter.workbook import Workbook
from optparse import OptionParser

def processFile(inFileName):

  infile = open(inFileName, 'r')
  lines = infile.readlines()
  infile.close()

  dataList = []
  headerDone = False

  for line in lines:
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
  
  
  
def dataToCSV(dataList, outFileName):
  outfile = open(outFileName, 'w')
  for a in dataList:
    writeString = ""
    for b in a:
      outfile.write(b + ',')
    outfile.write('\n')
  outfile.close()

def dataToXLSX(dataList, outFileName):

  workbook = Workbook(outFileName)
  worksheetData = workbook.add_worksheet('Data')
  
  columnMap = {0:"A", 1:"B", 2:"C", 3:"D", 4:"E", 5:"F", 6:"G", 7:"H", 8:"I", 9:"J", 10:"K"}
  bold = workbook.add_format({'bold': 1})

  # Format the worksheet1 data cells that the chart will refer to
  headings = dataList[0]
  
  # Set number format "constructor"
  #formatcells1 = workbook.add_format()
  #formatcells2 = workbook.add_format()
  
  # Format to 3 decimal places
  #formatcells1.set_num_format('0.000')
  
  # Columns A to F width set to 25. Note that it is not possible to "auto-fit" as this only happens at run-time in Excel
  worksheetData.set_column('A:%s' % columnMap[len(dataList[0])], 15)
  #worksheetData.set_column('B:Q', 18)
  
  # Set horizontal and vertical alignment
  #formatcells1.set_align('center') 
  #formatcells1.set_align('vcenter')
  #formatcells2.set_align('center') 
  #formatcells2.set_align('vcenter')
  
  # Write headings row to workbook
  worksheetData.write_row('A1', headings, bold) # note that alignment not applicable to the write_row method
  
  
  ####repeat lots
  rowCount = 2
  for data in dataList[1:]:
    columnCount = 0
    for item in data:
      worksheetData.write('%s%i' % (columnMap[columnCount], rowCount), float(item))
      columnCount += 1
    rowCount += 1
      
    #worksheetData.write('D%i' % (rowCount + 1), float(configuredInputVolts), formatcells1)
    #self.worksheetData.write('A%i' % (rowCount), time.ctime(), self.formatcells2)
  
  worksheetData.autofilter('A1:%s%d' % (columnMap[len(dataList[0])-1], len(dataList)))
  worksheetData.freeze_panes(1, 0)
  workbook.close()

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
