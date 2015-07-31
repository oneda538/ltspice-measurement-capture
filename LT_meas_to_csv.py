import time
#from xlsxwriter.workbook import Workbook

def code(inFileName):

  infile = open(inFileName, 'r')

  lines = infile.readlines()
  infile.close()

  outFileName = "{0}_{1}.csv".format(lines[0].split('\\')[-1].split('.')[0], time.strftime('%Y-%m-%d_%H-%M-%S'))
  print outFileName

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
          measurementStartLines.append(i)

  #add headers header column
  for startLine in measurementStartLines:
      dataList[0].append(lines[startLine].split(" ")[1].strip())
      for i in range(len(dataList)-1):
          dataList[i+1].append(lines[startLine + 2 + i].split()[1])
          
  print dataList[:20]
  dataToCSV(dataList, outFileName)
  
  
def dataToCSV(dataList, outFileName):
  outfile = open(outFileName, 'w')
  for a in dataList:
      writeString = ""
      for b in a:
          outfile.write(b + ',')
      outfile.write('\n')

  outfile.close()




if __name__ == "__main__":
  inFileName = r"C:\Users\oneillda\AppData\Local\Temp\IEC61000-4-5_testbench.log"
  try:
    code(inFileName)
  except:
    print "Fail"
  finally:
    time.sleep(2)