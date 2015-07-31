import time
from xlsxwriter.workbook import Workbook

def code(inFileName):

  infile = open(inFileName, 'r')

  lines = infile.readlines()
  infile.close()

  outFileName = "{0}_{1}".format(lines[0].split('\\')[-1].split('.')[0], time.strftime('%Y-%m-%d_%H-%M-%S'))
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
  outfile = open(outFileName+".csv", 'w')
  for a in dataList:
      writeString = ""
      for b in a:
          outfile.write(b + ',')
      outfile.write('\n')

  outfile.close()

def dataToXLSX(dataList, outFileName):

  workbook = Workbook(outFileName+".xlsx")
  worksheetData = workbook.add_worksheet('Data')
  
  bold = workbook.add_format({'bold': 1})

  # Format the worksheet1 data cells that the chart will refer to
  headings = ['Time', 'Set Temp(C)', 'Actual Temp(C)', 'Set Vin(V)', 'Actual Vin(V)', 'Iin(A)', 'Vout(V)', 'Set Iout(A)', 
              'Actual Iout(A)', 'Vout AC(Vpp)', 'Vout AC(Vrms)', 'Efficiency', 'SpurFreq<200kHz(Hz)', 
              'SpurLevel<200kHz(dB)', 'SpurFreq340kHz(Hz)', 'SpurLevel340kHz(dB)', 'L1 Temperature(degC)']
  
  # Set number format "constructor"
  formatcells1 = workbook.add_format()
  formatcells2 = workbook.add_format()
  
  # Format to 3 decimal places
  formatcells1.set_num_format('0.000')
  
  # Columns A to F width set to 25. Note that it is not possible to "auto-fit" as this only happens at run-time in Excel
  worksheetData.set_column('A:A', 25)
  worksheetData.set_column('B:Q', 18)
  
  # Set horizontal and vertical alignment
  formatcells1.set_align('center') 
  formatcells1.set_align('vcenter')
  formatcells2.set_align('center') 
  formatcells2.set_align('vcenter')
  
  # Write headings row to workbook
  worksheetData.write_row('A1', headings, bold) # note that alignment not applicable to the write_row method
  
  
  ####repeat lots
  worksheetData.write('D%i' % (row_count), float(configuredInputVolts), formatcells1)
  self.worksheetData.write('A%i' % (self.row_count), time.ctime(), self.formatcells2)
  
  workbook.close()

  
if __name__ == "__main__":
  inFileName = r"C:\Users\oneillda\AppData\Local\Temp\IEC61000-4-5_testbench.log"
  try:
    code(inFileName)
  except:
    print "Fail"
  finally:
    print "Done"
    time.sleep(2)
