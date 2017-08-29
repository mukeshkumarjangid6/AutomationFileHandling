import os
import xlrd
import xlsxwriter

#Folder Name
outputDir = 'output'
courier = outputDir + '/courier'

#Create folder if not exist
if not os.path.isdir(outputDir):
    os.makedirs(courier)
    print("Home directory %s was created." % outputDir)

#File Refrences
inputFile = '2.xlsx'
firstOutput = '3.xlsx'
blueDartFile = 'BlueDart.xlsx'
nonBlueDartFile = 'NonBlueDart.xlsx'
pincodes = 'Pincodes.xlsx'

#Column Refrences
#column Name from inputFile
shipPostalCode = 'ship-postal-code'
#Columns to be inserted in firstOutput
rc = 'rc'
CourierName = 'Courier Name'
AWB = 'AWB'

#value to be inserted in Courier-Name column in firstOutput
BlueDart = 'BlueDart'
#Complete Done Programme
class AddColumn1:
    def __init__(self):#Method Constructor
        global inputFile
        # open original excelbook
        self.workbook = xlrd.open_workbook(inputFile)
        self.sheet = self.workbook.sheet_by_index(0)
        #Max Active Rows
        self.row_count = self.sheet.nrows
        #Max Active columns
        self.colmn_count = self.sheet.ncols
        ### new excelbook
        global firstOutput
        self.workbook = xlsxwriter.Workbook(firstOutput)
        self.worksheet = self.workbook.add_worksheet()
    def findAddColmIndex(self):#Method 1
        self.posOfAddCol = 0
        for i in range(0, self.colmn_count):
            data = [self.sheet.cell_value(0, col) for col in range(self.sheet.ncols)]
            for self.index, self.value in enumerate(data):
                if self.value.strip() == shipPostalCode:
                    self.posOfAddCol = self.index + 1
    def pro(self):#Method 2
        global CourierName
        global AWB
        for i in range(0, self.row_count):
            data1 = [self.sheet.cell_value(i, col) for col in range(self.sheet.ncols)]
            if i == 0:
                data1.insert(self.posOfAddCol, rc)
                data1.insert(self.posOfAddCol + 1, CourierName)
                data1.insert(self.posOfAddCol + 2, AWB)
            if i != 0:
                data1.insert(self.posOfAddCol, '')
                data1.insert(self.posOfAddCol + 1, '')
                data1.insert(self.posOfAddCol + 2, '')
            self.worksheet.write_row('A' + str(i + 1), data1)
        self.workbook.close()
        print("File 3 created with extra Three Columns. Total Rows Available: ", self.row_count, ".")
addColumn = AddColumn1()
addColumn.findAddColmIndex()
addColumn.pro()
#Complete Done Programme
class NewAddBlueDartFile:
    def __init__(self):#Method Constructor
        global firstOutput
        # open original excelbook
        self.workbook = xlrd.open_workbook(firstOutput)
        self.sheet = self.workbook.sheet_by_index(0)
        #Max Active Rows
        self.row_count = self.sheet.nrows
        #Max Active columns
        self.colmn_count = self.sheet.ncols
        ### new excelbook
        self.workbook = xlsxwriter.Workbook(firstOutput)
        self.worksheet = self.workbook.add_worksheet()
    def pin_Code_Lookup(self, pin):
        global pincodes
        wb = xlrd.open_workbook(pincodes)
        xlsname = pincodes
        book = xlrd.open_workbook(xlsname)
        pin = int(pin)
        sheet = book.sheet_by_index(0)
        pinCode = sheet.col_values(0)
        rC = sheet.col_values(1)
        try:
            #print(rC[pinCode.index(pin)])
            return rC[pinCode.index(pin)]
        except ValueError:
            return ''
    def findAddColmIndex(self):#Method 1
        self.posOfRc = 0
        self.CourierNameIndex = 0
        self.posOfPostCode = 0
        global shipPostalCode
        global rc
        global CourierName
        for i in range(0, self.colmn_count):
            data = [self.sheet.cell_value(0, col) for col in range(self.sheet.ncols)]
            for self.index, self.value in enumerate(data):
                if self.value.strip() == shipPostalCode:
                    self.posOfPostCode = self.index
                if self.value.strip() == rc:
                    self.posOfRc = self.index
                if self.value.strip() == CourierName:
                    self.CourierNameIndex = self.index
    def pro(self):#Method 2
        # Create file for Blank BlueDart Cell
        global outputDir
        global nonBlueDartFile
        global blueDartFile
        global BlueDart
        bookNonBlueDart = xlsxwriter.Workbook(outputDir + '/' + nonBlueDartFile)
        sheetNonBlueDart = bookNonBlueDart.add_worksheet()
        nonBlueDartIncr = 0
        dataNonBlueDart = []
        # Create file for BlueDart Cell
        bookBlueDart = xlsxwriter.Workbook(outputDir + '/' + blueDartFile)
        sheetBlueDart = bookBlueDart.add_worksheet()
        blueDartIncr = 0
        dataBlueDart = []
        print(self.row_count)
        for i in range(0, self.row_count):#Create
            data1 = [self.sheet.cell_value(i, col) for col in range(self.sheet.ncols)]
            if blueDartIncr == 0:
                dataBlueDart = list(data1)
                sheetBlueDart.write_row('A' + str(blueDartIncr + 1), dataBlueDart)
                blueDartIncr = blueDartIncr + 1
            if nonBlueDartIncr == 0:
                dataNonBlueDart = list(data1)
                sheetNonBlueDart.write_row('A' + str(nonBlueDartIncr + 1), dataNonBlueDart)
                nonBlueDartIncr = nonBlueDartIncr + 1
#            sheetBlueDart.write_row('A' + str(blueDartIncr + 1), dataBlueDart)
#            sheetNonBlueDart.write_row('A' + str(nonBlueDartIncr + 1), dataNonBlueDart)
            print("Row ", i, " is updating.")
            if i != 0:
                #data1.insert(self.posOfRc, self.pin_Code_Lookup(data1[self.posOfPostCode]))
                data1[self.posOfRc] = self.pin_Code_Lookup(data1[self.posOfPostCode])
                if data1[self.posOfRc] != '':#BlueDarts
                    #data1.insert(self.CourierNameIndex, 'BlueDart')
                    data1[self.CourierNameIndex] = BlueDart
                    dataBlueDart = list(data1)
                    sheetBlueDart.write_row('A' + str(blueDartIncr + 1), dataBlueDart)
                    blueDartIncr = blueDartIncr + 1
                elif data1[self.posOfRc] == '':#NonBlueDart
                    #data1.insert(self.CourierNameIndex, '')
                    data1[self.CourierNameIndex] = ''
                    dataNonBlueDart = list(data1)
                    sheetNonBlueDart.write_row('A' + str(nonBlueDartIncr + 1), dataNonBlueDart)
                    nonBlueDartIncr = nonBlueDartIncr + 1
            self.worksheet.write_row('A' + str(i + 1), data1)
        self.workbook.close()
newAddBlueDartFile = NewAddBlueDartFile()
newAddBlueDartFile.findAddColmIndex()
newAddBlueDartFile.pro()
print("Done2")
print(courier)
a = input('Enter to Close')