import arrow
import os
import xlrd
import xlsxwriter
#Folder Name
outputDir = 'output'
courier = outputDir + '/courier'

#File Refrence
blueDartFile = 'BlueDart.xlsx'
nonBlueDartFile = 'NonBlueDart.xlsx'
awbData = 'AWB DATA.xlsx'

#Column Refrences
#column Name from inputFile
#Common Column for corrier Files
orderId = 'order-id'
AWB = 'AWB'

#Column Name
#First File

#orderId = 'order-id'
orderItemId = 'order-item-id' # Blank
quantity = 'quantity' # Blank
shipDate = 'ship-date' # arrow.now().format('YYYY-MM-DD') after import arrow
carrierCode = 'carrier-code' # Blank
carrierName = 'carrier-name' # Courier Name
trackingNumber = 'tracking-number' # AWB
shipMethod = 'ship-method' # Blank

#Second File

#AWB = 'AWB'
#orderId = 'order-id'
recipientName = 'recipient-name'
shipAddress1 = 'ship-address-1'
shipAddress2 = 'ship-address-2'
shipCity = 'ship-city'
shipState = 'ship-state'
shipPostalCode = 'ship-postal-code'
buyerPhoneNumber = 'buyer-phone-number'

firstListCol = {orderId:'order-id', orderItemId:'order-item-id', quantity:'quantity', shipDate:'ship-date', carrierCode:'carrier-code', carrierName:'Courier Name', trackingNumber:'AWB', shipMethod:'ship-method'}
secondListCol = {AWB:'AWB', orderId:'order-id', recipientName:'recipient-name', shipAddress1:'ship-address-1', shipAddress2:'ship-address-2', shipCity:'ship-city', shipState:'ship-state', shipPostalCode:'ship-postal-code', buyerPhoneNumber:'buyer-phone-number'}


class AwbHandler:
    def awbToBlueDart(self):
        global awbData
        global outputDir
        global blueDartFile
        AWBworkbook = xlrd.open_workbook(awbData)# For Reading purpose
        AWBsheet = AWBworkbook.sheet_by_index(0)
        #Max Active Rows of AWB DATA
        AWBrow_count = AWBsheet.nrows
        #Max Active columns of AWB DATA
        AWBcolmn_count = AWBsheet.ncols
        
        blueDartWorkbook = xlrd.open_workbook(outputDir + '/' + blueDartFile)# For Reading purpose
        blueDartSheet = blueDartWorkbook.sheet_by_index(0)
        #Max Active Rows
        blueDartRow_count = blueDartSheet.nrows
        #Max Active columns
        blueDartColmn_count = blueDartSheet.ncols
        
        blueDartWorkbook1 = xlsxwriter.Workbook(outputDir + '/' + blueDartFile)# For write purpose
        blueDartSheet1 = blueDartWorkbook1.add_worksheet()
        awbDataCountOfRow, awbDataPosOfOrderId = self.lastIndexAWBDATA()# from AWB Data
        blueDartCountOfRow, blueDartPosOfAWB, posOfOrderId = self.findAWBLastCountIndex()# from blueDart AWB Column, count and posOfAWB
        listAwbFromAWBData = AWBsheet.col_values(awbDataPosOfOrderId, awbDataCountOfRow)#
        listAwbFromAWBDataIncr = 0
        if(AWBrow_count-awbDataCountOfRow) >= blueDartRow_count:
            for i in range(0, blueDartRow_count):
                dataFromBlueCart = [blueDartSheet.cell_value(i, col) for col in range(blueDartSheet.ncols)]
                if(i != 0):
                    dataFromBlueCart[blueDartPosOfAWB] = listAwbFromAWBData[i];
                blueDartSheet1.write_row('A' + str(i + 1), dataFromBlueCart)
                print("Row ", i, " is updating By AWB Data")
            return AWBrow_count
        else:
            print("Not Having Enough Tracking ID Available in AWB DATA.xlsx")
            print("Bluecart Rows Counting is More than Awb Data Counting")
            return 0
    def BlueDartToAwb(self):
        global outputDir
        global blueDartFile
        global awbData
        AWBrow_count = self.awbToBlueDart()
        if AWBrow_count != 0:
            blueDartWorkbook = xlrd.open_workbook(outputDir + '/' + blueDartFile)# For Reading purpose
            blueDartSheet = blueDartWorkbook.sheet_by_index(0)
            #Max Active Rows
            blueDartRow_count = blueDartSheet.nrows
            #Max Active columns
            blueDartColmn_count = blueDartSheet.ncols
            
            AWBworkbook = xlrd.open_workbook(awbData)# For Reading purpose
            AWBsheet = AWBworkbook.sheet_by_index(0)
            #Max Active Rows of AWB DATA
            AWBrow_count = AWBsheet.nrows
            #Max Active columns of AWB DATA
            AWBcolmn_count = AWBsheet.ncols
            AWBworkbook1 = xlsxwriter.Workbook(awbData)# For write purpose
            AWBsheet1 = AWBworkbook1.add_worksheet()
            
            awbDataCountOfRow, awbDataPosOfOrderId = self.lastIndexAWBDATA()# from AWB Data/ Column awbDataPosOfOrderId Index+1,  No of row 1 if one exist
            blueDartCountOfRow, blueDartPosOfAWB, posOfOrderId = self.findAWBLastCountIndex()# from blueDart AWB Column, count and posOfAWB
            
            listOrderIdFromBlueCart = blueDartSheet.col_values(posOfOrderId, 1)#, blueDartCountOfRow)#
            lenOfList = 0
            print('len of list of order id ', listOrderIdFromBlueCart, lenOfList)
            awbDataPosOfOrderId = awbDataPosOfOrderId + 1
            for i in range(0, AWBrow_count):
                dataFromAWB = [AWBsheet.cell_value(i, col) for col in range(AWBsheet.ncols)]
                if(i >= awbDataCountOfRow):#order-id count in awb data file
                    if lenOfList < len(listOrderIdFromBlueCart):
                        dataFromAWB[awbDataPosOfOrderId] = listOrderIdFromBlueCart[lenOfList]
                        lenOfList = lenOfList + 1
                    if lenOfList == 0:
                        dataFromAWB[awbDataPosOfOrderId] = ''
                AWBsheet1.write_row('A' + str(i + 1), dataFromAWB)
                print("Row ", i, " of AWB Data is updating By BlueDart")
                
        else:
            print("Not Having Enough Tracking ID Available in AWB DATA.xlsx")
            print("Bluecart Rows Counting is More than Awb Data Counting")
    def lastIndexAWBDATA(self):
        global awbData
        book = xlrd.open_workbook(awbData)
        sheet = book.sheet_by_index(0)
        #print(sheet.cell_value(rowx=2, colx=0))
        col1 = 0 #index of AWB 0th column No in AWB Data.xlsx
        col2 = 1 # index of order Id columns
        row_count = sheet.nrows
        count = 0
        i = 0
        msg = ""
        while(i < row_count):
            msg = sheet.cell_value(rowx=i, colx=col2)
            i = i + 1
            #print(msg)
            if(msg != ""):
                count = count + 1
                #print(count)
            elif msg == "":
                break
        #print("Returning Counting is ", count)
        return count, col1
    def findAWBLastCountIndex(self):
        global outputDir
        global blueDartFile
        global orderId
        workbook = xlrd.open_workbook(outputDir + '/' + blueDartFile)
        sheet = workbook.sheet_by_index(0)
        #Max Active Rows
        row_count = sheet.nrows
        #Max Active columns
        colmn_count = sheet.ncols
        posOfAWB = 0
        posOfOrderId = 0
        for i in range(0, colmn_count):
            data = [sheet.cell_value(0, col) for col in range(sheet.ncols)]
            for index, value in enumerate(data):
                if value.strip() == AWB:
                    posOfAWB = index
                if value.strip() == orderId:
                    posOfOrderId = index
        count = 0
        i = 0
        msg = ""
        while(i < row_count):
            msg = sheet.cell_value(rowx=i, colx=posOfAWB)
            i = i + 1
            #print(msg)
            if(msg != ""):
                count = count + 1
                #print(count)
            elif msg == "":
                break
        #print("Returning Counting is ", count)
        return count, posOfAWB, posOfOrderId
awbHandler = AwbHandler()
#awbHandler.awbToBlueDart()
awbHandler.BlueDartToAwb()

class CourierHandle:
    def __init__(self): # Method Constructor
        global outputDir
        global outputDir
        global blueDartFile
        # open original excelbook
        self.workbook = xlrd.open_workbook(outputDir + '/' + blueDartFile)
        self.sheet = self.workbook.sheet_by_index(0)
        # Max Active Rows
        self.row_count = self.sheet.nrows
        # Max Active columns
        self.colmn_count = self.sheet.ncols
    def findColmIndex(self, columnName): # Method 1
        index = -1
        for i in range(0, self.colmn_count):
            data = [self.sheet.cell_value(0, col) for col in range(self.sheet.ncols)]
            for self.index, self.value in enumerate(data):
                if self.value.strip() == columnName.strip():
                    # columnName = index
                    index = self.index
        return index
    def createFirstfile(self):
        global courier
        global blueDartFile
        global firstListCol
        global shipDate
        workbook1 = xlsxwriter.Workbook(courier + '/' + '1' + blueDartFile)
        worksheet1 = workbook1.add_worksheet()
        
        #Hold column Index No and n is for just only for increment
        colIndexList = []
        n = 0
        #Data List is for hold Single row or column from BlueDart.xlsx
        data = []
        #Holding refrence of ship-date column in new column for insert Current Date
        shipCurrDateIndex = -11
       
        #First Loop for Heading
        for key, val in firstListCol.items():#keys(): and items() and values()
            #print(val)
            colIndexList.insert(n, self.findColmIndex(val))
            data.insert(n, key)
            if(key == shipDate):
                shipCurrDateIndex = n
            n = n + 1
        worksheet1.write_row('A1', data)
        #Empty all items from data list
        del data[:]      # equivalent to   del data[0:len(data)]
        n = 0
        rCount = 2
        #Second Loop for Rest Rows except Heading
        for row in range(1, self.row_count):
            for col in colIndexList:
                value  = (self.sheet.cell(row, col).value)
                data.insert(n, value)
                n = n + 1
            if (shipCurrDateIndex != -1):
                data[shipCurrDateIndex] = arrow.now().format('YYYY-MM-DD')
            worksheet1.write_row('A' + str(rCount), data)
            rCount = rCount + 1
            del data[:]
        print("Courier File One Created")
    def createSecondFile(self):
        global courier
        global blueDartFile
        global secondListCol
        workbook2 = xlsxwriter.Workbook(courier + '/' + '2' + blueDartFile)
        worksheet2 = workbook2.add_worksheet()
        
        #Hold column Index No and n is for just only for increment
        colIndexList = []
        n = 0
        #Data List is for hold Single row or column from BlueDart.xlsx
        data = []
        #First Loop for Heading
        for key, val in secondListCol.items():#keys(): and items() and values()
            #print(val)
            colIndexList.insert(n, self.findColmIndex(val))
            data.insert(n, key)
            n = n + 1
        worksheet2.write_row('A1', data)
        #Empty all items from data list
        del data[:]      # equivalent to   del data[0:len(data)]
        
        n = 0
        rCount = 2
        #Second Loop for Rest Rows except Heading
        for row in range(1, self.row_count):
            for col in colIndexList:
                value  = (self.sheet.cell(row, col).value)
                data.insert(n, value)
                n = n + 1
            worksheet2.write_row('A' + str(rCount), data)
            rCount = rCount + 1
            del data[:]
        print("Courier File Two Created")
courierHandle = CourierHandle()
courierHandle.createFirstfile()
courierHandle.createSecondFile()

print("Complete")
a = input('Enter to Close')