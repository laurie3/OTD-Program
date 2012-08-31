import xlrd
import xlwt

def checkSourceFormat(filePath):

    book = xlrd.open_workbook(filePath)
    sheet = book.sheet_by_index(1)

    nameDict = {
                  '1':'Pur Grp.',
                  '2':'Pur Grp Name',
                  '3':'Vendor',
                  '4':'Vendor Name',
                  '5':'Purch Doc',
                  '11':'Mat.No',
                  '12':'Short Text',
                  '18':'Ord Qty',
                  '20':'Sched line',
                  '21':'Stat Del D',
                  '23':'Date Txn',
                  '33':'OTD Status GR',
                  '34':'OTD Status GR'
              }
    check = True
    for columnIndex, attriName in nameDict.iteritems():
        if sheet.cell(0,int(columnIndex)).value != attriName:
            print 'format check fail for '+ attriName
            check = False

    if check:
        print 'format check pass'
    else:
        print 'format check fail'

    #print sheet.cell(1,33).value, sheet.cell(1,34).value
    print sheet.nrows

'''class Order:

    def __init__(self, vendorNo, vendorName, otdStatus1, otdStatus2):
        self.vendorNo = vendorNo
        self.vendorName = vendorName
        self.otdStatus1 = otdStatus1
        self.otdStatus2 = otdStatus2
    def testGetVendorNo(self):
        return self.vendorNo'''

class OrderBySupplier:

    def __init__(self, vendorNo, vendorName, otdStatus1OnTimeCount, otdStatus2LateCount, totalDeliveryCount):
        self.vendorNo = vendorNo
        self.vendorName = vendorName
        self.otdStatus1OnTimeCount = otdStatus1OnTimeCount
        self.otdStatus2LateCount = otdStatus2LateCount
        self.totalDeliveryCount = totalDeliveryCount

'''def processExcel(filePath):

    book = xlrd.open_workbook(filePath)
    sheet = book.sheet_by_index(1)

    itemList = []
    supplierList = []
    for i in range(1,sheet.nrows):
        order = Order(sheet.cell(i,3), sheet.cell(i,4), sheet.cell(i,33), sheet.cell(i,34))
        itemList.append(order)
        supplierList.append(sheet.cell(i,3))'''
def processExcel(filePath):

    book = xlrd.open_workbook(filePath)
    sheet = book.sheet_by_index(1)

    orderBySupplierList = []
    ADDFLAG = True

    orderBySupplierList.append(OrderBySupplier(sheet.cell(1,3).value, sheet.cell(1,4).value, 0, 0, 0))
    for i in range(1,sheet.nrows):
        for j in range(len(orderBySupplierList)):
            if sheet.cell(i,3).value==orderBySupplierList[j].vendorNo:
                if sheet.cell(i,33).value == 'On Time':
                    orderBySupplierList[j].otdStatus1OnTimeCount += 1
                if sheet.cell(i,34).value == 'Late':
                    orderBySupplierList[j].otdStatus2LateCount += 1
                orderBySupplierList[j].totalDeliveryCount += 1
                ADDFLAG = False
                break
            else:
                ADDFLAG = True
                
        if ADDFLAG:
            orderBySupplierList.append(OrderBySupplier(sheet.cell(i,3).value, sheet.cell(i,4).value, 0, 0, 0)) 
            if sheet.cell(i,33).value == 'On Time':
                orderBySupplierList[len(orderBySupplierList)-1].otdStatus1OnTimeCount += 1
            if sheet.cell(i,34).value == 'Late':
                orderBySupplierList[len(orderBySupplierList)-1].otdStatus2LateCount += 1
            orderBySupplierList[len(orderBySupplierList)-1].totalDeliveryCount += 1
            ADDFLAG = True
            #print 'add sth'

    '''print len(orderBySupplierList)
    print orderBySupplierList[0].vendorNo
    print orderBySupplierList[0].vendorName
    print orderBySupplierList[0].otdStatus1OnTimeCount
    print orderBySupplierList[0].otdStatus2LateCount
    print orderBySupplierList[0].totalDeliveryCount
    for i in range(len(orderBySupplierList)):
        print orderBySupplierList[i].vendorNo
        print orderBySupplierList[i].vendorName
        print orderBySupplierList[i].otdStatus1OnTimeCount
        print orderBySupplierList[i].otdStatus2LateCount
        print orderBySupplierList[i].totalDeliveryCount'''

    
    book = xlwt.Workbook()
    sheetResultOnTime = book.add_sheet('On time')
    sheetResultLateConfirmedDate = book.add_sheet('Late Confirmed Date')
    sheetResultOnTime.row(0).write(0,'VendorNo')
    sheetResultOnTime.row(0).write(1,'VendorName')
    sheetResultOnTime.row(0).write(2,'On Time Deliveries')
    sheetResultOnTime.row(0).write(3,'Total Deliveries')
    sheetResultOnTime.row(0).write(4,'Percentage')  
    for i in range(1,len(orderBySupplierList)):
        sheetResultOnTime.row(i).write(0,orderBySupplierList[i].vendorNo)
        sheetResultOnTime.row(i).write(1,orderBySupplierList[i].vendorName)
        sheetResultOnTime.row(i).write(2,orderBySupplierList[i].otdStatus1OnTimeCount)
        sheetResultOnTime.row(i).write(3,orderBySupplierList[i].totalDeliveryCount)
        sheetResultOnTime.row(i).write(4,float(orderBySupplierList[i].otdStatus1OnTimeCount)/float(orderBySupplierList[i].totalDeliveryCount))

    sheetResultLateConfirmedDate.row(0).write(0,'VendorNo')
    sheetResultLateConfirmedDate.row(0).write(1,'VendorName')
    sheetResultLateConfirmedDate.row(0).write(2,'Late delivery on confirmed date')
    sheetResultLateConfirmedDate.row(0).write(3,'Total Deliveries')
    sheetResultLateConfirmedDate.row(0).write(4,'Percentage')  
    for i in range(1,len(orderBySupplierList)):
        sheetResultLateConfirmedDate.row(i).write(0,orderBySupplierList[i].vendorNo)
        sheetResultLateConfirmedDate.row(i).write(1,orderBySupplierList[i].vendorName)
        sheetResultLateConfirmedDate.row(i).write(2,orderBySupplierList[i].otdStatus2LateCount)
        sheetResultLateConfirmedDate.row(i).write(3,orderBySupplierList[i].totalDeliveryCount)
        sheetResultLateConfirmedDate.row(i).write(4,float(orderBySupplierList[i].otdStatus2LateCount)/float(orderBySupplierList[i].totalDeliveryCount))
        
    book.save('result.xls')
    print 'file created'

    

    '''backupList = itemList

    print itemList[0].vendorNo
    print itemList[0].vendorName
    print itemList[0].otdStatus1
    print itemList[0].otdStatus2
    print len(itemList)
    
    tempList = []
    while len(itemList)>0:
        orderItem = itemList[0]
        for i in range(len(itemList)):
            if itemList[i].vendorNo == orederItem.vendorNo:
'''                

    
