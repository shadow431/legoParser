from pdfminer.layout import LAParams
from pdfminer.converter import PDFPageAggregator
from pdfminer.pdfinterp import PDFResourceManager
from pdfminer.pdfinterp import PDFPageInterpreter
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfparser import PDFParser
from PIL import Image
from io import BytesIO
import pdfminer
from operator import itemgetter
import re, json, requests, urllib2,traceback, time, os


'''
get the smartsheet data
TODO: replace with sdk
'''
def getSheet(sheetID):
    url = 'https://api.smartsheet.com/2.0/sheets/'+str(sheetID)
    r = requests.get(url, headers=headers)
    rArr = r.json()
    return rArr

def getAttachments(sheetID):
    url = 'https://api.smartsheet.com/2.0/sheets/'+str(sheetID)+'/attachments?includeAll=True'
    r = requests.get(url, headers=headers)
    rArr = r.json()
    return rArr

def getAttachment(sheetID,attachmentID):
    url = 'https://api.smartsheet.com/2.0/sheets/'+str(sheetID)+'/attachments/'+str(attachmentID)
    r = requests.get(url, headers=headers)
    rArr = r.json()
    return rArr

def insertRows(sheetId,data):
    jsonData = json.dumps(data)
    url = 'https://api.smartsheet.com/2.0/sheets/'+str(sheetId)+'/rows'
    r = requests.post(url, data=jsonData, headers=headers)
    return r.json()

def updateRows(sheetId,data):
    data = json.dumps(data)
    url = 'https://api.smartsheet.com/2.0/sheets/'+str(sheetId)+'/rows'
    r = requests.put(url, data=data, headers=headers)
    return r.json()

def addCellImage(sheetID,lego,columnId,image,imageSize):
    url = 'https://api.smartsheet.com/2.0/sheets/'+str(sheetID)+'/rows/'+str(lego['row'])+'/columns/'+str(columnId['picture'])+'/cellimages?altText='+str(lego['id'])
    headers['Content-Type'] = "application/jpeg"
    headers['Content-Disposition'] = 'attachment; filename="%s.jpg"'%lego['id']
    headers['Content-Length'] = imageSize
    r = requests.post(url, data=image, headers=headers)
    return r.json


'''
using the sheet data, get a dictionary of columnId's that we care about
'''
def getColumns(sheet):
    columnId ={}
    for column in sheet['columns']:
        if column['title'] == 'Id':
            columnId['id'] = column['id']
        if column['title'] == 'Pieces':
            columnId['pieces'] = column['id']
        if column['title'] == 'Process':
            columnId['process'] = column['id']
        if column['title'] == 'Sets':
            columnId['sets'] = column['id']
        if column['title'] == 'Picture':
            columnId['picture'] = column['id']
    return columnId


'''
this function takes a pdf and pulls the data and returns the full set of legos from the pdf
'''
def getLegos(pdf):
    legos = []

    ''' Set parameters for pdf analysis.'''
    laparams = LAParams()
    rsrcmgr = PDFResourceManager()
    fp = open(pdf, 'rb')
    parser = PDFParser(fp)
    document = PDFDocument(parser)

    ''' Create a PDF page aggregator object.'''
    device = PDFPageAggregator(rsrcmgr, laparams=laparams)
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    pages = list(enumerate(PDFPage.create_pages(document)))
    pageCount=1
    totalPages = len(pages)

    '''process each page'''
    for pageNumber,page in pages:


        interpreter.process_page(page)
        # receive the LTPage object for the page.
        layout = device.get_result()
        pageList = []

        ''' go through everything, only grabbing the text'''
        for objType in layout:
            objDict = {}
            if (isinstance(objType, pdfminer.layout.LTTextBoxHorizontal)):
                 objDict['height'] = objType.y1
                 objDict['width'] = objType.x0
                 objDict['text'] = objType.get_text()
                 pageList.append(objDict)

        '''
        find the text we are looking for and ignore the rest
        '''
        pageLegos = []
        if debug == 'pdf':
            print pageCount
        for item in pageList:
            matchItem = False
            #matchItem = re.match(r'(\d+)x ?\n(\d+)\n',item['text'], re.M|re.I)
            matchItem = re.search(r'(\d+)x ?\n(\d+)\n',item['text'], re.M|re.I)
            if matchItem:
               lego = {}
               lego['pieces'] = int(matchItem.group(1))
               lego['id'] = matchItem.group(2)
               pageLegos.append(lego)
        if len(pageLegos) > 1:
            legos = legos + pageLegos
        pageCount += 1
    return legos

'''
prepare the data for smartsheet.
Here the columnId and lego dictionary keys need to match
this stiches everything together to build the smartsheet rows with parentId
'''
def prepData(legos, columnIds):
    ssdata = []
    for lego in legos:
        row = {}
        row['cells'] = []
        try:
            row['id'] = lego['row']
        except KeyError:
            row['toBottom'] = True
            pass
        for item in lego:
            try:
                columns ={}
                columns['columnId'] = columnIds[item]
                columns['value'] = lego[item]
                row['cells'].append(columns)
            except:
                pass
        ssdata.append(row)

    return ssdata

'''''
Get a list of the legos currently in the sheet to check for existing blocks
'''''
def getSSLegos(sheet,columnId,pictures):
    ssLegos = []
    for row in sheet['rows']:
        lego = {} 
        lego['row'] = row['id']
        lego['id'] = False
        lego['pieces'] = False
        for cell in row['cells']:
            if cell['columnId'] == columnId['id']:
                try:
                    lego['id'] = cell['displayValue']
                except KeyError:
                    continue
            if cell['columnId'] == columnId['pieces']:
                try:
                    lego['pieces'] = cell['value']
                except KeyError:
                    continue
            if cell['columnId'] == columnId['sets']:
                try:
                    lego['sets'] = cell['displayValue']
                except KeyError:
                    lego['sets'] = ''
            if cell['columnId'] == columnId['picture'] and pictures:
                try:
                    lego['picture'] = cell['image']
                except KeyError:
                    continue
        if lego['pieces'] == False:
            lego['pieces'] = 0
        if lego['id']:
            ssLegos.append(lego)        
    return ssLegos

'''
seperate out the legos into two groups
the ones that I already have some of,
and the ones i dont yet have any of
'''
def sortLegos(legos,ssLegos,legoSet):
    new = []
    old = []
    legos = sorted(legos,key=itemgetter('id'))
    ssLegos = sorted(ssLegos,key=itemgetter('id'))
    a = 0
    count = 0
    for lego in ssLegos:
        if a < len(legos):
            while legos[a]['id'] <= lego['id']:
                if legos[a]['id'] == lego['id']:
                    legos[a]['pieces'] += lego['pieces']
                    legos[a]['row'] = lego['row']
                    if re.match(r''+legoSet+'',lego['sets']):
                        legos[a]['sets'] = legoSet
                    else:
                        legos[a]['sets'] = lego['sets'] + " " + legoSet
                    legoHold = legos[a]
                    legoHold.pop('id')
                    old.append(legoHold)
                else:
                    legos[a]['sets'] = legoSet
                    new.append(legos[a])
                a += 1
                if a >= len(legos):
                    break
    while a < len(legos):
        legos[a]['sets'] = legoSet
        new.append(legos[a])
        a += 1
    return new,old

'''
go to the lego site and get the picture
NOTE: i just grabbed the URL, I dont know how long this url works for
'''
def getLegoImage(legoID):
    url = 'http://cache.lego.com/media/bricks/5/2/%s.jpg' % legoID
    r = requests.get(url)
    if len(r.content) > 0:
        byteImage = BytesIO(r.content)
        byteImage.seek(0, 2)
        size = byteImage.tell()
        image = r.content
    else:
        image = False
        size = False
    return image,size
    

'''
Main loop
'''
if __name__ == '__main__':

    '''bring in config'''
    execfile("legoParser.conf", locals())
    headers = {'Authorization': 'Bearer '+str(ssToken)}

    '''get sheet data'''
    sheet = getSheet(sheetID)
    if debug == 'smartsheet':
        print sheet
        raw_input("Press Enter to continue...")

    '''build list of columns'''
    columnId = getColumns(sheet)
    if debug == 'smartsheet':
        print columnId
        raw_input("Press Enter to continue...")

    '''Get list of attachments'''
    attachments = getAttachments(sheetID)
    if debug == 'smartsheet':
        print attachments
        raw_input("Press Enter to continue...")

    rows = []
    count = 0

    '''see if the row needs to be processed'''
    for each in sheet['rows']:
        row = {}
        rowId = False
        rowSet = False
        for cell in each['cells']:
            if (cell['columnId'] == columnId['process']):
                try:
                    if (cell['value'] == True):
                        rowId=each['id']
                except KeyError:
                    continue
            if (cell['columnId'] == columnId['id']):
                try:
                    rowSet=cell['displayValue']
                except KeyError:
                    continue
        if rowId and rowSet:
            row['id'] = rowId
            row['set'] = rowSet
            rows.append(row)
    if debug == 'smartsheet':
        print rows
        raw_input("Press Enter to continue...")
    '''
    Performance Help:
       Run through the list of rows to be processed and select out only the needed attachments?
    '''
    '''run through all sheet attacments'''
    attachments = sorted(attachments['data'],key=itemgetter('parentId'))
    rows = sorted(rows,key=itemgetter('id'))
    a = 0
    count = 0
    for row in rows:
        if a < len(attachments):
            while attachments[a]['parentId'] <= row['id']:
                if attachments[a]['parentId'] == row['id'] and attachments[a]['parentType'] == 'ROW' and attachments[a]['mimeType'] == 'application/pdf':
                    if debug == 'smartsheet':
                        print row
                        print attachments[a]
                    count += 1 #debug
                    if smartsheetDown == True:
                        '''get attachment url and download the pdf'''
                        attachmentObj = getAttachment(sheetID,attachments[a]['id'])
                        fh = urllib2.urlopen(attachmentObj['url'])
                        localfile = open('tmp.pdf','w')
                        localfile.write(fh.read())
                        localfile.close()
                    '''process the PDF and get the legos back'''
                    try:
                        legos = getLegos('tmp.pdf') #wass getMeals
                    except:
                        print "Failed: "+ str(row)
                        print traceback.print_exc()
                        break
                    blockCount = 0
                    for lego in legos:
                        blockCount += int(lego['pieces'])
                    print "File: " + attachments[a]['name']
                    print "Total: " + str(blockCount)
                    print "Types: " + str(len(legos))
                    if debug == 'approve':
                        raw_input("Press Enter For Next")
                    '''get the current lego list'''
                    ssLegos = getSSLegos(sheet,columnId,False)
                    if debug == 'smartsheet':
                        print ssLegos
                    ''' seperate out the legos we already have'''
                    newLegos,oldLegos =sortLegos(legos,ssLegos,row['set'])
                    print "New Legos: "+str(len(newLegos))
                    print "Old Legos: "+str(len(oldLegos))
                    print "Total: "+str(len(newLegos)+len(oldLegos)) +" (should equal above 'Types' count"
                    '''get the dictionary ready for smartsheet'''
                    ssDataNew = prepData(newLegos,columnId)
                    ssDataOld = prepData(oldLegos,columnId)
                    if debug == 'smartsheet':
                        print "New:"
                        print ssDataNew
                        print "Old:"
                        print ssDataOld
                    '''prepare to uncheck the box so it doesn't get reprocessed'''
                    checkData = {"id":attachments[a]['parentId'],"cells":[{"columnId":columnId['process'], "value":False}]}
                    if smartsheetUp == True:
                        '''upload the data'''
                        resultNew = insertRows(sheetID,ssDataNew)
                        if debug == 'requests':
                            print resultNew
                        if resultNew['resultCode'] == 0:
                            resultOld = updateRows(sheetID,ssDataOld)
                            if debug == 'requests':
                                print resultOld
                        '''if the save succeded uncheck the processing box'''
                        if resultNew['resultCode'] == 0 and resultOld['resultCode'] == 0:
                            updateRows(sheetID,checkData)
                    '''get a new copy of the sheet'''
                    sheet = getSheet(sheetID)
                    if debug == 'smartsheet':
                        print sheet
                        raw_input("Press Enter to continue...")
                    '''Stop after only some sets?'''
                    if countLimit == True:
                        if count > 0:
                            exit()
                a += 1
                if a>= len(attachments):
                    break
    '''
    Find, Download and attache the indivual images for the pieces
    '''
    ssLegos = getSSLegos(sheet,columnId,True)
    a = 0
    for lego in ssLegos:
        if 'picture' not in lego and a > 12:
            image,size = getLegoImage(lego['id'])
            if image:
                results = addCellImage(sheetID,lego,columnId,image,size)
        a += 1
