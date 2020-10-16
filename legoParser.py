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
import re, json, requests, urllib2,traceback, time, os, csv
import logging

logger = logging.getLogger('legoParser')
logger.setLevel(logging.DEBUG)
fh = logging.FileHandler('legoParser.log')
fh.setLevel(logging.DEBUG)
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
fh.setFormatter(formatter)
logger.addHandler(fh)
'''
get the smartsheet data
TODO: replace with sdk
'''
def getSheet(sheetID):
    headers = {'Authorization': 'Bearer '+str(ssToken)}
    url = 'https://api.smartsheet.com/2.0/sheets/'+str(sheetID)
    r = requests.get(url, headers=headers)
    rArr = r.json()
    return rArr

def copySheet(sheetID,data):
    headers = {'Authorization': 'Bearer '+str(ssToken)}
    jsonData = json.dumps(data)
    url = 'https://api.smartsheet.com/2.0/sheets/'+str(sheetID)+'/copy?include=data'
    r = requests.post(url,data=jsonData, headers=headers)
    rArr = r.json()
    return rArr

def getWorkspace(workspaceID):
    headers = {'Authorization': 'Bearer '+str(ssToken)}
    url = 'https://api.smartsheet.com/2.0/workspaces/'+str(workspaceID)
    r = requests.get(url, headers=headers)
    rArr = r.json()
    return rArr

def getAttachments(sheetID):
    headers = {'Authorization': 'Bearer '+str(ssToken)}
    url = 'https://api.smartsheet.com/2.0/sheets/'+str(sheetID)+'/attachments?includeAll=True'
    r = requests.get(url, headers=headers)
    rArr = r.json()
    return rArr

def getAttachment(sheetID,attachmentID):
    headers = {'Authorization': 'Bearer '+str(ssToken)}
    # These two lines enable debugging at httplib level (requests->urllib3->http.client)
    # You will see the REQUEST, including HEADERS and DATA, and RESPONSE with HEADERS but without DATA.
    # The only thing missing will be the response.body which is not logged.
#    try:
#        import http.client as http_client
#    except ImportError:
#        # Python 2
#        import httplib as http_client
#    http_client.HTTPConnection.debuglevel = 2
#    # You must initialize logging, otherwise you'll not see debug output.
#    logging.basicConfig()
#    logging.getLogger().setLevel(logging.DEBUG)
#    requests_log = logging.getLogger("requests.packages.urllib3")
#    requests_log.setLevel(logging.DEBUG)
#    requests_log.propagate = True

    url = 'https://api.smartsheet.com/2.0/sheets/'+str(sheetID)+'/attachments/'+str(attachmentID)
    r = requests.get(url, headers=headers)
    rArr = r.json()
    return rArr

def insertRows(sheetId,data):
    headers = {'Authorization': 'Bearer '+str(ssToken)}
    jsonData = json.dumps(data)
    url = 'https://api.smartsheet.com/2.0/sheets/'+str(sheetId)+'/rows'
    r = requests.post(url, data=jsonData, headers=headers)
    return r.json()

def updateRows(sheetId,data):
    headers = {'Authorization': 'Bearer '+str(ssToken)}
    data = json.dumps(data)
    url = 'https://api.smartsheet.com/2.0/sheets/'+str(sheetId)+'/rows'
    r = requests.put(url, data=data, headers=headers)
    return r.json()

def addCellImage(sheetID,lego,columnId,image,imageSize):
    headers = {'Authorization': 'Bearer '+str(ssToken)}
    url = 'https://api.smartsheet.com/2.0/sheets/'+str(sheetID)+'/rows/'+str(lego['row'])+'/columns/'+str(columnId['picture'])+'/cellimages?altText='+str(lego['id'])
    headers['Content-Type'] = "application/jpeg"
    headers['Content-Disposition'] = 'attachment; filename="%s.jpg"'%lego['id']
    headers['Content-Length'] = str(imageSize)
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
        if column['title'] == 'Spares':
            columnId['spares'] = column['id']
        if column['title'] == 'Extra':
            columnId['extra'] = column['id']
        if column['title'] == 'Process':
            columnId['process'] = column['id']
        if column['title'] == 'Sets':
            columnId['sets'] = column['id']
        if column['title'] == 'Picture':
            columnId['picture'] = column['id']
        if column['title'] == 'Description':
            columnId['description'] = column['id']
    return columnId


'''
this function takes a pdf and pulls the data and returns the full set of legos from the pdf
'''
def getLegos(pdf):
    legos = []
    l = 1

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
            logger.debug(pageCount)
        for item in pageList:
            matchItem = False
            #matchItem = re.match(r'(\d+)x ?\n(\d+)\n',item['text'], re.M|re.I)
            matchItem = re.search(r'(\d+)x ?\n(\d+)\n',item['text'], re.M|re.I)
            if matchItem:
               lego = {}
               lego['pieces'] = int(matchItem.group(1))
               lego['id'] = matchItem.group(2)
               lego['order'] = l
               pageLegos.append(lego)
               l += 1
        if len(pageLegos) > 1:
            legos = legos + pageLegos
        pageCount += 1
    return legos

def getLegosCSV(csvName,pieceType):
    legos = []
    fb = open(csvName, 'rb')
    reader = csv.DictReader(fb)
    for line in reader:
        line[pieceType] = int(line['pieces'])
        if pieceType != 'pieces':
          del line['pieces']
        legos.append(line)
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
        lego['spares'] = False
        lego['extra'] = False
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
            if cell['columnId'] == columnId['spares']:
                try:
                    lego['spares'] = cell['value']
                except KeyError:
                    continue
            if cell['columnId'] == columnId['extra']:
                try:
                    lego['extra'] = cell['value']
                except KeyError:
                    continue
#            if cell['columnId'] == columnId['sets']:
#                try:
#                    lego['sets'] = cell['displayValue']
#                except KeyError:
#                    lego['sets'] = ''
            if cell['columnId'] == columnId['picture'] and pictures:
                try:
                    lego['picture'] = cell['image']
                except KeyError:
                    continue
        if lego['pieces'] == False:
            lego['pieces'] = 0
        if lego['spares'] == False:
            lego['spares'] = 0
        if lego['extra'] == False:
            lego['extra'] = 0
        if lego['id']:
            ssLegos.append(lego)        
    return ssLegos

'''
seperate out the legos into two groups
the ones that I already have some of,
and the ones i dont yet have any of
'''
def sortLegos(legos,ssLegos,legoSet,pieceType):
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
                    legos[a][pieceType] = lego[pieceType]
                    legos[a]['row'] = lego['row']
                    #if re.match(r''+legoSet+'',lego['sets']):
                    #    legos[a]['sets'] = legoSet
                    #else:
                    #    legos[a]['sets'] = lego['sets'] + " " + legoSet
                    legoHold = legos[a]
                    legoHold.pop('id')
                    old.append(legoHold)
                else:
                    #legos[a]['sets'] = legoSet
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
    baseUrl = 'https://www.lego.com/service/bricks'
    urlNumbers = ['5','4']
    for i in urlNumbers:
      url = '%s/%s/2/%s' % (baseUrl, i, legoID)
      r = requests.get(url)
      if len(r.content) > 0:
          byteImage = BytesIO(r.content)
          byteImage.seek(0, 2)
          size = byteImage.tell()
          image = r.content
          break
      else:
          image = False
          size = False
    return image,size
    
'''
look to see if a sheet exists for this set,
if not create one
'''
def getSetSheet(thisSet):
    data = {} 
    setId = thisSet['set']
    sheetName = thisSet['desc'] + " - "  + str(setId)
    sheetId = False
    regex = re.escape(setId)
    workspace = getWorkspace(ssWorkspace)
    for sheet in workspace['sheets']:
      if re.search(regex,sheet['name']): #, re.M|re.I):
        sheetId = sheet['id']
    if sheetId == False:
      data['destinationType'] = "workspace"
      data['destinationId'] = ssWorkspace
      data['newName'] = sheetName
      newSheet = copySheet(setTemplate,data)
      sheetId = newSheet['result']['id']
    return sheetId

'''
Main loop
'''
if __name__ == '__main__':

    '''bring in config'''
    execfile("legoParser.conf", locals())

    '''get sheet data'''
    sheet = getSheet(sheetID)
    if debug == 'smartsheet':
        logger.debug(sheet)
        raw_input("Press Enter to continue...")

    '''build list of columns'''
    columnId = getColumns(sheet)
    if debug == 'smartsheet':
        logger.debug(columnId)
        raw_input("Press Enter to continue...")

    '''Get list of attachments'''
    attachments = getAttachments(sheetID)
    if debug == 'smartsheet':
        logger.debug(attachments)
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
                    if cell['value'] == True or cell['value'] == 'pdf' or cell['value'] == 'csv':
                        rowId=each['id']
                        procType = cell['value']
                except KeyError:
                    continue
            if (cell['columnId'] == columnId['id']):
                try:
                    rowSet=cell['displayValue']
                except KeyError:
                    continue
            if (cell['columnId'] == columnId['description']):
                try:
                    rowDesc=cell['displayValue']
                except KeyError:
                    continue
        if rowId and rowSet:
            row['id'] = rowId
            row['set'] = rowSet
            row['desc'] = rowDesc
            row['procType'] = str(procType)
            rows.append(row)
    if debug == 'smartsheet':
        logger.debug(rows)
        raw_input("Press Enter to continue...")
    '''
    Performance Help:
       Run through the list of rows to be processed and select out only the needed attachments?
    '''
    '''run through all sheet attacments'''
    attachments = sorted(attachments['data'],key=itemgetter('parentId','mimeType'))
    rows = sorted(rows,key=itemgetter('id'))
    a = 0
    count = 0
    for row in rows:
        logger.info("Set: " + str(row['desc']) + " - " + str(row['set']))
        if a < len(attachments):
            while attachments[a]['parentId'] <= row['id']:
                if attachments[a]['parentId'] == row['id'] and attachments[a]['parentType'] == 'ROW' and (attachments[a]['mimeType'] == 'application/pdf' or attachments[a]['mimeType'] == 'text/csv') :
                    pieceType = 'pieces'
                    if debug == 'smartsheet':
                        logger.debug(row)
                        logger.debug(attachments[a])
                    count += 1 #debug
                    if smartsheetDown == True:
                        if attachments[a]['mimeType'] == 'application/pdf':
                            '''get attachment url and download the pdf'''
                            attachmentObj = getAttachment(sheetID,attachments[a]['id'])
                            fh = urllib2.urlopen(attachmentObj['url'])
                            localfile = open('tmp.pdf','w')
                            localfile.write(fh.read())
                            localfile.close()
                        elif attachments[a]['mimeType'] == 'text/csv':
                            '''get attachment url and download the csv'''
                            attachmentObj = getAttachment(sheetID,attachments[a]['id'])
                            fh = urllib2.urlopen(attachmentObj['url'])
                            localfile = open('tmp.csv','w')
                            localfile.write(fh.read())
                            localfile.close()
                    if (attachments[a]['mimeType'] == 'application/pdf') and (row['procType'] == 'pdf' or row['procType'] == 'True' ):
                        '''process the PDF and get the legos back'''
                        try:
                            legos = getLegos('tmp.pdf')
                        except:
                            logger.error("Failed: "+ str(row))
                            logger.debug(traceback.print_exc())
                            break
                    elif (attachments[a]['mimeType'] == 'text/csv') and (row['procType'] == 'csv' or row['procType'] == 'True'):
                        if re.search(r'spares',attachments[a]['name'], re.I):
                          pieceType = 'spares'
                        if re.search(r'extra',attachments[a]['name'], re.I):
                          pieceType = 'extra'
                        '''process the CSV and get the legos back'''
                        try:
                            legos = getLegosCSV('tmp.csv',pieceType) 
                        except:
                            logger.error("Failed: "+ str(row))
                            logger.debug(traceback.print_exc())
                            break
                    else:
                        a += 1
                        continue
                    blockCount = 0
                    for lego in legos:
                        blockCount += int(lego[pieceType])
                    logger.info("File: " + attachments[a]['name'])
                    logger.info("Total: " + str(blockCount))
                    logger.info("Types: " + str(len(legos)))
                    if debug == 'approve':
                        raw_input("Press Enter For Next")
                    if blockCount > 0:
                      setSheetID = getSetSheet(row)
                      setSheet = getSheet(setSheetID)
                      '''build list of columns'''
                      setColumnId = getColumns(setSheet)
                      if debug == 'smartsheet':
                          logger.debug(columnId)
                          raw_input("Press Enter to continue...")

                      '''get the current lego list'''
                      ssLegos = getSSLegos(setSheet,setColumnId,False)
                      if debug == 'smartsheet':
                          logger.debug(ssLegos)
                      ''' seperate out the legos we already have'''
                      newLegos,oldLegos =sortLegos(legos,ssLegos,row['set'],pieceType)
                      loggger.info("New Legos: "+str(len(newLegos)))
                      loggger.info("Old Legos: "+str(len(oldLegos)))
                      loggger.info("Total: "+str(len(newLegos)+len(oldLegos)) +" (should equal above 'Types' count)")
                      '''get the dictionary ready for smartsheet'''
                      try:
                        newLegos = sorted(newLegos,key=itemgetter('order'))
                      except:
                        pass
                      ssDataNew = prepData(newLegos,setColumnId)
                      ssDataOld = prepData(oldLegos,setColumnId)
                      if debug == 'smartsheet':
                          logger.debug("New:")
                          logger.debug(ssDataNew)
                          logger.debug("Old:")
                          logger.debug(ssDataOld)
                      '''prepare to uncheck the box so it doesn't get reprocessed'''
                      checkData = {"id":attachments[a]['parentId'],"cells":[{"columnId":columnId['process'], "value":False}]}
                      if smartsheetUp == True:
                          '''upload the data'''
                          resultNew = insertRows(setSheetID,ssDataNew)
                          if debug == 'requests':
                              logger.debug(resultNew)
                          if resultNew['resultCode'] == 0:
                              resultOld = updateRows(setSheetID,ssDataOld)
                              if debug == 'requests':
                                  logger.debug(resultOld)
                          '''if the save succeded uncheck the processing box'''
                          if resultNew['resultCode'] == 0 and resultOld['resultCode'] == 0:
                              updateRows(sheetID,checkData)
                          '''
                          Find, Download and attache the indivual images for the pieces
                          '''
                          setSheet = getSheet(setSheetID)
                          ssLegos = getSSLegos(setSheet,setColumnId,True)
                          i = 0
                          for lego in ssLegos:
                              if 'picture' not in lego: # and a > 12:
                                  logger.debug(lego)
                                  image,size = getLegoImage(lego['id'])
                                  if image:
                                      results = addCellImage(setSheetID,lego,setColumnId,image,size)
                              i += 1
                      if debug == 'smartsheet':
                          logger.debug(sheet)
                          raw_input("Press Enter to continue...")
                      '''Stop after only some sets?'''
                      if countLimit == True:
                          if count > 0:
                              exit()
                a += 1
                if a>= len(attachments):
                    break
