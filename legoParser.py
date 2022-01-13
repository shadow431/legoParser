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
from urllib.error import HTTPError
import re, json, urllib.request, urllib.parse,traceback, time, os, csv, rebrick, requests
import logging
from smartsheet import smartsheet

logFile='legoParser.log'
#logging.basicConfig(level=logging.DEBUG,filename=logFile)
logger = logging.getLogger('legoparser')
logger.setLevel(logging.DEBUG)
fh = logging.FileHandler(logFile)
#fh.setLevel(logging.DEBUG)
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
fh.setFormatter(formatter)
logger.addHandler(fh)

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
        if column['title'] == 'Color':
            columnId['color'] = column['id']
        if column['title'] == 'Release':
            columnId['release'] = column['id']
    return columnId

'''
Fill in the blanks
'''
def legoDetail(legos,columns,rebrickableAPIKey):
  i = 0
  while True:
    getDetails = False
    setDesc = False
    setColor = False
    setPic = False
    if i >= len(legos):
        break
    logger.debug(legos[i])
    try:
      legos[i]['description']
    except KeyError:
      logger.debug('no desc')
      getDetails = True
      setDesc = True
    try:
      legos[i]['color']
    except KeyError:
      getDetails = True
      setColor = True
      logger.debug('no color')
    if not ( 'picture' or 'url') in legos[i]:
      logger.debug('no pic or url')
      getDetails = True
      setPic = True
    if getDetails == True:
      details = getElementDetails(legos[i]['id'],rebrickableAPIKey)
      if setDesc == True:
        legos[i]['description'] = details['part']['name']
      if setColor == True:
        legos[i]['color'] = details['color']['name']
      if setPic == True:
        legos[i]['picture'] = details['element_img_url']
      time.sleep(1)
    logger.debug(legos[i])
    i += 1
  return legos


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
prepare the data for ss.
Here the columnId and lego dictionary keys need to match
this stiches everything together to build the smartsheet rows with parentId
'''
def prepData(legos, columnIds):
    logging.debug('prepData')
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
            if item == 'picture' and isinstance(lego[item],dict):
                continue
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
        #lego['url'] = False
        for cell in row['cells']:
            if cell['columnId'] == columnId['description']:
                try:
                    lego['description'] = cell['displayValue']
                except KeyError:
                    continue
            if cell['columnId'] == columnId['color']:
                try:
                    lego['color'] = cell['displayValue']
                except KeyError:
                    continue
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
            if cell['columnId'] == columnId['picture']:# and pictures:
                logger.debug(cell)
                try:
                    lego['picture'] = cell['image']
                except KeyError:
                    try:
                      lego['url'] = cell['value']
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
                    #legoHold = legos[a]
                    legoHold = lego
                    #legoHold.pop('id') #why?
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
Get the lego part details from rebrickable
'''
def getElementDetails(legoID,rebrickAPIKey):
    waitTime = 10
    rebrick.init(rebrickAPIKey)
    while True:
      try:
        response = rebrick.lego.get_element(legoID)
      except HTTPError as err:
          if err.code == 429:
            logger.info(f'Waiting {waitTime}secs for 429: Too many requests')
            time.sleep(waitTime)
            waitTime += waitTime
            continue
          else:
            raise
      break
    parts = json.loads(response.read())
    logger.debug(parts)
    return parts
    
'''
Get the lego set details from rebrickable
'''
def getSets(legoID,rebrickAPIKey):
    waitTime = 10
    rebrick.init(rebrickAPIKey)
    while True:
      try:
        response = rebrick.lego.get_sets(search=legoID)
      except HTTPError as err:
          if err.code == 429:
            logger.info(f'Waiting {waitTime}secs for 429: Too many requests')
            time.sleep(waitTime)
            waitTime += waitTime
            continue
          else:
            raise
      break
    parts = json.loads(response.read())
    #logger.debug(parts)
    return parts

'''
Get the lego part and images from rebrickable
'''
def getLegoImage(url):
    r = requests.get(url)
    logger.debug(f"image response: {r.status_code}")
    if len(r.content) > 0 and r.status_code == 200:
        byteImage = BytesIO(r.content)
        byteImage.seek(0, 2)
        size = byteImage.tell()
        image = r.content
    else:
        logger.debug(f"No Piece Image")
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
    workspace = ss.getWorkspace(ssWorkspace)
    for sheet in workspace['sheets']:
      if re.search(regex,sheet['name']): #, re.M|re.I):
        sheetId = sheet['id']
    if sheetId == False:
      data['destinationType'] = "workspace"
      data['destinationId'] = ssWorkspace
      data['newName'] = sheetName
      newSheet = ss.copySheet(setTemplate,data)
      sheetId = newSheet['result']['id']
    return sheetId

'''
Main loop
'''
if __name__ == '__main__':
    logger.info("Starting Lego Parser")
    '''bring in config'''
    logger.info("Reading Config")
    exec(compile(open("legoParser.conf").read(), "legoParser.conf", 'exec'), locals())
    ss = smartsheet(ssToken)
    '''get sheet data'''
    logger.info("Downloading the Sheet")
    sheet = ss.getSheet(sheetID)
    if debug == 'smartsheet':
        logger.debug(sheet)
        input("Press Enter to continue...")

    '''build list of columns'''
    logger.info("Getting Sheet Columns")
    columnId = getColumns(sheet)
    if debug == 'smartsheet':
        logger.debug(columnId)
        input("Press Enter to continue...")

    '''Get list of attachments'''
    logger.info("Getting the attachments")
    attachments = ss.getAttachments(sheetID)
    if debug == 'smartsheet':
        logger.debug(attachments)
        input("Press Enter to continue...")

    rows = []
    updates = []
    count = 0

    '''see if the row needs to be processed'''
    logger.info("Searching the rows for something to process")
    for each in sheet['rows']:
        row = {}
        update = {}
        rowId = False
        rowSet = False
        rowPhoto = False
        rowDesc = False
        rowRelease = False
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
            if (cell['columnId'] == columnId['release']):
                try:
                    rowRelease=cell['displayValue']
                except KeyError:
                    continue
            if (cell['columnId'] == columnId['picture']):
                try:
                    rowPhoto=cell['image']
                except KeyError:
                    continue
        if rowId and rowSet:
            row['id'] = rowId
            row['set'] = rowSet
            row['desc'] = rowDesc
            row['procType'] = str(procType)
            rows.append(row)
        if rowSet and (rowDesc == False or rowPhoto == False):
            update['id'] = each['id']
            update['set'] = rowSet
            update['desc'] = rowDesc
            update['photo'] = rowPhoto
            update['release'] = rowRelease
            updates.append(update)
    logger.info(f"{len(rows)} rows to process")
    if debug == 'smartsheet':
        logger.debug(rows)
        input("Press Enter to continue...")
    '''
    Performance Help:
       Run through the list of rows to be processed and select out only the needed attachments?
    '''
    '''run through all sheet attacments'''
    logger.info(f"Going through attachments")
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
                            logger.debug(f"Getting attachmentID: {attachments[a]['id']}")
                            '''get attachment url and download the pdf'''
                            attachmentObj = ss.getAttachment(sheetID,attachments[a]['id'])
                            fh = urllib.request.urlopen(attachmentObj['url'])
                            localfile = open('tmp.pdf','wb')
                            localfile.write(fh.read())
                            localfile.close()
                        elif attachments[a]['mimeType'] == 'text/csv':
                            '''get attachment url and download the csv'''
                            attachmentObj = ss.getAttachment(sheetID,attachments[a]['id'])
                            fh = urllib.request.urlopen(attachmentObj['url'])
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
                        input("Press Enter For Next")
                    if blockCount > 0:
                      setSheetID = getSetSheet(row)
                      setSheet = ss.getSheet(setSheetID)
                      '''build list of columns'''
                      setColumnId = getColumns(setSheet)
                      if debug == 'smartsheet':
                          logger.debug(columnId)
                          input("Press Enter to continue...")

                      '''get the current lego list'''
                      ssLegos = getSSLegos(setSheet,setColumnId,False)
                      if debug == 'smartsheet':
                          logger.debug(ssLegos)
                      ''' seperate out the legos we already have'''
                      newLegos,oldLegos = sortLegos(legos,ssLegos,row['set'],pieceType)
                      logger.info("New Legos: "+str(len(newLegos)))
                      logger.info("Old Legos: "+str(len(oldLegos)))
                      logger.info("Total: "+str(len(newLegos)+len(oldLegos)) +" (should equal above 'Types' count)")
                      '''get the dictionary ready for smartsheet'''
                      try:
                        newLegos = sorted(newLegos,key=itemgetter('order'))
                      except:
                        pass
                      logger.info('process pieces for full details')
                      fullDataNew = legoDetail(newLegos,setColumnId,rebrickableAPIKey)
                      fullDataOld = legoDetail(oldLegos,setColumnId,rebrickableAPIKey)
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
                          resultNew = ss.insertRows(setSheetID,ssDataNew)
                          if debug == 'requests':
                              logger.debug(resultNew)
                          if resultNew['resultCode'] == 0:
                              resultOld = ss.updateRows(setSheetID,ssDataOld)
                              logger.debug(resultOld)
                          '''if the save succeded uncheck the processing box'''
                          if resultNew['resultCode'] == 0 and resultOld['resultCode'] == 0:
                              ss.updateRows(sheetID,checkData)
                          '''
                          Find, Download and attache the indivual images for the pieces
                          '''
                          logger.info("Downloading Piece images")
                          setSheet = ss.getSheet(setSheetID)
                          ssLegos = getSSLegos(setSheet,setColumnId,True)
                          i = 0
                          for lego in ssLegos:
                              if 'url' in lego: # and a > 12:
                                  logger.debug(lego)
                                  image,size = getLegoImage(lego['url'])
                                  if image:
                                      logger.info(f"Uploading Image with size of {size}")
                                      results = ss.addCellImage(setSheetID,lego,setColumnId,image,size)
                              i += 1
                      if debug == 'smartsheet':
                          logger.debug(sheet)
                          input("Press Enter to continue...")
                      '''Stop after only some sets?'''
                      if countLimit == True:
                          if count > 0:
                              exit()
                a += 1
                if a>= len(attachments):
                    break
    sets = []
#    d = 0
    for update in updates:
        setDetails = {}
        setDetails['row'] = update['id']
        setDetails['id'] = update['set']
        logger.debug("Update: "+json.dumps(update)+"\n")
        legoSet = getSets(update['set'],rebrickableAPIKey)
        logger.debug(legoSet)
        if legoSet['count'] == 1:
            #setDetails['id'] = legoSet['results'][0]['set_num'] #rebrickable setnum
            if update['desc'] == False:
                setDetails['description'] = legoSet['results'][0]['name']
            elif update['desc'] != legoSet['results'][0]['name'] and re.search(r'\(',update['desc']) == None:
                setDetails['description'] = legoSet['results'][0]['name'] + " ("+update['desc'] +")"
            if update['photo'] == False:
                image,size = getLegoImage(legoSet['results'][0]['set_img_url'])
                if image:
                    logger.info(f"Uploading Image with size of {size}")
                    results = ss.addCellImage(sheetID,setDetails,columnId,image,size)
            if update['release'] == False:
                setDetails['release'] = legoSet['results'][0]['year']
        elif legoSet['count'] > 1:
            logger.debug("To May options for "+update['set'])
            continue
        elif legoSet['count'] == 0:
            logger.debug("No Options for "+update['set'])
            continue
        del setDetails['id']
        logger.info(setDetails)
        if len(setDetails) > 1:
          sets.append(setDetails)
#        d += 1
#        if d >= 5:
#            break
    logger.debug(sets)
    ssSetDetails = prepData(sets,columnId)
    logger.debug(ssSetDetails)
    result = ss.updateRows(sheetID,ssSetDetails)
    logger.debug(result)
    '''if the save succeded uncheck the processing box'''
    if result['resultCode'] != 0:
        logger.error(result)
