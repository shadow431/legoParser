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
import sys
from operator import itemgetter
from urllib.error import HTTPError
import re, json, urllib.request, urllib.parse,traceback, time, os, csv, rebrick, requests
import logging
from smartsheet import smartsheet
from aws import aws
from dotenv import load_dotenv

'''
Setup Logging
'''

logFile='legoParser.log'
#logging.basicConfig(level=logging.DEBUG,filename=logFile)
logger = logging.getLogger('legoparser')
streamHandler = logging.StreamHandler(sys.stdout)
logger.setLevel(logging.INFO)
#fh = logging.FileHandler(logFile)
#fh.setLevel(logging.DEBUG)
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(lineno)d - %(message)s')
#fh.setFormatter(formatter)
streamHandler.setFormatter(formatter)
#logger.addHandler(fh)
logger.addHandler(streamHandler)


'''
SMARTSHEET STUFF
'''

'''
using the sheet data, get a dictionary of columnId's that we care about
'''
def getColumns(sheet):
    columnId ={}
    for column in sheet['columns']:
        #SETLIST,SAMLIST,SETS,ELEMENTS
        if column['title'] == 'Id':
            columnId['id'] = column['id']
        if column['title'] == 'Description':
            columnId['description'] = column['id']
        if column['title'] == 'Picture':
            columnId['picture'] = column['id']
        if column['title'] == 'Release':
            columnId['release'] = column['id']
        if column['title'] == 'Pieces':
            columnId['pieces'] = column['id']
        if column['title'] == 'Color':
            columnId['color'] = column['id']

        #SETLIST,SAMLIST
        if column['title'] == 'Lego Theme':
            columnId['theme'] = column['id']
        if column['title'] == 'Themes':
            columnId['themes'] = column['id']
        if column['title'] == 'MSRP':
            columnId['msrp'] = column['id']
        if column['title'] == 'Value':
            columnId['value'] = column['id']

        #SETS,ELEMENTS
        if column['title'] == 'Design':
            columnId['design'] = column['id']

        #SETS
        if column['title'] == 'Spares':
            columnId['spares'] = column['id']
        if column['title'] == 'Extra':
            columnId['extra'] = column['id']

        #SETLIST
        if column['title'] == 'Process':
            columnId['process'] = column['id']

        #UNUSED
        if column['title'] == 'Sets':
            columnId['sets'] = column['id']

    return columnId

'''
look to see if a sheet exists for this set,
if not create one
'''
def getSetSheet(set_id, desc):
    data = {}
    id_length = len(set_id)
    sep_length = 3
    max_length = 50
    remain_length = max_length - (id_length + sep_length)
    sheetName = f"{desc[0:remain_length]} - {set_id}"
    sheetId = False
    regex = re.escape(set_id)
    workspace = ss.getWorkspace(ssWorkspace)
    for sheet in workspace['sheets']:
      if re.search(regex,sheet['name']): #, re.M|re.I):
        sheetId = sheet['id']
    if sheetId == False:
      data['destinationType'] = "folder"
      data['destinationId'] = ssSetsFolder
      data['newName'] = sheetName
      newSheet = ss.copySheet(setTemplate,data)
      logger.debug(newSheet)
      sheetId = newSheet['result']['id']
    return sheetId

'''
REBRICKABLE STUFF
'''

'''
Get the lego part details from rebrickable
'''
def getElementDetails(legoID,rebrickAPIKey):
    return getRebrickData(legoID,'elements',rebrickAPIKey)

'''
Get the lego sets from rebrickable
'''
def getSets(legoID,rebrickAPIKey):
    return getRebrickData(legoID,'sets',rebrickAPIKey)

'''
Get the lego set details from rebrickable
'''
def getSet(set_id,rebrickAPIKey):
    return getRebrickData(set_id,'set',rebrickAPIKey)

'''
Get the lego Theme details from rebrickable
'''
def getTheme(themeID,rebrickAPIKey):
    return getRebrickData(themeID,'theme',rebrickAPIKey)

'''
Get the lego Theme details from rebrickable
'''
def getRebrickData(itemID,callType,rebrickAPIKey):
    waitTime = 10
    rebrick.init(rebrickAPIKey)
    while True:
      try:
          if callType == 'elements':
            response = rebrick.lego.get_element(itemID)
          elif callType == 'sets':
            response = rebrick.lego.get_sets(search=itemID)
          elif callType == 'set':
            response = rebrick.lego.get_set(itemID)
          elif callType == 'theme':
            response = rebrick.lego.get_theme(itemID)
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
PROCESSING STUFF
'''

'''
Fill in the blanks for the bricks
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
      try:
        details = getElementDetails(legos[i]['id'],rebrickableAPIKey)
      except:
        i += 1
        continue
      if setDesc == True:
        legos[i]['description'] = details['part']['name']
      if setColor == True:
        legos[i]['color'] = details['color']['name']
      if setPic == True:
        legos[i]['picture'] = details['element_img_url']
      logger.debug(details)
      time.sleep(1)
    logger.debug(legos[i])
    i += 1
  return legos

'''
prepare the data for Smartsheet.
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
Get a list of the bricks currently in the sheet to check for existing blocks
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
seperate out the bricks into two groups
the ones that I already have some of,
and the ones i dont yet have any of
Is this really needed any more? was for when it was a single sheet for all the pieces, will i ever want to increment the list?
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
If More then one set result from rebrickable, try and find the correct one
'''
def pick_a_set(sets, set_id):
  for each in sets:
    if "-" in set_id:
      logger.debug(f"Found hypen in {set_id}, using full {each['set_num']}")
      derived_set_id = [each['set_num']]
    else:
      logger.debug(f"No hypen in {set_id}, using split {each['set_num']}")
      derived_set_id = each['set_num'].split('-')
    if set_id == derived_set_id[0]:
        match = each
        break
    else:
        match = None
  return match

'''
Take a sets details and update them with rebrickable info
'''
def setUpdate(rowId, itemID, desc, photo, release, theme, pieces, rebrickableAPIKey, data, columnId, ss):
  sheetID = data['id']
  logger.debug("Update: "+json.dumps({'row': rowId, 'item': itemID, 'photo': photo, 'release': release, 'theme': theme, 'pieces': pieces}))
  setDetails = {}
  setDetails['row'] = rowId
  setDetails['id'] = itemID
  legoDetails = getRebrickData(itemID,data['type'],rebrickableAPIKey)
  logger.debug(legoDetails)
  if legoDetails['count'] == 1:
    rebrickDetail = legoDetails['results'][0]
  elif legoDetails['count'] > 1:
    logger.debug("To May options for "+itemID)
    rebrickDetail = pick_a_set(legoDetails['results'], itemID)
    logger.debug(f"Posible Match: {rebrickDetail}")
  elif legoDetails['count'] == 0:
    logger.debug("No Options for "+itemID)
    rebrickDetail = {}
  #setDetails['id'] = legoDetails['results'][0]['set_num'] #rebrickable setnum
  if len(rebrickDetail) > 0:
    logger.debug("Updating Set info")
    if desc == False:
      setDetails['description'] = rebrickDetail['name']
    elif desc != rebrickDetail['name'] and re.search(r'\(',desc) == None:
      setDetails['description'] = rebrickDetail['name'] + " ("+desc +")"
    if photo == False:
      image,size = getLegoImage(rebrickDetail['set_img_url'])
      if image:
        setDetails['set_img_url'] = rebrickDetail['set_img_url']
        logger.info(f"Uploading Image with size of {size}")
        results = ss.addCellImage(sheetID,setDetails,columnId,image,size)
    if release == False:
      setDetails['release'] = rebrickDetail['year']
    if pieces == False:
      setDetails['pieces'] = rebrickDetail['num_parts']
    if theme == False:
      rbTheme = getTheme(rebrickDetail['theme_id'],rebrickableAPIKey)
      logger.debug(rbTheme)
      if rbTheme['parent_id'] != None:
        rbTheme = getTheme(rbTheme['parent_id'],rebrickableAPIKey)
        logger.debug(rbTheme)
      setDetails['theme'] = rbTheme['name']
    del setDetails['id']
    logger.info(setDetails)
  else:
    setDetails = {}
  return setDetails

'''
Take an elemenets details and update them with rebrickable info
'''
def elementUpdate(rowId, itemID, desc, photo, release,rebrickableAPIKey, data, columnId, ss, design, color):
  sheetID = data['id']
  logger.debug("Update: "+json.dumps({'row': rowId, 'item': itemID, 'Description': desc, 'photo': photo, 'release': release, 'design': design, 'color': color}))
  elementDetails = {}
  elementDetails['row'] = rowId
  elementDetails['id'] = itemID
  legoDetails = getRebrickData(itemID,data['type'],rebrickableAPIKey)
  logger.debug(legoDetails)
  rebrickDetail = legoDetails
  if len(rebrickDetail) > 0:
    logger.debug("Updating Set info")
    if desc == False:
      elementDetails['description'] = rebrickDetail['part']['name']
    elif desc != rebrickDetail['part']['name'] and re.search(r'\(',desc) == None:
      elementDetails['description'] = rebrickDetail['part']['name'] + " ("+desc +")"
    if photo == False:
      if rebrickDetail['element_img_url'] != None:
        url = rebrickDetail['element_img_url']
      else:
        logger.info(f"Falling back from element_img_url to part[part_img_url] for {itemID}")
        url = rebrickDetail['part']['part_img_url']
      image,size = getLegoImage(url)
      if image:
        logger.info(f"Uploading Image with size of {size}")
        results = ss.addCellImage(sheetID,elementDetails,columnId,image,size)
    if release == False:
      elementDetails['release'] = rebrickDetail['part']['year_from']
    if design == False:
      elementDetails['design'] = rebrickDetail['design_id']
    if color == False:
      elementDetails['color'] = rebrickDetail['color']['name']
    del elementDetails['id']
    logger.info(elementDetails)
  else:
    elementDetails = {}
  return elementDetails

'''
Process a row for attachments and create/update a Set inventory sheet based off them
'''
def row_process(ss, sheet_id, columnId, row_id, set_id, title, proc_type, rebrickableAPIKey, download=True, upload=True):
  logger.info("Processing Set {set_id}")
  proc_type = str(proc_type)
  logger.info("Getting Attachments")
  attachmentsObj = ss.getAttachments(sheet_id,row_id)
  if attachmentsObj['totalPages'] > 1:
      logger.error("Row attachemts Greater the 1 Page")
  logger.debug(attachmentsObj)
  for attachment in attachmentsObj['data']:
    logger.info("Check Attachment Type")
    if attachment['mimeType'] == 'application/pdf' or attachment['mimeType'] == 'text/csv':
      pieceType = 'pieces'
      logger.debug(f'Row: {row_id}, Set: {set_id}')
      logger.debug(attachment)
      #count += 1 #debug gone
      if download == True:
          logger.info("Downloading Attachment")
          if attachment['mimeType'] == 'application/pdf':
              logger.debug(f"Getting attachmentID: {attachment['id']}")
              '''get attachment url and download the pdf'''
              attachmentObj = ss.getAttachment(sheet_id,attachment['id'])
              fh = urllib.request.urlopen(attachmentObj['url'])
              localfile = open('./tmp.pdf','wb')
              localfile.write(fh.read())
              localfile.close()
          elif attachment['mimeType'] == 'text/csv':
              '''get attachment url and download the csv'''
              attachmentObj = ss.getAttachment(sheet_id,attachment['id'])
              fh = urllib.request.urlopen(attachmentObj['url'])
              localfile = open('tmp.csv','wb')
              localfile.write(fh.read())
              localfile.close()
      if (attachment['mimeType'] == 'application/pdf') and (proc_type == 'pdf' or proc_type == 'True' ):
          '''process the PDF and get the legos back'''
          try:
              legos = getLegos('./tmp.pdf')
          except:
              logger.error(f"Failed: {row_id}")
              logger.debug(traceback.print_exc())
              break
      elif (attachment['mimeType'] == 'text/csv') and (proc_type == 'csv' or proc_type == 'True'):
          if re.search(r'spares',attachment['name'], re.I):
            pieceType = 'spares'
          if re.search(r'extra',attachment['name'], re.I):
            pieceType = 'extra'
          '''process the CSV and get the legos back'''
          try:
              legos = getLegosCSV('tmp.csv',pieceType)
          except:
              logger.error(f"Failed: {row_id}")
              logger.debug(traceback.print_exc())
              break
      else:
        logger.error(f"No useable attachments for Set {set_id}, Row {row_id}")

      blockCount = 0
      for lego in legos:
          blockCount += int(lego[pieceType])
      logger.info("File: " + attachment['name'])
      logger.info("Total: " + str(blockCount))
      logger.info("Types: " + str(len(legos)))
      if blockCount > 0:
        setSheetID = getSetSheet(set_id, title)
        setSheet = ss.getSheet(setSheetID)
        '''build list of columns'''
        setColumnId = getColumns(setSheet)
        logger.debug(setColumnId)

        '''get the current lego list'''
        ssLegos = getSSLegos(setSheet,setColumnId,False)
        logger.debug(ssLegos)
        ''' seperate out the legos we already have'''
        newLegos,oldLegos = sortLegos(legos,ssLegos,set_id,pieceType)
        logger.info("New Legos: "+str(len(newLegos)))
        logger.info("Old Legos: "+str(len(oldLegos)))
        logger.info("Total: "+str(len(newLegos)+len(oldLegos)) +" (should equal above 'Types' count)")
        '''get the dictionary ready for smartsheet'''
        try:
          newLegos = sorted(newLegos,key=itemgetter('order'))
        except:
          pass
#        logger.info('process pieces for full details')
#        fullDataNew = legoDetail(newLegos,setColumnId,rebrickableAPIKey)
#        fullDataOld = legoDetail(oldLegos,setColumnId,rebrickableAPIKey)
        ssDataNew = prepData(newLegos,setColumnId)
        ssDataOld = prepData(oldLegos,setColumnId)
        logger.debug("New:")
        logger.debug(ssDataNew)
        logger.debug("Old:")
        logger.debug(ssDataOld)
        '''prepare to uncheck the box so it doesn't get reprocessed'''
        checkData = {"id":attachment['parentId'],"cells":[{"columnId":columnId['process'], "value":False}]}
        if upload == True:
            '''upload the data'''
            resultNew = ss.insertRows(setSheetID,ssDataNew)
            logger.debug(resultNew)
            if resultNew['resultCode'] == 0:
                resultOld = ss.updateRows(setSheetID,ssDataOld)
                logger.debug(resultOld)
            '''if the save succeded uncheck the processing box'''
            if resultNew['resultCode'] == 0 and resultOld['resultCode'] == 0:
                ss.updateRows(sheet_id,checkData)
            sheetInfo = {'id': setSheetID ,'type':'elements'}
            sheet_proc(ss,sheetInfo,rebrickableAPIKey,download,upload)
#            '''
#            Find, Download and attache the indivual images for the pieces
#            '''
#            logger.info("Downloading Piece images")
#            setSheet = ss.getSheet(setSheetID)
#            ssLegos = getSSLegos(setSheet,setColumnId,True)
#            i = 0
#            for lego in ssLegos:
#                if 'url' in lego: # and a > 12:
#                    logger.debug(lego)
#                    image,size = getLegoImage(lego['url'])
#                    if image:
#                        logger.info(f"Uploading Image with size of {size}")
#                        results = ss.addCellImage(setSheetID,lego,setColumnId,image,size)
#                i += 1
  return

'''
Process through the rows in a sheet to get set and element details
'''

def sheet_proc(ss, data,rebrickableAPIKey,smartsheetDown,smartsheetUp,countLimit):
  '''get sheet data'''
  logger.info("Downloading the Sheet")
  sheet = ss.getSheet(data['id'])
  logger.debug(sheet)

  '''build list of columns'''
  logger.info("Getting Sheet Columns")
  columnId = getColumns(sheet)
  logger.debug(columnId)

  sets = []
  count = 0

  '''see if the row needs to be processed'''
  logger.info("Searching the rows for something to process")
  for each in sheet['rows']:
      rowId = False
      rowSet = False
      rowPhoto = False
      rowDesc = False
      rowRelease = False
      rowDesign = False
      rowColor = False
      rowTheme = False
      rowPieces = False

      '''Check the data for a set and see if its marked for processing, or if it has missing fields'''
      for cell in each['cells']:
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
          try:
            if (cell['columnId'] == columnId['process']):
                try:
                    if cell['value'] == True or cell['value'] == 'pdf' or cell['value'] == 'csv':
                        rowId=each['id']
                        procType = cell['value']
                except KeyError:
                    continue
          except KeyError:
              pass
          try:
            if (cell['columnId'] == columnId['design']):
              try:
                  rowDesign=cell['displayValue']
              except KeyError:
                  continue
          except KeyError:
              pass
          try:
            if (cell['columnId'] == columnId['color']):
              try:
                  rowColor=cell['displayValue']
              except KeyError:
                  continue
          except KeyError:
              pass
          try:
            if (cell['columnId'] == columnId['theme']):
              try:
                  rowTheme=cell['displayValue']
              except KeyError:
                  continue
          except KeyError:
              pass
          try:
            if (cell['columnId'] == columnId['pieces']):
              try:
                  rowPieces=cell['displayValue']
              except KeyError:
                  continue
          except KeyError:
              pass
      '''If rowId is set, meaning process column is checked, proccess the attachments to create an inventory sheet'''
      if rowId and rowSet and rowDesc != False:
          row_process(ss, data['id'], columnId, rowId, rowSet, rowDesc, procType, rebrickableAPIKey, download=smartsheetDown, upload=smartsheetUp)
      '''If the set is missing a photo or description check rebrickable to try and fill them in'''
      if rowSet and (rowDesc == False or rowPhoto == False):
          if data['type'] == 'sets':
            setDetails = setUpdate(each['id'], rowSet, rowDesc, rowPhoto, rowRelease, rowTheme, rowPieces, rebrickableAPIKey, data, columnId, ss)
            logger.debug(len(setDetails))
          elif data['type'] == 'elements':
            logger.info("do something elemenetal here")
            setDetails = elementUpdate(each['id'], rowSet, rowDesc, rowPhoto, rowRelease, rebrickableAPIKey, data, columnId, ss, rowDesign, rowColor)
          else:
            logger.info("I don't know what to do here")
            continue
          if len(setDetails) > 1:
            logger.info("Appending Set")
            sets.append(setDetails)
          count = count + 1
          if (countLimit != False) and (count > countLimit):
            break

  '''update any sets with updated info from rebrickabl'''
  if len(sets) > 0:
    logger.debug(sets)
    ssSetDetails = prepData(sets,columnId)
    logger.debug(ssSetDetails)
    result = ss.updateRows(data['id'],ssSetDetails)
    logger.debug(result)
    if result['resultCode'] != 0:
        logger.error(result)
  return

'''
ATTACHMENTS PROCESSING
'''

'''
this function takes a pdf and pulls the data and returns the full set of lego bricks from the pdf
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
    fb = open(csvName, 'r')
    reader = csv.DictReader(fb)
    for line in reader:
        line[pieceType] = int(line['pieces'])
        if pieceType != 'pieces':
          del line['pieces']
        legos.append(line)
    return legos


'''
Main loop
'''
def handler(event, context):
    logger.info("Starting Lego Parser")
    '''bring in config'''
    logger.info("Reading Config")
    #exec(compile(open("legoParser.conf").read(), "legoParser.conf", 'exec'), locals())
    load_dotenv()
    sheetID = os.getenv('SHEETID')
    sheet_type = os.getenv('TYPE') #sets/elements
    setTemplate = os.getenv('SETTEMPLATE')
    ssWorkspace = os.getenv('SSWORKSPACE')
    ssSetsFolder = os.getenv('SSSETSFOLDER')
    
    ssTokenName = os.getenv('SSTOKEN_NAME')
    rebrickableAPIKeyName = os.getenv('REBRICKABLEAPIKEY_NAME')
    
    countLimit = os.getenv('COUNTLIMIT')
    
    logLevel = os.getenv('LOGLEVEL')
    smartsheetDown = bool(os.getenv('SMARTSHEETDOWN'))
    smartsheetUp = bool(os.getenv('SMARTSHEETUP'))

    logger.info("====ENVIRONMENT====")
    for k, v in sorted(os.environ.items()):
        logger.info(k+':', v)

    ssToken = aws.get_ssm_parameter(ssTokenName)
    rebrickableAPIKey = aws.get_parameter(rebrickableAPIKeyName)

    if logLevel == "DEBUG":
      logger.setLevel(logging.DEBUG)
    if logLevel == "INFO":
      logger.setLevel(logging.INFO)

    
    if countLimit:
      logger.info("Limiiting row count to "+ countLimit)
    ss = smartsheet(ssToken)
    logger.debug(ss.listWebhooks())
    #sheets ={'Individuals': {'id': elementsID, 'type': 'elements'} }
    sheet_proc(ss, {'id': sheetID, 'type': sheet_type},rebrickableAPIKey,smartsheetDown,smartsheetUp)
       
#handler(none,none)
    
#    sheet_proc(ss,sheetID,sheetType)
###########################################################
#    logger.info("Proccessing Elements sheet")
#    logger.info("Downloading the Sheet")
#    elements = ss.getSheet(elementsID)
#    if debug == 'smartsheet':
#        logger.debug(elements)
#        input("Press Enter to continue...")
#
#    '''build list of columns'''
#    logger.info("Getting Element Columns")
#    elementColumns = getColumns(elements)
#    if debug == 'smartsheet':
#        logger.debug(elementColumns)
#        input("Press Enter to continue...")
#
#    ssLegos = getSSLegos(elements,elementColumns,False)
#    fullDataOld = legoDetail(ssLegos,elementColumns,rebrickableAPIKey)
#    ssDataOld = prepData(fullDataOld,elementColumns)
#    resultOld = ss.updateRows(elementsID,ssDataOld)
#    elements = ss.getSheet(elementsID)
#    ssLegos = getSSLegos(elements,elementColumns,True)
#    i = 0
#    for lego in ssLegos:
#        if 'url' in lego: # and a > 12:
#            logger.debug(lego)
#            image,size = getLegoImage(lego['url'])
#            if image:
#                logger.info(f"Uploading Image with size of {size}")
#                results = ss.addCellImage(elementsID,lego,elementColumns,image,size)
#        i += 1
