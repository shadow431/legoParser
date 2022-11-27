import logging, requests, json

class smartsheet:


  def __init__ (self,ssToken):
    self.ssToken = ssToken
    self.logger = logging.getLogger('legoparser.smartsheet')

  def smartsheetRequest(self,endpoint,endpointID,data=None,method='GET',action=None,headers={}):
        headers['Authorization'] = f'Bearer {self.ssToken}'
        url = f'https://api.smartsheet.com/2.0/{endpoint}'
        if endpointID:
          url += f'/{endpointID}'
        self.logger.debug(url)
        if action:
          url += action
        if method == 'GET':
          r = requests.get(url, headers=headers)
        elif method == 'POST':
          r = requests.post(url, data=data, headers=headers)
        elif method == 'PUT':
          r = requests.put(url, data=data, headers=headers)
        rArr = r.json()
        return rArr

  def getSheet(self,sheetID):
      return self.smartsheetRequest('sheets',sheetID)
  
  def copySheet(self,sheetID,data):
      jsonData = json.dumps(data)
      return self.smartsheetRequest('sheets',sheetID,data=jsonData,method='POST',action='/copy?include=data')
  
  def getWorkspace(self,workspaceID):
      return self.smartsheetRequest('workspaces',workspaceID)

  def getAttachments(self,sheetID,row_id=None):
      action='/attachments?includeAll=True'
      if row_id:
        action = f'/rows/{row_id}{action}'
      return self.smartsheetRequest('sheets',sheetID,action=action)
  
  def getAttachment(self,sheetID,attachmentID):
      return self.smartsheetRequest('sheets',sheetID,action=f'/attachments/{attachmentID}')
  
  def listWebhooks(self):
      return self.smartsheetRequest('webhooks',None)

  def insertRows(self,sheetId,data):
      jsonData = json.dumps(data)
      return self.smartsheetRequest('sheets',sheetId,action='/rows',data=jsonData,method='POST')
   
  def updateRows(self,sheetId,data):
      data = json.dumps(data)
      self.logger.debug(data)
      return self.smartsheetRequest('sheets',sheetId,action='/rows',method='PUT',data=data)
  
  def addCellImage(self,sheetID,lego,columnId,image,imageSize):
      self.logger.info(lego)
      headers = {}
      headers['Content-Type'] = "application/jpeg"
      headers['Content-Disposition'] = 'attachment; filename="%s.jpg"'%lego['id']
      headers['Content-Length'] = str(imageSize)
      return self.smartsheetRequest('sheets',sheetID,action=f"/rows/{lego['row']}/columns/{columnId['picture']}/cellimages?altText={lego['set_img_url']}",headers=headers,method='POST',data=image)
