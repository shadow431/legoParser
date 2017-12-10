<h3>Install</h3> *venv doesnt seem to work*
  virtualenv \<env\><br />
  source ./\<env\>/bin/activate<br />
  pip install -r requirements.txt

<h3>legoPaser.conf options</h3>
  **sheetID**
      This is the sheetID were it will grab the attachments, and upload the parsed data
  
  **ssToken**
      This will be your smartsheet API token
  
  **countLimit**
      activates a count in the row attachment processor to force a premeture completion after x lego sets
  
  **debug**
      use for see the data at various steps.
      Current options are 'pdf, smartsheet, approve'
        pdf: prints out the data as it is pulled and manaipulated,
        smartsheet: prints the data as it is retrieved, prepared, and submited for/to smartsheet,
        approve: itterates through each component of the menus so you can see if it is parsing right.
        requests: prints out the response from smartsheet for row inserts
  
  **smartsheetDown**
      Boolean used to set whether to pull data from smartsheet for processing
      if False it will download the sheet and process the row/attachemnt info, but will use the existing pdf document.
      Note: if Flase I recommend setting countLimit = True
  
  **smartsheetUP**
      Boolean used to set whether or not to upload the data once processed back up to smartsheet
