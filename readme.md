<h3>Install</h3> *venv doesnt seem to work*
  virtualenv \<env\><br />
  source ./\<env\>/bin/activate<br />
  pip install -r requirements.txt

<h3>legoPaser.conf options</h3>
  **sheetID**
      This is the sheetID were it will grab the attachments

  **setTemplate**
      This is the sheet that will be coppied to be used for new inventory sheets
  
  **ssToken**
      This will be your smartsheet API token

  **ssWorkspace**
      This is the workspace that will be search for an existing set sheet, and will be used to place new set sheets
  
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


<h3> ToDo</h3>
 - add requests debugging to the output for debug=requests
 - add switch for weather or not to download the images
 - add retry logic to api/web calls
   - perhaps a sing function to make web calls used by api specific functions
 - add exception handling/logging to pdf/set name parsing
 - add logic to remove "delete" rows, and add in the set name in row1 desc column, and setID in row1 Id column
    - idea: when looping though original rows create an array of rowIds that contain the word "delete" in the row for use at the end of processing
         this prevents us from looping the whole sheet trying to delete the rows every time even after their are gone.
 - workspace change management
   - one script that takes arguments?

 - update expected column formula to include the missing column
 - Move Desc colunm to first position
 - Put "Summary:" into the Desc column of the first row
 - going through attachments even though nothing to do??
