2022-08-19
------
Added DotEnv support instead of loading config.
Added function to smartsheet.py to list webhooks
Added Docker file for build image for running the parser

Chnaged Logging to stdout
Changed `smartsheet.smartsheetRequest endpointID to optional` to support Webhook list

2021-05-26
------
Bugs:
  - fixed smartsheetRequest indentation error

Features:
 - Get set details from rebrickable based on setnumber
   - added function getSets to get set details from rebrickable
   - added Release column to column ids to added release year
   - when processing sheet for rows to process added subsequent list of sets missing info (Photo or Desc)
