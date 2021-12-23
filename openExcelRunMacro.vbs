Option Explicit

Dim xlApp, xlBook

Set xlApp = CreateObject("Excel.Application")

'~~> Change Path here
custFilepath = "\\netfolder\net subfolder\file.xlsb")

Set xlBook = xlApp.Workbooks.Open(custFilepath, 0, False)

'delay of 6 seconds = sleep 6000
WScript.Sleep 6000
xlbook.refreshall
xlapp.run "macroName"

xlBook.Close false   'or true if you want to save the workbook on close
xlApp.Quit

Set xlBook = Nothing
Set xlApp = Nothing



WScript.Quit
