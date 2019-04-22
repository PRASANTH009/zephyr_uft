'Create QTP object
Set QTP = CreateObject("QuickTest.Application")
QTP.Launch
QTP.Visible = TRUE
 
'Open QTP Test
QTP.Open "D:\DevAppPro\Test123\Sample", TRUE 'Set the QTP test path

'Run QTP test
QTP.Test.Run Action1
  
'Close QTP
QTP.Test.Close
QTP.Quit