channelNumber = Application.DDEInitiate( _ 
 app:="WinWord", _ 
 topic:="C:\WINWORD\SALES.DOC") 
Set rangeToPoke = Worksheets("Sheet1").Range("A1") 
Application.DDEPoke channelNumber, "\StartOfDoc", rangeToPoke 
Application.DDETerminate channelNumber