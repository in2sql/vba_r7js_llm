# PivotCache SourceConnectionFile Property

## Business Description
Returns or sets a String indicating the Microsoft Office Data Connection file or similar file that was used to create the PivotTable. Read/write.

## Behavior
Returns or sets aStringindicating the Microsoft Office Data Connection file or similar file that was used to create the PivotTable. Read/write.

## Example Usage
```vba
Sub CheckSourceConnection() 
 
 Dim pvtCache As PivotCache 
 
 Set pvtCache = Application.ActiveWorkbook.PivotCaches.Item(1) 
 
 On Error GoTo No_Connection 
 
 MsgBox "The source connection is: " & pvtCache.SourceConnectionFileExit Sub 
 
No_Connection: 
 MsgBox "PivotCache source can not be determined." 
 
End Sub
```