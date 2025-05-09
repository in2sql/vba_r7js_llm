# PivotCache RobustConnect Property

## Business Description
Returns or sets how the PivotTable cache connects to its data source. Read/write XlRobustConnect.

## Behavior
Returns or sets how the PivotTable cache connects to its data source. Read/writeXlRobustConnect.

## Example Usage
```vba
Sub CheckRobustConnect() 
 
 Dim pvtCache As PivotCache 
 
 Set pvtCache = Application.ActiveWorkbook.PivotCaches.Item(1) 
 
 ' Determine the connection robustness and notify user. 
 Select Case pvtCache.RobustConnectCase xlAlways 
 MsgBox "The PivotTable cache is always connected to its source." 
 Case xlAsRequired 
 MsgBox "The PivotTable cache is connected to its source as required." 
 Case xlNever 
 MsgBox "The PivotTable cache is never connected to its source." 
 End Select 
 
End Sub
```