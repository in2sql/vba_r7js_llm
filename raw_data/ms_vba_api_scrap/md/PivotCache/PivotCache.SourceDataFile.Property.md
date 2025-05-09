# PivotCache SourceDataFile Property

## Business Description
Returns a String value that indicates the source data file for the cache of the PivotTable.

## Behavior
Returns aStringvalue that indicates the source data file for the cache of the PivotTable.

## Example Usage
```vba
Sub CheckSourceConnection() 
 
 Dim pvtCache As PivotCache 
 
 Set pvtCache = Application.ActiveWorkbook.PivotCaches.Item(1) 
 
 On Error GoTo No_Connection 
 
 MsgBox "The data source connection is: " & _ 
 pvtCache.SourceDataFileExit Sub 
 
No_Connection: 
 MsgBox "PivotCache source cannot be determined." 
 
End Sub
```