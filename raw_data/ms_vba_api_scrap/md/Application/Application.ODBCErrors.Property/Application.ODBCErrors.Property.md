# Application.ODBCErrors property (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
With Worksheets(1).QueryTables(1) 
 .Refresh 
 Set errs = Application.ODBCErrors 
 If errs.Count > 0 Then 
 Set r = .Destination.Cells(1) 
 r.Value = "The following errors occurred:" 
 c = 0 
 For Each er In errs 
 c = c + 1 
 r.offset(c, 0).value = er.ErrorString 
 r.offset(c, 1).value = er.SqlState 
 Next 
 Else 
 MsgBox "Query complete: all records returned." 
 End If 
End With
```

## Remarks
If there's more than one query running at the same time, the ODBCErrors collection contains the ODBC errors from the query that finished last.

## Example
```vba
With Worksheets(1).QueryTables(1) 
 .Refresh 
 Set errs = Application.ODBCErrors 
 If errs.Count > 0 Then 
 Set r = .Destination.Cells(1) 
 r.Value = "The following errors occurred:" 
 c = 0 
 For Each er In errs 
 c = c + 1 
 r.offset(c, 0).value = er.ErrorString 
 r.offset(c, 1).value = er.SqlState 
 Next 
 Else 
 MsgBox "Query complete: all records returned." 
 End If 
End With
```

