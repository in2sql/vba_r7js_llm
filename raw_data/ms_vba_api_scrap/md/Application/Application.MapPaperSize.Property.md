# Application.MapPaperSize property (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
Sub UseMapPaperSize() 
 
 ' Determine setting and notify user. 
 If Application.MapPaperSize = True Then 
 MsgBox "Microsoft Excel automatically " & _ 
 "adjusts the paper size according to the country/region setting." 
 Else 
 MsgBox "Microsoft Excel does not " & _ 
 "automatically adjusts the paper size according to the country/region setting." 
 End If 
 
End Sub
```

## Example
```vba
Sub UseMapPaperSize() 
 
 ' Determine setting and notify user. 
 If Application.MapPaperSize = True Then 
 MsgBox "Microsoft Excel automatically " & _ 
 "adjusts the paper size according to the country/region setting." 
 Else 
 MsgBox "Microsoft Excel does not " & _ 
 "automatically adjusts the paper size according to the country/region setting." 
 End If 
 
End Sub
```

