# AboveAverage.Creator property (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
Sub FindCreator() 
 
 Dim myObject As Excel.Workbook 
 Set myObject = ActiveWorkbook 
 If myObject.Creator = &h5843454c Then 
 MsgBox "This is a Microsoft Excel object." 
 Else 
 MsgBox "This is not a Microsoft Excel object." 
 End If 
 
End Sub
```

## Remarks
If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The Creator property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.

## Example
```vba
Sub FindCreator() 
 
 Dim myObject As Excel.Workbook 
 Set myObject = ActiveWorkbook 
 If myObject.Creator = &h5843454c Then 
 MsgBox "This is a Microsoft Excel object." 
 Else 
 MsgBox "This is not a Microsoft Excel object." 
 End If 
 
End Sub
```

