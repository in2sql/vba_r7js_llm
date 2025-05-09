# AboveAverage object (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Remarks
All conditional formatting objects are contained within a FormatConditions collection object, which is a child of a Range collection.

## Example
```vba
Sub AboveAverageCF() 
 
' Building data for Melanie 
 Range("A1").Value = "Name" 
 Range("B1").Value = "Number" 
 Range("A2").Value = "Melanie-1" 
 Range("A2").AutoFill Destination:=Range("A2:A26"), Type:=xlFillDefault 
 Range("B2:B26").FormulaArray = "=INT(RAND()*101)" 
 Range("B2:B26").Select 
 
' Applying Conditional Formatting to items above the average. Should appear green fill and dark green font. 
 Selection.FormatConditions.AddAboveAverage 
 Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority 
 Selection.FormatConditions(1).AboveBelow = xlAboveAverage 
 With Selection.FormatConditions(1).Font 
 .Color = -16752384 
 .TintAndShade = 0 
 End With 
 With Selection.FormatConditions(1).Interior 
 .PatternColorIndex = xlAutomatic 
 .Color = 13561798 
 .TintAndShade = 0 
 End With 
MsgBox "Added an Above Average Conditional Format to Melanie's data. Press F9 to update values.", vbInformation 
 
End Sub
```

