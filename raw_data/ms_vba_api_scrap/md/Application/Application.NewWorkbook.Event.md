# Range Object (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Remarks
The following properties and methods for returning a Range object are described in the examples section:

## Example
```vba
Sub SetUpTable() 
Worksheets("Sheet1").Activate 
For TheYear = 1 To 5 
    Cells(1, TheYear + 1).Value = 1990 + TheYear 
Next TheYear 
For TheQuarter = 1 To 4 
    Cells(TheQuarter + 1, 1).Value = "Q" & TheQuarter 
Next TheQuarter 
End Sub
```

```vba
Dim r1 As Range, r2 As Range, myMultiAreaRange As Range 
Worksheets("sheet1").Activate 
Set r1 = Range("A1:B2") 
Set r2 = Range("C3:D4") 
Set myMultiAreaRange = Union(r1, r2) 
myMultiAreaRange.Select
```

```vba
Sub NoMultiAreaSelection() 
    NumberOfSelectedAreas = Selection.Areas.Count 
    If NumberOfSelectedAreas > 1 Then 
        MsgBox "You cannot carry out this command " & _ 
            "on multi-area selections" 
    End If 
End Sub
```

```vba
Sub Create_Unique_List_Count()
    'Excel workbook, the source and target worksheets, and the source and target ranges.
    Dim wbBook As Workbook
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim rnSource As Range
    Dim rnTarget As Range
    Dim rnUnique As Range
    'Variant to hold the unique data
    Dim vaUnique As Variant
    'Number of unique values in the data
    Dim lnCount As Long
    
    'Initialize the Excel objects
    Set wbBook = ThisWorkbook
    With wbBook
        Set wsSource = .Worksheets("Sheet1")
        Set wsTarget = .Worksheets("Sheet2")
    End With
    
    'On the source worksheet, set the range to the data stored in column A
    With wsSource
        Set rnSource = .Range(.Range("A1"), .Range("A100").End(xlUp))
    End With
    
    'On the target worksheet, set the range as column A.
    Set rnTarget = wsTarget.Range("A1")
    
    'Use AdvancedFilter to copy the data from the source to the target,
    'while filtering for duplicate values.
    rnSource.AdvancedFilter Action:=xlFilterCopy, _
                            CopyToRange:=rnTarget, _
                            Unique:=True
                            
    'On the target worksheet, set the unique range on Column A, excluding the first cell
    '(which will contain the "List" header for the column).
    With wsTarget
        Set rnUnique = .Range(.Range("A2"), .Range("A100").End(xlUp))
    End With
    
    'Assign all the values of the Unique range into the Unique variant.
    vaUnique = rnUnique.Value
    
    'Count the number of occurrences of every unique value in the source data,
    'and list it next to its relevant value.
    For lnCount = 1 To UBound(vaUnique)
        rnUnique(lnCount, 1).Offset(0, 1).Value = _
            Application.Evaluate("COUNTIF(" & _
            rnSource.Address(External:=True) & _
            ",""" & rnUnique(lnCount, 1).Text & """)")
    Next lnCount
    
    'Label the column of occurrences with "Occurrences"
    With rnTarget.Offset(0, 1)
        .Value = "Occurrences"
        .Font.Bold = True
    End With

End Sub
```

