# Application.InputBox method (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
Set myRange = Application.InputBox(prompt := "Sample", type := 8)
```

## Parameters
- **Prompt**: Required
- **Title**: Optional
- **Default**: Optional
- **Left**: Optional
- **Top**: Optional
- **HelpFile**: Optional
- **HelpContextID**: Optional
- **Type**: Optional

## Return Value
Variant

## Remarks
The following table lists the values that can be passed in the Type argument. Can be one or a sum of the values. For example, for an input box that can accept both text and numbers, set Type to 1 + 2.

## Example
```vba
Set myRange = Application.InputBox(prompt := "Sample", type := 8)
```

```vba
Worksheets("Sheet1").Activate 
Set myCell = Application.InputBox( _ 
    prompt:="Select a cell", Type:=8)
```

```vba
Sub Cbm_Value_Select()
   'Set up the variables.
   Dim rng As Range
   
   'Use the InputBox dialog to set the range for MyFunction, with some simple error handling.
   Set rng = Application.InputBox("Range:", Type:=8)
   If rng.Cells.Count <> 3 Then
     MsgBox "Length, width and height are needed -" & _
         vbLf & "please select three cells!"
      Exit Sub
   End If
   
   'Call MyFunction by value using the active cell.
   ActiveCell.Value = MyFunction(rng)
End Sub

Function MyFunction(rng As Range) As Double
   MyFunction = rng(1) * rng(2) * rng(3)
End Function
```

