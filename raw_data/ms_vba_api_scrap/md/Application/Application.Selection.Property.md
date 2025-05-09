# Application.Selection property (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
Worksheets("Sheet1").Activate 
Selection.Clear
```

## Remarks
The returned object type depends on the current selection (for example, if a cell is selected, this property returns a Range object). The Selection property returns Nothing if nothing is selected.

## Example
```vba
Sub TestSelection(  )
    Dim str As String
    Select Case TypeName(Selection)
    Case "Nothing"
        str = "No selection made."
    Case "Range"
        str = "You selected the range: " & Selection.Address
    Case "Picture"
        str = "You selected a picture."
    Case Else
        str = "You selected a " & TypeName(Selection) & "."
    End Select
    MsgBox str
End Sub
```

