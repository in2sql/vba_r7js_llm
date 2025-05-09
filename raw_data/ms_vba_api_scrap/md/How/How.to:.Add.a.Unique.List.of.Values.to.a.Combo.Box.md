# How to: Add a Unique List of Values to a Combo Box

## Business Description
These examples show different approaches for taking a list from a spreadsheet and using it to populate a combo box control using only the unique values.

## Behavior
These examples show different approaches for taking a list from a spreadsheet and using it to populate a combo box control using only the unique values. The first example uses theAdvancedFiltermethod of the Range object and the second uses the Collection object.

## Example Usage
```vba
Sub Populate_Combobox_Worksheet()
    'The Excel workbook and worksheets that contain the data, as well as the range placed on that data
    Dim wbBook As Workbook
    Dim wsSheet As Worksheet
    Dim rnData As Range

    'Variant to contain the data to be placed in the combo box.
    Dim vaData As Variant

    'Initialize the Excel objects
    Set wbBook = ThisWorkbook
    Set wsSheet = wbBook.Worksheets("Sheet1")

    'Set the range equal to the data, and then (temporarily) copy the unique values of that data to the L column.
    With wsSheet
        Set rnData = .Range(.Range("A1"), .Range("A100").End(xlUp))
        rnData.AdvancedFilter Action:=xlFilterCopy, _
                          CopyToRange:=.Range("L1"), _
                          Unique:=True
        'store the unique values in vaData
        vaData = .Range(.Range("L2"), .Range("L100").End(xlUp)).Value
        'clean up the contents of the temporary data storage
        .Range(.Range("L1"), .Range("L100").End(xlUp)).ClearContents
    End With

    'display the unique values in vaData in the combo box already in existence on the worksheet.
    With wsSheet.OLEObjects("ComboBox1").Object
        .Clear
        .List = vaData
        .ListIndex = -1
    End With

End Sub
```