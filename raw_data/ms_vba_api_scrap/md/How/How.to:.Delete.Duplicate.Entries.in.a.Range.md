# How to: Delete Duplicate Entries in a Range

## Business Description
The following example shows how to take a range of data in column A and delete duplicate entries. This example uses the AdvancedFilter method of the Range object with the Unique parameter equal to True to get the unique list of data.

## Behavior
The following example shows how to take a range of data in column A and delete duplicate entries. This example uses theAdvancedFiltermethod of theRangeobject with theUniqueparameter equal toTrueto get the unique list of data. TheActionparameter equalsxlFilterInPlace, specifying that the data is filtered in place. If you want to retain your original data, set theActionparameter equal toxlFilterCopyand specify the location where you want the filtered data copied in theCopyToRangeparameter. Once the unique values are filtered, this example uses theSpecialCellsmethod of theRangeobject to find any remaining blank rows and deletes them.

## Example Usage
```vba
Sub DeleteDuplicates()
    With Application
        ' Turn off screen updating to increase performance
        .ScreenUpdating = False
        Dim LastColumn As Integer
        LastColumn = Cells.Find(What:="*", After:=Range("A1"), SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column + 1
        With Range("A1:A" & Cells(Rows.Count, 1).End(xlUp).Row)
            ' Use AdvanceFilter to filter unique values
            .AdvancedFilter Action:=xlFilterInPlace, Unique:=True
            .SpecialCells(xlCellTypeVisible).Offset(0, LastColumn - 1).Value = 1
            On Error Resume Next
            ActiveSheet.ShowAllData
            'Delete the blank rows
            Columns(LastColumn).SpecialCells(xlCellTypeBlanks).EntireRow.Delete
            Err.Clear
        End With
        Columns(LastColumn).Clear
        .ScreenUpdating = True
    End With
End Sub
```