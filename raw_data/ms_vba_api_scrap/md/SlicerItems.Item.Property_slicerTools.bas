Attribute VB_Name = "SlicerTools"
'slicerTools.bas
'Collection of Subroutines and UDFs for querying and printing Slicer information.
'********************************************************************************************************************************************
'--------------------------------------------------------------------------------------------------------------------------------------------
'-------------------------------------------- Module Settings and Global Variables ----------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------------------------------------
'********************************************************************************************************************************************
Option Explicit
'******************************************************************************************************************
'------------------------------------------------------------------------------------------------------------------
'-------------------------------------------------- GetSelectedSlicerItems ----------------------------------------
'------------------------------------------------------------------------------------------------------------------
'******************************************************************************************************************
'UDFs.bas contains this UDF and provides user access to it
'------------------------------------------------------------------------------------------------------------------
'https://answers.microsoft.com/en-us/msoffice/forum/msoffice_excel-mso_winother/display-slicer-selections/9fcaa6fa-9f21-4197-924b-35487b6b8f82
'------------------------------------------------------------------------------------------------------------------
'******************************************************************************************************************
'------------------------------------------------------------------------------------------------------------------
'---------------------------------------- Output Selected Slicer Values -------------------------------------------
'------------------------------------------------------------------------------------------------------------------
'******************************************************************************************************************
Public Function GetSelectedSlicerItems(slicerName As String) As String
    Dim oSc As SlicerCache
    Dim oSi As SlicerItem
    Dim lCt As Long
    On Error Resume Next
    Application.Volatile
    Set oSc = ThisWorkbook.SlicerCaches(slicerName)
    If Not oSc Is Nothing Then
        For Each oSi In oSc.SlicerItems
            If oSi.Selected Then
                GetSelectedSlicerItems = GetSelectedSlicerItems & oSi.Name & ", "
                lCt = lCt + 1
            ElseIf oSi.HasData = False Then
                lCt = lCt + 1
            End If
        Next
        If Len(GetSelectedSlicerItems) > 0 Then
            If lCt = oSc.SlicerItems.count Then
                GetSelectedSlicerItems = "All Items"
            Else
                GetSelectedSlicerItems = Left(GetSelectedSlicerItems, Len(GetSelectedSlicerItems) - 2)
            End If
        Else
            GetSelectedSlicerItems = "No items selected"
        End If
    Else
        GetSelectedSlicerItems = "No slicer with name '" & slicerName & "' was found"
    End If
End Function
''******************************************************************************************************************
'------------------------------------------------------------------------------------------------------------------
'-------------------------------------------------- SlicerEnum ----------------------------------------------------
'------------------------------------------------------------------------------------------------------------------
'******************************************************************************************************************
Public Function SlicerEnum()
'    Dim i As Integer
'    'Slicer_Equipment_Category
'    'Slicer_Test
'    'Worksheets("Sheet11").Activate
'    'With ActiveWorkbook.SlicerCaches("Slicer_Equipment_Category")
''    With ThisWorkbook.SlicerCaches("Slicer_Equipment_Category")
''      For i = 1 To ThisWorkbook.SlicerCaches("Slicer_Equipment_Category").SlicerItems.Count
''      'For i = 1 To ActiveWorkbook.SlicerCaches("Slicer_Equipment_Category").VisibleSlicerItemsList.
''      'For i = 1 To ActiveWorkbook.SlicerCaches("Slicer_Equipment_Category").VisibleSlicerItems.Item.Value
''       ''Range("B" & i) = .SlicerItems(i).Name '& " " & .SlicerItems(i).Selected & " " & .SlicerItems(i).Value & " " & .SlicerItems(i).SourceName & " " & .SlicerItems(i).Creator & " " & .SlicerItems(i).HasData & " " & .SlicerItems(i).SourceNameStandard & " " & .SlicerItems(i).Parent & " " & .SlicerItems(i).Caption
''       Range("A1") = .SlicerItems(i).Name '& " " & .SlicerItems(i).Value & " " & .SlicerItems(i).SourceName & " " & .SlicerItems(i).Creator & " " & .SlicerItems(i).HasData & " " & .SlicerItems(i).SourceNameStandard & " " & .SlicerItems(i).Parent & " " & .SlicerItems(i).Caption
''       'Range("A" & i).Value = .SlicerItems(i).Name '& " " & .SlicerItems(i).Value & " " & .SlicerItems(i).SourceName & " " & .SlicerItems(i).Creator & " " & .SlicerItems(i).HasData & " " & .SlicerItems(i).SourceNameStandard & " " & .SlicerItems(i).Parent & " " & .SlicerItems(i).Caption
''       'Range("B" & i) = .SlicerItems(i).Selected '& " " & .SlicerItems(i).Value & " " & .SlicerItems(i).SourceName & " " & .SlicerItems(i).Creator & " " & .SlicerItems(i).HasData & " " & .SlicerItems(i).SourceNameStandard & " " & .SlicerItems(i).Parent & " " & .SlicerItems(i).Caption
''       'Range("C" & i) = "Value=" & .SlicerItems(i).Value '& " " & .SlicerItems(i).Value & " " & .SlicerItems(i).SourceName & " " & .SlicerItems(i).Creator & " " & .SlicerItems(i).HasData & " " & .SlicerItems(i).SourceNameStandard & " " & .SlicerItems(i).Parent & " " & .SlicerItems(i).Caption
''       'Range("D" & i) = ActiveWorkbook.SlicerCaches("Slicer_Equipment_Category").VisibleSlicerItems.Count
''       ''Range("D" & i) = ActiveWorkbook.SlicerCaches("Slicer_Equipment_Category").VisibleSlicerItemsList
''       ''Range("B" & i) = .Slicer  Items(i).Application.PivotTableSelection
''      'MsgBox ActiveWorkbook.SlicerCaches("Slicer_Equipment_Category").VisibleSlicerItemsList
''      'MsgBox .SlicerItems(i).Name
''    Next
''    End With
'    'SlicerEnum
'    'Worksheets(11).Cells(1, 1).Value = ThisWorkbook.SlicerCaches("Slicer_Equipment_Category").SlicerItems.SlicerItems(1).Name
'    'ActiveWorksheet.Cells(1, 1).Value = ThisWorkbook.SlicerCaches("Slicer_Equipment_Category").SlicerItems.SlicerItems(1).Name
'    'ActiveSheet.Cells(2, 2).Value = ThisWorkbook.SlicerCaches("Slicer_Equipment_Category").SlicerItems.SlicerItems(1).Name
'    'ActiveSheet.Cells(2, 2) = ThisWorkbook.SlicerCaches("Slicer_Equipment_Category").SlicerItems.SlicerItems(1).Name
'    'SlicerEnum = "Jello"
'    'Range("D10") = ActiveWorkbook.SlicerCaches("Slicer_Equipment_Category").SlicerItems(2).Name & ActiveWorkbook.SlicerCaches("Slicer_Equipment_Category").SlicerItems(2).Selected
'    'ActiveSheet.Cells(2, 2) = ActiveWorkbook.SlicerCaches("Slicer_Equipment_Category").SlicerItems(2).Name & ActiveWorkbook.SlicerCaches("Slicer_Equipment_Category").SlicerItems(2).Selected
'    SlicerEnum = ActiveWorkbook.SlicerCaches("Slicer_Equipment_Category").SlicerItems(ActiveCell.Row).Value & ActiveWorkbook.SlicerCaches("Slicer_Equipment_Category").SlicerItems(ActiveCell.Row).Selected
'    'MsgBox ActiveWorkbook.SlicerCaches("Slicer_Equipment_Category").VisibleSlicerItemsList
'    'SlicerEnum = ThisWorkbook.SlicerCaches("Slicer_Equipment_Category").SlicerItems.Count
'    'Range("$B$5").Value = ThisWorkbook.SlicerCaches("Slicer_Equipment_Category").SlicerItems.SlicerItems(1).Name '& " " & .SlicerItems(i).Value & " " & .SlicerItems(i).SourceName & " " & .SlicerItems(i).Creator & " " & .SlicerItems(i).HasData & " " & .SlicerItems(i).SourceNameStandard & " " & .SlicerItems(i).Parent & " " & .SlicerItems(i).Caption
    Evaluate ("ListSlicerItems()")
End Function
'******************************************************************************************************************
'------------------------------------------------------------------------------------------------------------------
'--------------------------------------------- ListSlicerItems ----------------------------------------------------
'------------------------------------------------------------------------------------------------------------------
'******************************************************************************************************************
'ListSlicerItems.bas contains this Subroutine and provides user access to it
'------------------------------------------------------------------------------------------------------------------
'https://chandoo.org/forum/threads/extract-list-slicer-values-using-vba.21226/
'------------------------------------------------------------------------------------------------------------------
Sub ListSlicerItems()
        Dim i As Integer
        'Slicer_Equipment_Category
        'Slicer_Test
    With ActiveWorkbook.SlicerCaches("Slicer_Equipment_Category")
                For i = 1 To ActiveWorkbook.SlicerCaches("Slicer_Equipment_Category").SlicerItems.count
                'For i = 1 To ActiveWorkbook.SlicerCaches("Slicer_Equipment_Category").VisibleSlicerItemsList.
                'For i = 1 To ActiveWorkbook.SlicerCaches("Slicer_Equipment_Category").VisibleSlicerItems.Item.Value
                'Range("B" & i) = .SlicerItems(i).Name '& " " & .SlicerItems(i).Selected & " " & .SlicerItems(i).Value & " " & .SlicerItems(i).SourceName & " " & .SlicerItems(i).Creator & " " & .SlicerItems(i).HasData & " " & .SlicerItems(i).SourceNameStandard & " " & .SlicerItems(i).Parent & " " & .SlicerItems(i).Caption
                Range("A" & i) = .SlicerItems(i).Name '& " " & .SlicerItems(i).Value & " " & .SlicerItems(i).SourceName & " " & .SlicerItems(i).Creator & " " & .SlicerItems(i).HasData & " " & .SlicerItems(i).SourceNameStandard & " " & .SlicerItems(i).Parent & " " & .SlicerItems(i).Caption
                Range("B" & i) = .SlicerItems(i).Selected '& " " & .SlicerItems(i).Value & " " & .SlicerItems(i).SourceName & " " & .SlicerItems(i).Creator & " " & .SlicerItems(i).HasData & " " & .SlicerItems(i).SourceNameStandard & " " & .SlicerItems(i).Parent & " " & .SlicerItems(i).Caption
                Range("C" & i) = "Value=" & .SlicerItems(i).value '& " " & .SlicerItems(i).Value & " " & .SlicerItems(i).SourceName & " " & .SlicerItems(i).Creator & " " & .SlicerItems(i).HasData & " " & .SlicerItems(i).SourceNameStandard & " " & .SlicerItems(i).Parent & " " & .SlicerItems(i).Caption
                Range("D" & i) = ActiveWorkbook.SlicerCaches("Slicer_Equipment_Category").VisibleSlicerItems.count
       'Range("D" & i) = ActiveWorkbook.SlicerCaches("Slicer_Equipment_Category").VisibleSlicerItemsList
       'Range("B" & i) = .SlicerItems(i).Application.PivotTableSelection
                'MsgBox ActiveWorkbook.SlicerCaches("Slicer_Equipment_Category").VisibleSlicerItemsList
                'MsgBox .SlicerItems(i).Name
    Next
    End With
End Sub
'******************************************************************************************************************
'------------------------------------------------------------------------------------------------------------------
'----------------------------------------- SelectedSlicers/RevisedSlicers -----------------------------------------
'------------------------------------------------------------------------------------------------------------------
'******************************************************************************************************************
'https://www.mrexcel.com/forum/excel-questions/563503-possible-list-selected-slicer-values-cell.html
'------------------------------------------------------------------------------------------------------------------
'Dim cache As Excel.SlicerCache
'Set cache = ActiveWorkbook.SlicerCaches("Slicer_Sales_Type")
'Dim sItem As Excel.SlicerItem
'For Each sItem In cache.SlicerItems
'MsgBox sItem.Name
'Next sItem
'------------------------------------------------------------------------------------------------------------------
'SelectedSlicers.bas contains this Subroutine and provides user access to it
'------------------------------------------------------------------------------------------------------------------
Sub SelectedSlicers()
    Dim cache As Excel.SlicerCache
    Set cache = ActiveWorkbook.SlicerCaches("Slicer_Sales_Type")
    Dim sItem As Excel.SlicerItem
    For Each sItem In cache.SlicerItems
        If sItem.Selected = True Then MsgBox sItem.Name
    Next sItem
End Sub
'------------------------------------------------------------------------------------------------------------------
Sub RevisedSlicers()
    Dim sAvail As Long
    Dim sSelect As Long
    Dim cache As Excel.SlicerCache
    Set cache = ActiveWorkbook.SlicerCaches("Slicer_TYPE")
    Dim sItem As Excel.SlicerItem
    sAvail = 0
    sSelect = 0
    For Each sItem In cache.SlicerItems
    sAvail = sAvail + 1
    If sItem.Selected = True Then Range("B5000").End(xlUp).Offset(1, 0).value = sItem.value
    sSelect = sSelect + 1
    Next sItem
    If Range("B20").value < sAvail Then
    Exit Sub
    End If
    Range("B22:B5000").ClearContents
End Sub
'******************************************************************************************************************
'------------------------------------------------------------------------------------------------------------------
'---------------------------------------- Example Subroutines/UDF Functions ---------------------------------------
'------------------------------------------------------------------------------------------------------------------
'******************************************************************************************************************
''------------------------------------------------------------------------------------------------------------------
''https://everydayexcel123.blogspot.com/
''------------------------------------------------------------------------------------------------------------------
Sub MultiplePivotSlicerCaches()
    Dim oSlicer As Slicer
    Dim oSlicercache As SlicerCache
    Dim oPT As PivotTable
    Dim oSh As Worksheet
    Dim row As Integer
    Dim ObjChart As Chart
    Dim x As String
    'Create New Worksheet & Table Structure
    Sheets.Add.Name = "HIDDENSlicerConnections"
    Sheets("HIDDENSlicerConnections").Cells(1, 1).value = "Slicer Name"
    Sheets("HIDDENSlicerConnections").Cells(1, 2).value = "Pivot Parent Name"
    Sheets("HIDDENSlicerConnections").Cells(1, 3).value = "Pivot Name"
    Sheets("HIDDENSlicerConnections").Cells(1, 4).value = "Chart Title"
    'Add Slicer Cache Name and Pivot Table Name
    row = 2
' '   For Each oSlicercache In ThisWorkbook.SlicerCaches
'        'For Each oPT In oSlicercache.PivotTables
'        '    oPT.Parent.Activate
' '           Sheets("HIDDENSlicerConnections").Cells(Row, 1).Value = oSlicercache.Name
'        '    Sheets("HIDDENSlicerConnections").Cells(Row, 2).Value = oPT.Parent.Name
'        '    Sheets("HIDDENSlicerConnections").Cells(Row, 3).Value = oPT.Name
'
'        '    If InStr(oPT.Name, "Chart") < 1 Then GoTo Skip
'        '    Sheets("HIDDENSlicerConnections").Cells(Row, 4).Value = oPT.PivotChart.Chart.ChartTitle.Caption
'
''Skip:
'            Row = Row + 1
'        'Next
''    Next
'
End Sub
'********************************************************************************************************************************************
'--------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------------- Notes ----------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------------------------------------
'********************************************************************************************************************************************


