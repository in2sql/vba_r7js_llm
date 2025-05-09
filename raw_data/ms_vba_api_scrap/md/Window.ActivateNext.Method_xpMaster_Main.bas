Attribute VB_Name = "Main"
'// Event routines are in the Workbook object
Option Explicit

Public ctrlXpsearch As Office.CommandBarControl

Private Const TAGXP As String = "XP"         'name of this control Addin
Private Const XPSEARCH As String = "XpSearch"           'name of operational Addin

Public Sub deleteXPcontrols()    '// delete all cutsom controls - 'XP' control
    Dim col As CommandBarControls
    Dim c

    Set col = CommandBars.FindControls(tag:=TAGXP)
    If Not col Is Nothing Then
        For Each c In col
            c.Delete
        Next c
    End If
End Sub

Public Sub xpF3()
    If Not ActiveWindow Is Nothing Then ActiveWindow.ActivateNext
End Sub

Public Sub xpF6()   '// Toggle AutoFilter, TopRow, SplitWindow, Freeze panes
    If Not TypeName(ActiveSheet) = "Worksheet" Then Exit Sub
    With ActiveWindow
        Select Case False
            Case .Split
                .ScrollRow = .ActiveSheet.UsedRange.Cells(1).Row
                .SplitColumn = 0
                .SplitRow = 1
                .FreezePanes = True
            Case .ActiveSheet.AutoFilterMode Or IsEmpty(.ActiveSheet.UsedRange.Cells(1))
                .ActiveSheet.UsedRange.Cells(1).AutoFilter
            Case Else
                .ActiveSheet.AutoFilterMode = False
                .FreezePanes = False
                .Split = False
        End Select
    End With
End Sub

Public Sub xpF5()
    Select Case True
        Case Not ActiveChart Is Nothing     '// on a chart - save to png
            xpChartSavedAsPng
        Case Not ActiveCell Is Nothing      '// in a worksheet cell
            Select Case True
                Case ActiveCell.Hyperlinks.Count = 1
                    xpFollowHyperlink
                Case ActiveCell.ListObject Is Nothing
                Case ActiveCell.ListObject.AutoFilter Is Nothing
                Case Else
                    ActiveCell.ListObject.AutoFilter.ApplyFilter
            End Select
    End Select
End Sub

Public Sub xpF7()
    If TypeName(ActiveSheet) = "Worksheet" Then ActiveSheet.UsedRange.Select
End Sub

Public Sub xpF8()
    If TypeName(ActiveSheet) = "Worksheet" Then ActiveSheet.UsedRange.Select
End Sub

Public Sub xpBuiltInMenusPopup()    'F1 key pulls up All menus :)
    Application.CommandBars("Built-in Menus").ShowPopup
End Sub

Public Sub xpFollowHyperlink()  'F5 follows hyperlink in cell - Same as Regedit 'ForceShellExecute' key
    With ActiveCell
        If .Hyperlinks.Count = 1 Then
            Shell Environ("ProgramW6432") & "\Mozilla Firefox\firefox.exe " & .Hyperlinks(1).Address
            .Font.ThemeColor = xlThemeColorFollowedHyperlink
        ElseIf InStr(.Text, "linkedin.com") > 0 Then
            Shell Environ("ProgramW6432") & "\Mozilla Firefox\firefox.exe " & .Text
            .Font.ThemeColor = xlThemeColorFollowedHyperlink
        End If
    End With
''    ActiveWorkbook.FollowHyperlink "https://www.linkedin.com/in/reidhoffman/"
End Sub

Public Sub xpChartSavedAsPng()
    Dim s As String
''    Const EXT As String = "jpg"
    Const EXT As String = "png" '// 'gif'
''    Const EXT As String = "gif" '// 'gif'
    
''For Each ch In ActiveWorkbook.Charts: ch.Export Filename:=ActiveWorkbook.Path & "\" & ch.Name & ".png": Next
    s = ActiveWorkbook.Path & "\" & ActiveChart.Name & "." & EXT
    ActiveChart.Export FileName:=s, FilterName:=EXT
    MsgBox s
End Sub
