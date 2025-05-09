Public columnsQty As Integer
Public player As Integer
Public xPlayerScore, oPlayerScore As Integer

Sub Initialize()
    columnsQty = 3: player = 1: CreateTable
    xPlayerScore = 0: oPlayerScore = 0
End Sub

Private Sub Worksheet_SelectionChange(ByVal target As Range)
    'get the cell that the player selected:
    Dim selectedCell As Range: Set selectedCell = target
    Dim value As String: value = selectedCell.value
    Dim column, row, i As Integer:
    column = selectedCell.column
    row = selectedCell.row
    Dim result As Boolean 'if false = invalid
    result = CheckMove(column, row, value, target)
    If result = True Then
        'main game functions
        DrawMove selectedCell
        ChangePlayer
        CheckWin
    End If
End Sub

Function CreateTable()
    Dim i, j As Integer
    Dim cellSize As Integer: cellSize = 5
    RestoreTable
    'calculation for cell size:
    For i = 1 To columnsQty
        Columns(i).ColumnWidth = cellSize
        Rows(i).RowHeight = cellSize * 6.666
    Next: FormatTable: FormatMenu
End Function

Function RestoreTable()
    Application.ScreenUpdating = False
    Dim i, j As Integer
    For i = 1 To 10
        For j = 1 To 10
            Cells(i, j).ClearFormats
            Cells(i, j).ColumnWidth = 8.43: Cells(i, j).RowHeight = 15
            Cells(i, j).value = ""
    Next: Next: Application.ScreenUpdating = True
End Function

Function RestoreGame(playerString, CheckWin)
    If CheckWin = True Then
        Select Case playerString
            Case "X": MsgBox ("Player X win!")
                xPlayerScore = xPlayerScore + 1
            Case "O": MsgBox ("Player O win!")
                oPlayerScore = oPlayerScore + 1
        End Select:
    Else: MsgBox ("Draw!"): End If
RestoreTable: CreateTable
End Function

Function DrawMove(selectedCell)
    Select Case player
        Case 1: selectedCell.value = "X"
        Case 2: selectedCell.value = "O"
    End Select
End Function

Function ChangePlayer()
    Select Case player
        Case 1: player = 2
        Case 2: player = 1
    End Select
End Function

Function CheckMove(column, row, value, target) As Boolean
    If target.Cells.Count > 1 Then
        MsgBox "Do not select more than one cell"
        CheckMove = False
    ElseIf column > columnsQty Or row > columnsQty Then
        MsgBox "Out of Table"
        CheckMove = False
    ElseIf value = "X" Or value = "O" Then
        MsgBox "Place already taken"
        CheckMove = False
    '---------end of conditions---------
    Else: CheckMove = True: End If
End Function

Function CheckWin() As Boolean
    Dim playerString As String
    Select Case player
        Case 1: playerString = "O"
        Case 2: playerString = "X"
    End Select
    'if sum = columnsQty, the player will win
    Dim i, j, sum As Integer
    For i = 1 To columnsQty
        sum = 0
        'restore sum for the row
        For j = 1 To columnsQty
            If Cells(i, j).value = playerString Then
                sum = sum + 1
            End If
        Next j
        If sum = columnsQty Then
            CheckWin = True: RestoreGame playerString, CheckWin
        End If
        'restore sum for the columns
        sum = 0
        For j = 1 To columnsQty
            If Cells(j, i).value = playerString Then
                sum = sum + 1
            End If
        Next j
        If sum = columnsQty Then
            CheckWin = True: RestoreGame playerString, CheckWin
        End If
    Next i
    'restore sum for the first diagonal
    sum = 0
    For i = 1 To columnsQty
        If Cells(i, i).value = playerString Then
            sum = sum + 1
        End If
    Next i
    If sum = columnsQty Then
        CheckWin = True: RestoreGame playerString, CheckWin
    End If
    'restore sum for the second diagonal
    sum = 0
    For i = 1 To columnsQty
        If Cells(i, 4 - i).value = playerString Then
            sum = sum + 1
        End If
    Next i
    If sum = columnsQty Then
        CheckWin = True: RestoreGame playerString, CheckWin
    End If
    'check draw
    Dim drawSum As Integer
    For i = 1 To columnsQty
        For j = 1 To columnsQty
            If Cells(i, j).value <> "" Then
                drawSum = drawSum + 1
            End If
        Next j
    Next i
    Debug.Print drawSum
    'draw if all tiles filled
    If drawSum = columnsQty * columnsQty Then
        CheckWin = False: RestoreGame playerString, CheckWin
    End If
End Function

Function FormatTable()
    Dim i, j As Integer
    For i = 1 To columnsQty
        For j = 1 To columnsQty
            Cells(i, j).Interior.ColorIndex = 16
            Cells(i, j).Font.ColorIndex = 2
            Cells(i, j).Borders.LineStyle = xlContinuous
            Cells(i, j).Borders.Weight = xlMedium
            Cells(i, j).Borders.ColorIndex = 1
            Cells(i, j).HorizontalAlignment = xlCenter
            Cells(i, j).VerticalAlignment = xlCenter
            Cells(i, j).Font.Size = 20
    Next: Next
End Function

Function FormatMenu()
    Dim i, j, k As Integer: i = columnsQty + 1
    'codes:
    'i, i = player x text
    'i, i + 1 = player o text
    'i + 1, i = x score var
    'i + 1, i + 1 = o score var
    Cells(i, i).value = "Player X Score:"
    Cells(i + 1, i).value = xPlayerScore
    Cells(i, i + 1).value = "Player O Score:"
    Cells(i + 1, i + 1).value = oPlayerScore
    Cells(i, i).Columns.AutoFit
    Cells(i, i + 1).Columns.AutoFit
    For j = i To i + 1
        For k = i To i + 1
            Cells(j, k).Interior.ColorIndex = 16
            Cells(j, k).Interior.ColorIndex = 16
            Cells(j, k).Font.ColorIndex = 2
            Cells(j, k).Borders.Weight = xlMedium
            Cells(j, k).Borders.ColorIndex = 1
            Cells(j, k).HorizontalAlignment = xlCenter
            Cells(j, k).VerticalAlignment = xlCenter
            Cells(j, k).Font.Bold = True
        Next: Next
    If xPlayerScore > oPlayerScore Then
        Cells(i + 1, i).Font.ColorIndex = 4
        Cells(i + 1, i + 1).Font.ColorIndex = 3
    ElseIf xPlayerScore < oPlayerScore Then
        Cells(i + 1, i).Font.ColorIndex = 3
        Cells(i + 1, i + 1).Font.ColorIndex = 4
    End If
End Function
