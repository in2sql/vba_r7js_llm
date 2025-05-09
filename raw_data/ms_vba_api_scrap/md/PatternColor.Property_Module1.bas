Attribute VB_Name = "Module1"
Sub ScopeTable(ips As String)
    ' Split the IPs into an array based on line breaks
    Dim ipArray() As String
    ipArray = Split(ips, vbCrLf)

    ' Calculate the number of rows needed (add 1 for the title row)
    Dim numRows As Integer
    numRows = ((UBound(ipArray) + 1) + 3) \ 4

    ' Insert the table
    Dim tbl As Table
    Set tbl = ActiveDocument.Tables.Add(Range:=Selection.Range, numRows:=numRows + 1, NumColumns:=4)

    ' Format the title row
    tbl.Cell(1, 1).Merge MergeTo:=tbl.Cell(1, 4)
    tbl.Cell(1, 1).Range.Text = "Scope"
    tbl.Cell(1, 1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    tbl.Cell(1, 1).VerticalAlignment = wdCellAlignVerticalCenter
    tbl.Cell(1, 1).Range.Font.Bold = True
    tbl.Cell(1, 1).Range.Font.Color = wdColorWhite
    tbl.Cell(1, 1).Shading.BackgroundPatternColor = RGB(89, 89, 89) ' #595959

    ' Format the table (customize this part as needed)
    tbl.Borders.Enable = True
    tbl.Borders.OutsideColor = wdColorWhite
    tbl.Borders.InsideColor = wdColorWhite

    ' Set row height and allow rows to break across pages
    Dim r As row
    For Each r In tbl.Rows
        r.Height = CentimetersToPoints(0.7)
        r.AllowBreakAcrossPages = True
        r.Alignment = wdAlignRowCenter
    Next r

    ' Fill the table with IPs
    Dim i As Integer
    Dim row As Integer
    Dim col As Integer
    row = 2 ' Start from the second row, as the first row is the title
    col = 1
    For i = LBound(ipArray) To UBound(ipArray)
        If Trim(ipArray(i)) <> "" Then ' Ignore empty lines
            tbl.Cell(row, col).Range.Text = Trim(ipArray(i))
            tbl.Cell(row, col).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
            tbl.Cell(row, col).VerticalAlignment = wdCellAlignVerticalCenter
            tbl.Cell(row, col).Shading.BackgroundPatternColor = RGB(242, 242, 242) ' #F2F2F2
            col = col + 1
            If col > 4 Then
                col = 1
                row = row + 1
            End If
        End If
    Next i

    ' Fill any remaining cells with the desired color
    For i = row To numRows + 1
        For j = 1 To 4
            tbl.Cell(i, j).Shading.BackgroundPatternColor = RGB(242, 242, 242) ' #F2F2F2
            tbl.Cell(i, j).VerticalAlignment = wdCellAlignVerticalCenter
        Next j
    Next i
End Sub






Sub ShowIPForm()
    UserForm1.Show
End Sub
Private Sub CommandButton1_Click()
    Debug.Print TextBox1.Text ' This will print the TextBox content to the Immediate Window
    ScopeTable TextBox1.Text
    Unload Me
End Sub

