VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub btnClear_Click()
    'Clear cell content and formating
    Range("A19:AC100").Clear
    Range("A19:AC100").ClearFormats
    Range("L1:AC18").Clear
    Range("L1:AC18").ClearFormats
    
    'Remove all previous shapes
    Dim Shp As Shape
    For Each Shp In ActiveSheet.Shapes
        If Shp.Type = msoAutoShape Then Shp.Delete
    Next Shp
    
    'Reset row height and column width
    Rows.UseStandardHeight = True
    Columns.UseStandardWidth = True
End Sub

Private Sub btnFillTree_Click()
    If Range("start").Interior.Color <> 255 Then
        MsgBox "Please generate the tree before trying to fill it."
        Exit Sub
    End If
    
    Call FillTree
End Sub

Private Sub BtnGenerate_Click()
    'Clear cell content and formating
    Range("A19:AC100").Clear
    Range("A19:AC100").ClearFormats
    Range("L1:AC18").Clear
    Range("L1:AC18").ClearFormats
    
    'Remove all previous shapes
    Dim Shp As Shape
    For Each Shp In ActiveSheet.Shapes
        If Shp.Type = msoAutoShape Then Shp.Delete
    Next Shp
    
    'Reset row height and column width
    Rows.UseStandardHeight = True
    Columns.UseStandardWidth = True
    
    Application.Wait Now + #12:00:01 AM# '1 sec delay
    
    For Each nm In ThisWorkbook.Names
        If Left(nm.Name, 5) <> "start" Then GoTo Skip1
        With Range(nm)
            .Merge
            '.Value = 5
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Size = 16
            .Font.Bold = True
            .Interior.Color = RGB(255, 0, 0)
        End With
        
Skip1:
    Next nm
    
    Call DrawArrows(Range("start"), Range("yesExp"))
    Call DrawArrows(Range("start"), Range("noExp"))
    
    Application.Wait Now + #12:00:01 AM#
    
    For Each nm In ThisWorkbook.Names
        If (Left(nm.Name, 3) <> "yes" And Left(nm.Name, 2) <> "no") Then GoTo Skip2
        With Range(nm)
            .Merge
            '.Value = 5
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Size = 16
            .Font.Bold = True
            .Interior.Color = RGB(131, 186, 106)
        End With
        
Skip2:
    Next nm
        
    Call DrawArrows(Range("yesExp"), Range("yesExp2"))
    Call DrawArrows(Range("noExp"), Range("noExp2"))
    
    Call DrawArrows(Range("yesExp2"), Range("exper1"))
    Call DrawArrows(Range("yesExp2"), Range("exper2"))
    Call DrawArrows(Range("yesExp2"), Range("exper3"))
    
    Call DrawArrows(Range("noExp2"), Range("exper0"))
    
    For Each nm In ThisWorkbook.Names
        If (Left(nm.Name, 5) <> "exper") Then GoTo Skip3
        With Range(nm)
            .Merge
            '.Value = 5
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Size = 16
            .Font.Bold = True
            .Interior.Color = RGB(255, 138, 130)
        End With
        
Skip3:
    Next nm
    
    Call DrawArrows(Range("exper1"), Range("alter11"))
    Call DrawArrows(Range("exper1"), Range("alter12"))
    
    Call DrawArrows(Range("exper2"), Range("alter21"))
    Call DrawArrows(Range("exper2"), Range("alter22"))
    
    Call DrawArrows(Range("exper3"), Range("alter31"))
    Call DrawArrows(Range("exper3"), Range("alter32"))
    
    Call DrawArrows(Range("exper0"), Range("alter01"))
    Call DrawArrows(Range("exper0"), Range("alter02"))
    
    Application.Wait Now + #12:00:01 AM#
      
    For Each nm In ThisWorkbook.Names
        If (Left(nm.Name, 5) <> "alter") Then GoTo Skip4
        With Range(nm)
            .Merge
            '.Value = 5
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Size = 16
            .Font.Bold = True
            .Interior.Color = RGB(255, 228, 25)
        End With
        
Skip4:
    Next nm
    
    Call DrawArrows(Range("alter11"), Range("sast111"))
    Call DrawArrows(Range("alter11"), Range("sast112"))
    Call DrawArrows(Range("alter11"), Range("sast113"))
    Call DrawArrows(Range("alter11"), Range("sast114"))
    
    Call DrawArrows(Range("alter12"), Range("sast121"))
    Call DrawArrows(Range("alter12"), Range("sast122"))
    Call DrawArrows(Range("alter12"), Range("sast123"))
    Call DrawArrows(Range("alter12"), Range("sast124"))
    
    Call DrawArrows(Range("alter21"), Range("sast211"))
    Call DrawArrows(Range("alter21"), Range("sast212"))
    Call DrawArrows(Range("alter21"), Range("sast213"))
    Call DrawArrows(Range("alter21"), Range("sast214"))
    
    Call DrawArrows(Range("alter22"), Range("sast221"))
    Call DrawArrows(Range("alter22"), Range("sast222"))
    Call DrawArrows(Range("alter22"), Range("sast223"))
    Call DrawArrows(Range("alter22"), Range("sast224"))
    
    Call DrawArrows(Range("alter31"), Range("sast311"))
    Call DrawArrows(Range("alter31"), Range("sast312"))
    Call DrawArrows(Range("alter31"), Range("sast313"))
    Call DrawArrows(Range("alter31"), Range("sast314"))
    
    Call DrawArrows(Range("alter32"), Range("sast321"))
    Call DrawArrows(Range("alter32"), Range("sast322"))
    Call DrawArrows(Range("alter32"), Range("sast323"))
    Call DrawArrows(Range("alter32"), Range("sast324"))
    
    Call DrawArrows(Range("alter01"), Range("sast011"))
    Call DrawArrows(Range("alter01"), Range("sast012"))
    Call DrawArrows(Range("alter01"), Range("sast013"))
    Call DrawArrows(Range("alter01"), Range("sast014"))
    
    Call DrawArrows(Range("alter02"), Range("sast021"))
    Call DrawArrows(Range("alter02"), Range("sast022"))
    Call DrawArrows(Range("alter02"), Range("sast023"))
    Call DrawArrows(Range("alter02"), Range("sast024"))
    
    For Each nm In ThisWorkbook.Names
        If (Left(nm.Name, 4) <> "sast") Then GoTo Skip5
        With Range(nm)
            '.Merge
            '.Value = 5
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Size = 16
            .Font.Bold = True
            .Cells(1).Interior.Color = RGB(56, 158, 209)
        End With
        
Skip5:
    Next nm
    
    For Each nm In ThisWorkbook.Names
        If (Right(nm.Name, 5) <> "Label") Then GoTo Skip10
        With Range(nm)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Size = 16
            .Font.Bold = True
        End With
        
Skip10:
    Next nm
    
    Range("V1").ColumnWidth = Range("V1").ColumnWidth * 2
    
End Sub

Private Sub FillTree()
       
    For Each nm In ThisWorkbook.Names
    If (Left(nm.Name, 4) <> "sast") Then GoTo Skip9
        With Range(nm)
            .Cells(2).Value = "I" + Right(nm.Name, 1)
            .Cells(2).HorizontalAlignment = xlCenter
        End With
Skip9:
    Next nm

    Dim i As Integer
    i = 8
    
    For Each nm In ThisWorkbook.Names
        If (Left(nm.Name, 4) <> "sast" Or Mid(nm.Name, 6, 1) <> 1) Then GoTo Skip6
        Range(nm).Cells(1).Formula = "=B" + CStr(i)
        If i = 11 Then i = 8 Else i = i + 1
Skip6:
    Next nm
    
    For Each nm In ThisWorkbook.Names
        If (Left(nm.Name, 4) <> "sast" Or Mid(nm.Name, 6, 1) <> 2 Or Mid(nm.Name, 5, 1) <> 0) Then GoTo Skip7
        Range(nm).Cells(1).Formula = "=C" + CStr(i)
        If i = 11 Then i = 8 Else i = i + 1
Skip7:
    Next nm
    
    For Each nm In ThisWorkbook.Names
        If (Left(nm.Name, 4) <> "sast" Or Mid(nm.Name, 6, 1) <> 2 Or Mid(nm.Name, 5, 1) = 0) Then GoTo Skip8
        Range(nm).Cells(1).Value = "0"
Skip8:
    Next nm
    
    Application.Wait (Now + TimeValue("0:00:01"))
    
    Range("alter11").Formula = "=SUMPRODUCT(" + Range("e1_").Address + "," + Range("groupSast11").Address + ")"
    Range("alter12").Formula = "=SUMPRODUCT(" + Range("e1_").Address + "," + Range("groupSast12").Address + ")"
    Range("alter21").Formula = "=SUMPRODUCT(" + Range("e2_").Address + "," + Range("groupSast21").Address + ")"
    Range("alter22").Formula = "=SUMPRODUCT(" + Range("e2_").Address + "," + Range("groupSast22").Address + ")"
    Range("alter31").Formula = "=SUMPRODUCT(" + Range("e3_").Address + "," + Range("groupSast31").Address + ")"
    Range("alter32").Formula = "=SUMPRODUCT(" + Range("e3_").Address + "," + Range("groupSast32").Address + ")"
    Range("alter01").Formula = "=SUMPRODUCT(" + Range("pi_").Address + "," + Range("groupSast01").Address + ")"
    Range("alter02").Formula = "=SUMPRODUCT(" + Range("pi_").Address + "," + Range("groupSast02").Address + ")"
    
    Application.Wait (Now + TimeValue("0:00:01"))
    
    Range("exper1").Formula = "=MAX(" + Range("alter11").Address + "," + Range("alter12").Address + ")"
    Range("exper2").Formula = "=MAX(" + Range("alter21").Address + "," + Range("alter22").Address + ")"
    Range("exper3").Formula = "=MAX(" + Range("alter31").Address + "," + Range("alter32").Address + ")"
    Range("exper0").Formula = "=MAX(" + Range("alter01").Address + "," + Range("alter02").Address + ")"
    
    Application.Wait (Now + TimeValue("0:00:01"))
    
    Range("noExp2").Value = "=" + Range("exper0").Cells(1).Address
    Range("noExp").Value = "=" + Range("noExp2").Cells(1).Address
    
    Range("yesExp2").Formula = "=" + Range("pe_").Cells(1).Address + "*" + Range("exper1").Cells(1).Address + "+" _
                                    + Range("pe_").Cells(2).Address + "*" + Range("exper2").Cells(1).Address + "+" _
                                    + Range("pe_").Cells(3).Address + "*" + Range("exper3").Cells(1).Address
    
    Range("yesExp").Formula = "=" + Range("yesExp2").Cells(1).Address + "-" + Range("K4").Address
    
    Application.Wait (Now + TimeValue("0:00:01"))
    
    Range("start").Formula = "=MAX(" + Range("yesExp").Cells(1).Address + "," + Range("noExp").Cells(1).Address + ")"
    
    Application.Wait (Now + TimeValue("0:00:01"))
    Application.Wait (Now + TimeValue("0:00:01"))
    
    Range("keLabel").Value = "K(e)"
    Range("ke0Label").Value = "K(e) = 0"
    Range("e1Label").Value = "e1"
    Range("e2Label").Value = "e2"
    Range("e3Label").Value = "e3"
    Range("e0Label").Value = "e0"
    Range("v11Label").Value = "V1"
    Range("v21Label").Value = "V2"
    Range("v12Label").Value = "V1"
    Range("v22Label").Value = "V2"
    Range("v13Label").Value = "V1"
    Range("v23Label").Value = "V2"
    Range("v10Label").Value = "V1"
    Range("v20Label").Value = "V2"
       
End Sub

Private Sub DrawArrows(FromRange As Range, ToRange As Range, Optional RGBcolor As Long, Optional LineType As String)

Dim dleft1 As Double, dleft2 As Double
Dim dtop1 As Double, dtop2 As Double
Dim dheight1 As Double, dheight2 As Double
Dim dwidth1 As Double, dwidth2 As Double
dleft1 = FromRange.Left
dleft2 = ToRange.Left
dtop1 = FromRange.Top
dtop2 = ToRange.Top
dheight1 = FromRange.Height
dheight2 = ToRange.Height
dwidth1 = FromRange.Width
dwidth2 = ToRange.Width
 
ActiveSheet.Shapes.AddConnector(msoConnectorStraight, dleft1 + dwidth1, dtop1 + dheight1 / 2, dleft2, dtop2 + dheight2 / 2).Select
'format line
With Selection.ShapeRange.Line
    .BeginArrowheadStyle = msoArrowheadNone
    .EndArrowheadStyle = msoArrowheadOpen
    .Weight = 1.75
    .Transparency = 0
    If UCase(LineType) = "DOUBLE" Then 'double arrows
        .BeginArrowheadStyle = msoArrowheadOpen
    ElseIf UCase(LineType) = "LINE" Then 'Line (no arows)
        .EndArrowheadStyle = msoArrowheadNone
    Else 'single arrow
        'defaults to an arrow with one head
    End If
    'color arrow
    If RGBcolor <> 0 Then
        .ForeColor.RGB = RGBcolor 'custom color
    Else
        .ForeColor.RGB = RGB(0, 31, 156) '(228, 108, 10) 'orange (DEFAULT)
    End If
End With
 
End Sub


Private Sub btnTracePath_Click()
    If (Range("start").Interior.Color <> 255 Or Range("keLabel") <> "K(e)") Then
        MsgBox "Please generate and fill the tree before trying to trace the path."
        Exit Sub
    End If
    
    'Remove all previous shapes
    Dim Shp As Shape
    For Each Shp In ActiveSheet.Shapes
        If Shp.AutoShapeType = msoShapeNoSymbol Or Shp.AutoShapeType = msoShapeChevron Then
            Shp.Delete
        End If
    Next Shp
    
    Dim yesExp As Range, noExp, alter11, alter12, alter21, alter22, alter31, alter32, alter01, alter02 As Range
    Set yesExp = Range("D24")
    Set noExp = Range("D34")
    Set alter11 = Range("Q7")
    Set alter12 = Range("Q10")
    Set alter21 = Range("Q18")
    Set alter22 = Range("Q21")
    Set alter31 = Range("Q29")
    Set alter32 = Range("Q32")
    Set alter01 = Range("Q40")
    Set alter02 = Range("Q43")
    
    Application.Wait Now + #12:00:01 AM#

    If Range("yesExp").Cells(1).Value = Range("start").Cells(1).Value Then
        Shapes.AddShape msoShapeNoSymbol, noExp.Left, noExp.Top, Application.CentimetersToPoints(2.87), Application.CentimetersToPoints(2.61)
        Shapes.AddShape msoShapeChevron, yesExp.Left, yesExp.Top, Application.CentimetersToPoints(2.57), Application.CentimetersToPoints(2.73)
    Else
        Shapes.AddShape msoShapeNoSymbol, yesExp.Left, yesExp.Top, Application.CentimetersToPoints(2.87), Application.CentimetersToPoints(2.61)
        Shapes.AddShape msoShapeChevron, noExp.Left, noExp.Top, Application.CentimetersToPoints(2.57), Application.CentimetersToPoints(2.73)
    End If
    
    ShapeFormat
    Application.Wait Now + #12:00:01 AM#
    
    If Range("alter11").Cells(1).Value = Range("exper1").Cells(1).Value Then
        Shapes.AddShape msoShapeNoSymbol, alter12.Left, alter12.Top, Application.CentimetersToPoints(1.32), Application.CentimetersToPoints(1.21)
        Shapes.AddShape msoShapeChevron, alter11.Left, alter11.Top, Application.CentimetersToPoints(1.23), Application.CentimetersToPoints(1.28)
    Else
        Shapes.AddShape msoShapeNoSymbol, alter11.Left, alter11.Top, Application.CentimetersToPoints(1.32), Application.CentimetersToPoints(1.21)
        Shapes.AddShape msoShapeChevron, alter12.Left, alter12.Top, Application.CentimetersToPoints(1.23), Application.CentimetersToPoints(1.28)
    End If
    
    ShapeFormat
    Application.Wait Now + #12:00:01 AM#
    
    If Range("alter21").Cells(1).Value = Range("exper2").Cells(1).Value Then
        Shapes.AddShape msoShapeNoSymbol, alter22.Left, alter22.Top, Application.CentimetersToPoints(1.32), Application.CentimetersToPoints(1.21)
        Shapes.AddShape msoShapeChevron, alter21.Left, alter21.Top, Application.CentimetersToPoints(1.23), Application.CentimetersToPoints(1.28)
    Else
        Shapes.AddShape msoShapeNoSymbol, alter21.Left, slter21.Top, Application.CentimetersToPoints(1.32), Application.CentimetersToPoints(1.21)
        Shapes.AddShape msoShapeChevron, alter22.Left, alter22.Top, Application.CentimetersToPoints(1.23), Application.CentimetersToPoints(1.28)
    End If
    
    ShapeFormat
    Application.Wait Now + #12:00:01 AM#
    
    If Range("alter31").Cells(1).Value = Range("exper3").Cells(1).Value Then
        Shapes.AddShape msoShapeNoSymbol, alter32.Left, alter32.Top, Application.CentimetersToPoints(1.32), Application.CentimetersToPoints(1.21)
        Shapes.AddShape msoShapeChevron, alter31.Left, alter31.Top, Application.CentimetersToPoints(1.23), Application.CentimetersToPoints(1.28)
    Else
        Shapes.AddShape msoShapeNoSymbol, alter31.Left, alter31.Top, Application.CentimetersToPoints(1.32), Application.CentimetersToPoints(1.21)
        Shapes.AddShape msoShapeChevron, alter32.Left, alter32.Top, Application.CentimetersToPoints(1.23), Application.CentimetersToPoints(1.28)
    End If
    
    ShapeFormat
    Application.Wait (Now + TimeValue("0:00:01"))
    ShapeFormat
    Application.Wait (Now + TimeValue("0:00:01"))
    
    If Range("alter01").Cells(1).Value = Range("exper0").Cells(1).Value Then
        Shapes.AddShape msoShapeNoSymbol, alter02.Left, alter02.Top, Application.CentimetersToPoints(1.32), Application.CentimetersToPoints(1.21)
        Shapes.AddShape msoShapeChevron, alter01.Left, alter01.Top, Application.CentimetersToPoints(1.23), Application.CentimetersToPoints(1.28)
    Else
        Shapes.AddShape msoShapeNoSymbol, alter01.Left, alter01.Top, Application.CentimetersToPoints(1.32), Application.CentimetersToPoints(1.21)
        Shapes.AddShape msoShapeChevron, alter02.Left, alter02.Top, Application.CentimetersToPoints(1.23), Application.CentimetersToPoints(1.28)
    End If
    
    ShapeFormat
    

    
End Sub

Private Sub ShapeFormat()
        Dim ShpNo As Shape
        For Each ShpNo In ActiveSheet.Shapes
            If ShpNo.AutoShapeType = msoShapeNoSymbol Then
                ShpNo.Fill.ForeColor.RGB = RGB(252, 3, 3)
            End If
        Next ShpNo
    
        Dim ShpYes As Shape
        For Each ShpYes In ActiveSheet.Shapes
            If ShpYes.AutoShapeType = msoShapeChevron Then
                ShpYes.Rotation = 90
                ShpYes.Adjustments.Item(1) = 0.6
                ShpYes.Fill.ForeColor.RGB = RGB(24, 105, 63)
            End If
    Next ShpYes
End Sub
