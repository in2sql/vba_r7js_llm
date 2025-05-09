VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TitleScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private WithEvents Button As CommandButton
Attribute Button.VB_VarHelpID = -1
Private WithEvents Clock As Timer
Attribute Clock.VB_VarHelpID = -1
Private Chart As PictureBox


Private Palette(1 To 27) As Long
Private Head As Tripod


Public Sub Setup()
Dim i, CTLS
  ' Grab control references
  Set CTLS = Main.Controls
  Set Button = CTLS("BTN1")
  Set Clock = CTLS("TMR1")
  Set Chart = CTLS("PIC1")
  Chart.Visible = True
  CTLS("BTN2").Visible = False
  ' Build color palette
  For i = 0 To 26
    Palette(i + 1) = CTLS("PAL1").Point(i, 0)
  Next
  ' Build title screen
  Title
End Sub

Public Sub Title()
  ' Title screen graphics
  Set Main.Picture = Nothing
  Main.Move (Screen.Width - 8000) / 2, (Screen.Height - 8000) / 2, 8000, 6000
  Main.BackColor = vbApplicationWorkspace
  Main.Font.Name = "Tahoma"
  Main.Font.Size = 36
  Main.Font.Bold = True
  Main.ForeColor = vb3DDKShadow
  Main.PSet (390, 0), vbBlack
  Main.Print "PROGNOSTICATOR!"
  Main.ForeColor = vb3DHighlight
  Main.PSet (360, -30), vbBlack
  Main.Print "PROGNOSTICATOR!"
  Chart.Move 270, 870, 7400, 2000
  Chart.BackColor = vbActiveTitleBar
  Chart.Scale (-1, 1200)-(30, -80)
  With Button.Font
    .Name = "Tahoma"
    .Size = 10
    .Bold = False
  End With
  Button.Caption = "Start"
  Button.Move 3600, 5000, 990, 360
  Main.Font.Size = 16
  CText 3000, "BUSINESS CYCLES GO UP AND DOWN"
  CText 3330, "CAUSING STOCK PRICES TO RISE AND FALL"
  Main.Font.Size = 12
  Main.Font.Bold = False
  CText 4060, "Buy low and sell high to earn the most money"
  CText 4400, "See how much money you can make before the clock runs out"
End Sub

Public Sub Execute()
Dim i
  ' Build stock entities
  Set Head = New Tripod
  Head.Target = 1000
  Head.Value = 20
  Head.Build 2
  ' Advance from initial state
  Head.Target = 10000
  For i = 1 To 50
    Head.Tick (Rnd < 0.3)
  Next
  ' Set up timer
  Clock.Interval = 40
  Clock.Enabled = True
End Sub

Sub CText(ByVal Y&, TXT As String)
  ' Center text on form
  Main.PSet ((Main.ScaleWidth - Main.TextWidth(TXT)) / 2, Y), Main.BackColor
  Main.Print TXT
End Sub

Private Sub Button_Click()
  Clock.Enabled = False
  Main.StateChange New Game
End Sub

Private Sub Clock_Timer()
Dim V As New Collection
Dim i
  ' Occasional target adjustments
  If (Int(Rnd * 100) Mod 7) = 0 Then
    Head.Target = 10000
    Head.Tick True
    Head.Tick True
  ElseIf (Int(Rnd * 100) Mod 17) = 0 Then
    Head.Target = -2000
    Head.Tick True
  ElseIf (Int(Rnd * 100) Mod 5) = 0 Then
    Head.Target = Head.Target + 100
    Head.Tick True
  Else
    Head.Tick False
  End If
  ' Get stock values
  Head.Tally V
  ' Display stock values
  For i = 1 To 27
    Chart.Line (i, -20)-Step(1, 20 + V(i)), Palette(i), BF
    Chart.Line (i, V(i))-(i + 1, Chart.ScaleTop), vbActiveTitleBar, BF
  Next
End Sub
