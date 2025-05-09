VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm Parent 
   BackColor       =   &H8000000F&
   Caption         =   " Code Library 3.0"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8100
   Icon            =   "Parent.frx":0000
   LinkTopic       =   "MDIForm1"
   NegotiateToolbars=   0   'False
   ScrollBars      =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   5940
      Width           =   8100
      _ExtentX        =   14288
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2514
            MinWidth        =   2514
            Text            =   " Press F1 for help"
            TextSave        =   " Press F1 for help"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   88194
            MinWidth        =   88194
            Text            =   "File info:"
            TextSave        =   "File info:"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Parent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'-----------
'Parent Form
'-----------

Dim Temp1 As String * 75        'Variable saved in ini file
Dim getnumPages As Long         'Return variable for .ini file procedures
Dim Y As Integer                'For/Next loop variable
Dim file As String              'Full path of .ini file

Private Sub MDIForm_Load()
  frmCodeLib.Show
  frmCodeLib.rtb1(frmCodeLib.TabStrip1.SelectedItem.Index - 1).SetFocus
  Me.Show

'--------------WINDOW STATE
  Temp1 = ""
  Ini_File
  getnumPages = GetPrivateProfileString("Preferences", "WindowState", file, Temp1, Len(Temp1), file)
  Y = InStr(Temp1, ".ini")
  
  If Y = 0 Then
    Me.WindowState = Left(Temp1, getnumPages)
  Else
    Me.WindowState = vbMaximized
  End If

End Sub

Public Sub Ini_File()
  filepath = App.Path
  
  If Right(filepath, 1) <> "\" Then
    filepath = App.Path & "\"
  End If
  
  file = filepath & "Library.ini"
End Sub

Private Sub StatusBar1_DblClick()
  
  If Me.WindowState = vbMaximized Then
    Me.WindowState = vbNormal
  Else
    Me.WindowState = vbMaximized
  End If
  
End Sub

Private Sub StatusBar1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  If Button = 2 Then
    frmCodeLib.PopUpToolBar.Bands("Toolbar").TrackPopup -1, -1
  End If
  
End Sub
