VERSION 5.00
Begin VB.MDIForm frmMain 
   BackColor       =   &H80000001&
   Caption         =   "Picture Viewer"
   ClientHeight    =   6285
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9300
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   230
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   9270
      TabIndex        =   0
      Top             =   6060
      Width           =   9300
      Begin VB.Label xyaxis 
         Alignment       =   1  'Right Justify
         Caption         =   "(0,0)"
         Height          =   240
         Left            =   7680
         TabIndex        =   1
         Top             =   0
         Width           =   1575
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^C
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHowTo 
         Caption         =   "How to use"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
         Shortcut        =   ^E
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub MDIForm_Load()
    frmMain.WindowState = vbMaximized
    XRes = Screen.Width / Screen.TwipsPerPixelX
    YRes = Screen.Height / Screen.TwipsPerPixelY
End Sub
Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    xyaxis.Caption = "(" & x & "," & y & ")"
End Sub
Private Sub MDIForm_Resize()
    xyaxis.Left = (frmMain.Width - xyaxis.Width - 150)
End Sub

Private Sub mnuAbout_Click()
    Load frmAbout
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuClose_Click()
    Unload frmMain.ActiveForm
End Sub

Private Sub mnuExit_Click()
    Dim Title, Message, Buttons, Reply
    Title = "Quit"
    Message = "Are you sure you want to quit?"
    Buttons = vbYesNo
    Reply = MsgBox(Message, Buttons, Title)
    If Reply = vbYes Then End
End Sub

Private Sub mnuHowTo_Click()
    Load frmHowTo
    frmHowTo.Show vbModal, Me
End Sub

Private Sub mnuOpen_Click()
    Load frmLoadPic
    frmLoadPic.Show vbModal, Me
End Sub
