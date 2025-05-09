VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.MDIForm frmOldMain 
   BackColor       =   &H8000000C&
   Caption         =   "EDM"
   ClientHeight    =   7200
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10710
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox desk 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000C&
      Height          =   4335
      Left            =   0
      ScaleHeight     =   4275
      ScaleWidth      =   10650
      TabIndex        =   0
      Top             =   0
      Width           =   10710
   End
   Begin MSCommLib.MSComm theoport 
      Left            =   960
      Top             =   6480
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   360
      Top             =   6480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu filemenu 
      Caption         =   "&File"
      Begin VB.Menu filenewsite 
         Caption         =   "&New site"
         Shortcut        =   ^N
      End
      Begin VB.Menu filenewpoints 
         Caption         =   "New &points file"
      End
      Begin VB.Menu donothing1 
         Caption         =   "-"
      End
      Begin VB.Menu fileopensite 
         Caption         =   "Open site"
         Shortcut        =   ^O
      End
      Begin VB.Menu fileopenpoints 
         Caption         =   "&Open points file"
      End
      Begin VB.Menu donothing3 
         Caption         =   "-"
      End
      Begin VB.Menu fileimport 
         Caption         =   "&Import"
      End
      Begin VB.Menu fileexport 
         Caption         =   "&Export"
      End
      Begin VB.Menu donothing4 
         Caption         =   "-"
      End
      Begin VB.Menu filelist 
         Caption         =   "p1"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu filelist 
         Caption         =   "p2"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu filelist 
         Caption         =   "p3"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu filelist 
         Caption         =   "p4"
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu filelist 
         Caption         =   "p5"
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu donothing5 
         Caption         =   "-"
      End
      Begin VB.Menu fileexit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu stationmenu 
      Caption         =   "&Setup"
      Begin VB.Menu stationinitialize 
         Caption         =   "&Station"
      End
      Begin VB.Menu setupunits 
         Caption         =   "&Units"
      End
      Begin VB.Menu setupprinter 
         Caption         =   "&Printer"
      End
      Begin VB.Menu setuptheodolite 
         Caption         =   "&Theodolite"
      End
      Begin VB.Menu setupfields 
         Caption         =   "&Fields"
      End
   End
   Begin VB.Menu recordpoints 
      Caption         =   "&Record"
      Begin VB.Menu recordmenu 
         Caption         =   "Artifact"
         Index           =   0
         Shortcut        =   ^A
      End
      Begin VB.Menu recordmenu 
         Caption         =   "Topography"
         Index           =   1
         Shortcut        =   ^T
      End
      Begin VB.Menu recordmenu 
         Caption         =   "Sample"
         Index           =   2
         Shortcut        =   ^S
      End
      Begin VB.Menu recordmenu 
         Caption         =   "Bucket"
         Index           =   3
         Shortcut        =   ^B
      End
      Begin VB.Menu recordpoint 
         Caption         =   "X-shot"
         Index           =   4
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu viewmenu 
      Caption         =   "&View"
      Begin VB.Menu viewplan 
         Caption         =   "&Plan (X-Y)"
      End
      Begin VB.Menu viewsagital 
         Caption         =   "&Sagital (Y-Z)"
      End
      Begin VB.Menu viewfrontal 
         Caption         =   "&Frontal (X-Z)"
      End
      Begin VB.Menu viewnothing 
         Caption         =   "-"
      End
      Begin VB.Menu viewstatus 
         Caption         =   "&Status"
      End
   End
   Begin VB.Menu windowmenu 
      Caption         =   "&Window"
      Begin VB.Menu windowpsheet 
         Caption         =   "&Points"
      End
      Begin VB.Menu windowpmap 
         Caption         =   "Point &map"
      End
      Begin VB.Menu windowdsheet 
         Caption         =   "Datums"
      End
      Begin VB.Menu windowdmap 
         Caption         =   "&Datum map"
      End
      Begin VB.Menu polesheet 
         Caption         =   "Poles"
      End
      Begin VB.Menu unitsheet 
         Caption         =   "Units"
      End
      Begin VB.Menu donothing2 
         Caption         =   "-"
      End
      Begin VB.Menu windowtile 
         Caption         =   "Tile"
      End
   End
   Begin VB.Menu helpmenu 
      Caption         =   "&Help"
      Begin VB.Menu helpabout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmOldMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub MDIForm_Load()

'Public variables initialized
BannerHeight = 400
BannerWidth = 150

inifile$ = fixpath(App.Path) + "edm.ini"
Call ReadEDMini(inifile$)
desk.Height = Me.Height

Me.Show

EDMName$ = readcfg("theodolite")
If EDMName$ = "" Then
    EDMName$ = "None"
    Call writecfg("theodolite", "None")
End If
comport$ = readcfg("comport")

temp$ = readcfg("vhd")
If temp$ = "0" Or temp$ = "" Then
    vhd = False
Else
    vhd = True
End If

CurrentStation.name = readcfg("stationname")
CurrentStation.x = Val(readcfg("lastx"))
CurrentStation.y = Val(readcfg("lasty"))
CurrentStation.z = Val(readcfg("lastz"))

'need code here to deal with initializing the edm
'depending on which one they selected

Call printstatus

End Sub

Private Sub viewstatus_Click()

Me.desk.Cls
Me.desk.Print "EDM for Windows - Version 1.0"
Me.desk.Print "by Shannon McPherron and Harold Dibble"
Me.desk.Print
Me.desk.Print "Total Station : " + EDMName$
Me.desk.Print "Current site : " + SiteDBname$
Me.desk.Print "Current points file : " + PointTableName$
Me.desk.Print "Station name : " + CurrentStation.name
Me.desk.Print "Station X : " + Str$(CurrentStation.x)
Me.desk.Print "Station Y : " + Str$(CurrentStation.y)
Me.desk.Print "Station Z : " + Str$(CurrentStation.z)

End Sub

Private Sub windowdmap_Click()

frmDatummap.Show

End Sub

Private Sub windowpmap_Click()

frmPointmap.Show

End Sub

Sub printstatus()

Me.desk.Print "EDM for Windows - Version 1.0"
Me.desk.Print "by Shannon McPherron and Harold Dibble"
Me.desk.Print
Me.desk.Print "Total Station : " + EDMName$
Me.desk.Print "Current site : " + SiteDBname$
Me.desk.Print "Current points file : " + PointTableName$
Me.desk.Print "Station name : " + CurrentStation.name
Me.desk.Print "Station X : " + Str$(CurrentStation.x)
Me.desk.Print "Station Y : " + Str$(CurrentStation.y)
Me.desk.Print "Station Z : " + Str$(CurrentStation.z)
        
End Sub

