Attribute VB_Name = "PlantumlFTP"
Const ASCII_TRANSFER = 1
Const BINARY_TRANSFER = 2
Const INTERNET_FLAG_RELOAD = &H80000000
Const UserName = "plantuml"
Const Pass = "plantuml"
Const useFTP = True

'Open the Internet object
 Private Declare Function InternetOpen _
   Lib "wininet.dll" _
     Alias "InternetOpenA" _
       (ByVal sAgent As String, _
        ByVal lAccessType As Long, _
        ByVal sProxyName As String, _
        ByVal sProxyBypass As String, _
        ByVal lFlags As Long) As Long

'Connect to the network
 Private Declare Function InternetConnect _
   Lib "wininet.dll" _
     Alias "InternetConnectA" _
       (ByVal hInternetSession As Long, _
        ByVal sServerName As String, _
        ByVal nServerPort As Integer, _
        ByVal sUsername As String, _
        ByVal sPassword As String, _
        ByVal lService As Long, _
        ByVal lFlags As Long, _
        ByVal lContext As Long) As Long

'Get a file using FTP
 Private Declare Function FtpGetFile _
   Lib "wininet.dll" _
     Alias "FtpGetFileA" _
       (ByVal hFtpSession As Long, _
        ByVal lpszRemoteFile As String, _
        ByVal lpszNewFile As String, _
        ByVal fFailIfExists As Boolean, _
        ByVal dwFlagsAndAttributes As Long, _
        ByVal dwFlags As Long, _
        ByVal dwContext As Long) As Boolean

'Send a file using FTP
 Private Declare Function FtpPutFile _
   Lib "wininet.dll" _
     Alias "FtpPutFileA" _
       (ByVal hFtpSession As Long, _
        ByVal lpszLocalFile As String, _
        ByVal lpszRemoteFile As String, _
        ByVal dwFlags As Long, _
        ByVal dwContext As Long) As Boolean

'Close the Internet object
 Private Declare Function InternetCloseHandle _
   Lib "wininet.dll" _
     (ByVal hInet As Long) As Integer


Function testgetServerPort()
    Dim servername As String
    Dim serverport As Integer
    serverport = 123
    Debug.Assert getServerPort("127.0.0.1", servername, serverport) = True
    Debug.Assert servername = "127.0.0.1"
    Debug.Assert serverport = 123
    
    
    Debug.Assert getServerPort("127.0.0.1:4242", servername, serverport) = True
    Debug.Assert servername = "127.0.0.1"
    Debug.Assert serverport = 4242
    servername = ""
    serverport = 4242
    Debug.Assert getServerPort("127.0.0.1", servername, serverport) = True
    Debug.Assert servername = "127.0.0.1"
    Debug.Assert serverport = 4242
    
    Debug.Assert getServerPort("http://127.0.0.1:4242", servername, serverport) = False
    Debug.Assert getServerPort("www.nowhere.com:1234", servername, serverport) = True
    Debug.Assert servername = "www.nowhere.com"
    Debug.Assert serverport = 1234
    
    
End Function

Function getServerPort(url As String, ByRef servername As String, ByRef serverport As Integer) As Boolean
  Dim params() As String
  Dim RE As RegExp
  Dim match
  getServerPort = False
  Set RE = New RegExp
  url = LCase(url)
  If InStr("://", url) Then
    If Left(url, 6) = "ftp://" Then
       url = Mid(url, 7)
    Else
       Exit Function
    End If
  End If
  
  params = Split(url, ":")
  If UBound(params) <= 1 Then
     servername = params(0)
     If UBound(params) = 1 Then
        serverport = Val(params(1))
    End If
    RE.Pattern = "[^:/\\ \t\n\r\%\&]+"
    If RE.Test(servername) And ((UBound(params) = 0) Or (serverport > 0)) Then
       getServerPort = True
    End If
  End If
End Function

Function ftpOpen(FTPURL As String) As Long
  Dim INet As Long
  Dim INetConn As Long
  Dim RetVal As Long
  Dim Success As Long
  Dim servername As String
  Dim serverport As Integer
  ftpOpen = 0
  INetConn = -1
  serverport = 4242 ' default
  If getServerPort(FTPURL, servername, serverport) Then
    
    INet = InternetOpen("MyFTP Control", 1&, vbNullString, vbNullString, 0&)
    If INet > 0 Then
       INetConn = InternetConnect(INet, servername, serverport, UserName, Pass, 1&, 0&, 0&)
       ftpOpen = INetConn
      Debug.Print "FtpOpen(" & FTPURL & ") -> success"
    Else
      Debug.Print "FtpOpen(" & FTPURL & ") -> failed"
    End If
  Else
    Debug.Print "FtpOpen(" & FTPURL & ") -> ill configured server/port"
   End If
End Function

Function ftpClose(handle As Long)
  If handle > 0 Then
       RetVal = InternetCloseHandle(handle)
  End If
  Debug.Print "FtpClose(" & handle & ")"
End Function
' =========================================================
' Store a File to a FTP server
Function FtpStor(INetConn As Long, localFile, hostFile)

  Dim RetVal As Long
  Dim Success As Long
  
  RetVal = False
  FtpStor = True
    
    If INetConn > 0 Then
        Success = FtpPutFile(INetConn, localFile, hostFile, BINARY_TRANSFER Or INTERNET_FLAG_RELOAD, 0&)
       FtpStor = True
    End If
  Debug.Print "FtpStor(" & localFile & " , " & hostFile & ") -> " & RetVal & "success=" & Success

End Function


' =========================================================
' Retrieve a File from a FTP server
Function FtpRetr(INetConn As Long, localFile, hostFile)

  Dim INet As Long
  Dim RetVal As Long
  Dim Success As Long

  RetVal = False
  FtpRetr = RetVal
  If INetConn > 0 Then
      FtpRetr = True
      Success = FtpGetFile(INetConn, hostFile, localFile, False, 0, BINARY_TRANSFER Or INTERNET_FLAG_RELOAD, 0&)
      Debug.Print "FtpRetr(" & localFile & " , " & hostFile & ") -> " & Success
    
  End If

End Function




' =========================================================

Function Macro_UML(scope)
' Generate diagrams image from a PlantUML source textual description in the Word Document
' Scope can be "parg" or "all"
'
' - Initialisations
'
     Call ToolbarInit
    Set statusButton = CommandBars("UML").Controls(6)
    
    Call CreateStyle
    Call CreateStyleImg
    Call ShowPlantuml

    Call ShowHiddenText
    Selection.Range.Select
'
' documentId is the filename with its path, without extension
'
    documentId = ActiveDocument.Name
    documentId = Left(documentId, Len(documentId) - 4)
    
    ' Check for the presente of plantuml.jar
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    jarPath = getJarPath()
    If fs.FileExists(jarPath & "\plantuml.jar") = False Then
        MsgBox jarPath
        GoTo Macro_UML_exit
    End If
    
' - Phase 1
' We create a file text per bloc of diagrams
' We look for @startuml
' We open the textfile in background (visible:=false)
' We add to the name a number on 4 digit
' The text bloc is put on "PlantUML" style
' Then the bloc is copied into the text file

    statusButton.Caption = "Extract"
    statusButton.Visible = False
    statusButton.Visible = True
    If scope = "all" Then
        Set parsedtext = ActiveDocument.Content
        isForward = True
    Else
        Set parsedtext = Selection.Range
        parsedtext.Collapse
        isForward = False
    End If

    parsedtext.Find.Execute FindText:=startuml, Forward:=isForward
    If parsedtext.Find.Found = True Then
        'We keep the the first line only "@startuml" with the carriage return
        Set singleparagraph = parsedtext.Paragraphs(1).Range
        singleparagraph.Collapse
    Else
        GoTo Macro_UML_exit
    End If
    
    Do While parsedtext.Find.Found = True And _
             (scope = "all" Or currentIndex < 1)
        statusButton.Caption = "Extract." & currentIndex + 1
        statusButton.Visible = False
        statusButton.Visible = True
        Set currentparagraph = parsedtext.Paragraphs(1)
        Set paragraphRange = currentparagraph.Range
        paragraphRange.Collapse
        jobDone = False
        Do Until jobDone
            If Left(currentparagraph.Range.Text, Len(startuml)) = startuml Then
                Set paragraphRange = currentparagraph.Range
                paragraphRange.Collapse
               
            End If
            paragraphRange.MoveEnd Unit:=wdParagraph
            If Left(currentparagraph.Range.Text, Len(enduml)) = enduml Then
                paragraphRange.Style = "PlantUML"
                paragraphRange.Copy
                Set textFile = Documents.Add(Visible:=False)
                textFile.Content.Paste
                currentIndex = currentIndex + 1
                textFileId = documentId & "_extr" & Right("000" & currentIndex, 4) & ".txt"
                textFile.SaveAs filename:=jarPath & "\" & textFileId, FileFormat:=wdFormatText, Encoding:=65001
                textFile.Close
                If useFTP Then
                  retValue = FtpStor(jarPath & "\" & textFileId, textFileId)
                  'MsgBox ("A")
                  'imageId = Left(textFileId, Len(textFileId) - 4) & ".png"
                  'imageName = jarPath & "\" & imageId
                  'retValue = FtpRetr(imageName, imageId)
                  'MsgBox ("B")
                End If
                jobDone = True
            End If
            
            Set currentparagraph = currentparagraph.Next
            
            If currentparagraph Is Nothing Then
                jobDone = True
            End If
        Loop
        parsedtext.Collapse Direction:=wdCollapseEnd
        If scope = "all" Then
            parsedtext.Find.Execute FindText:=startuml, Forward:=True
        End If
   Loop
'
' We create a lock file that will be deleted by the Java program to indicate the end of Java process
'
    statusButton.Caption = "Gener"
    statusButton.Visible = False
    statusButton.Visible = True
    Set lockFile = Documents.Add(Visible:=False)
    lockFile.SaveAs filename:=jarPath & "\javaumllock.tmp", FileFormat:=wdFormatText
    lockFile.Close

'
' Call to PlantUML to generate images from text descriptions
'
If useFTP Then
  For I = 1 To currentIndex
         imageId = documentId & "_extr" & Right("000" & I, 4) & ".png"
         imageName = jarPath & "\" & imageId
         retValue = FtpRetr(imageName, imageId)
  Next I
  'Sleep 200
End If

If useFTP = False Then
    JavaCommand = "java -classpath """ & jarPath & "\plantuml.jar;" & _
            jarPath & "\plantumlskins.jar"" net.sourceforge.plantuml.Run -charset UTF8 -word """ & jarPath & "/"""
    Shell (JavaCommand)
' This sleep is needed, but we don't know why...
    Sleep 500
'
' Phase 2 :
' Insertion of images into the word document
' We insert the image after the textual block that describe the diagram
'
    jobDone = False
    currentIndex = 0
    
' We wait for the file javaumllock.tmp to be deleted by Java
' which means that the process is ended
'
    Do
        currentIndex = currentIndex + 1
        statusButton.Caption = "Gener." & currentIndex
        statusButton.Visible = False
        statusButton.Visible = True
        DoEvents
        Sleep 1000
        If fs.FileExists(jarPath & "\javaumllock.tmp") = False Then
            jobDone = True
            Exit Do
        End If
        If currentIndex > 30 Then
            statusButton.Visible = False
            MsgBox ("Java Timeout. Aborted.")
            Exit Do
        End If
    Loop
    
    If jobDone = False Then
        End
    End If
End If
        
    statusButton.Caption = "Inser"
    statusButton.Visible = False
    statusButton.Visible = True
    
    If scope = "all" Then
        Set parsedtext = ActiveDocument.Content
        isForward = True
    Else
        Set parsedtext = singleparagraph
        isForward = True
    End If
    parsedtext.Find.Execute FindText:=enduml, Forward:=isForward
    currentIndex = 0
    Do While parsedtext.Find.Found = True And (scope = "all" Or currentIndex < 1)
        currentIndex = currentIndex + 1
        statusButton.Caption = "Inser." & currentIndex
        statusButton.Visible = False
        statusButton.Visible = True
        On Error GoTo LastParagraph
        Set currentparagraph = parsedtext.Paragraphs(1).Next.Range
        Do While currentparagraph.InlineShapes.Count > 0 And currentparagraph.Style = "PlantUMLImg"
            currentparagraph.Delete
            Set currentparagraph = parsedtext.Paragraphs(1).Next.Range
        Loop
        On Error GoTo 0
        Set currentRange = currentparagraph
        imagesDirectory = jarPath & "\" & documentId & "_extr" & Right("000" & currentIndex, 4) & "*.png"
        image = Dir(imagesDirectory)
        While image <> ""
            Set currentparagraph = ActiveDocument.Paragraphs.Add(Range:=currentRange).Range
            Set currentRange = currentparagraph.Paragraphs(1).Next.Range
            currentparagraph.Style = "PlantUMLImg"
            currentparagraph.Collapse
            
            Set image = currentparagraph.InlineShapes.AddPicture _
                (filename:=jarPath & "\" & image _
                , LinkToFile:=False, SaveWithDocument:=True)
            image.AlternativeText = "Generated by PlantUML"
            If image.ScaleHeight > 100 Or image.ScaleWidth > 100 Then
                image.Reset
            End If
            image = Dir()
        Wend
        parsedtext.Collapse Direction:=wdCollapseEnd
        parsedtext.Find.Execute FindText:=enduml, Forward:=True
   Loop
    
'
' Phase 3 : suppression of temporary files (texte and PNG)
'
Phase3:
    statusButton.Caption = "Delete"
    statusButton.Visible = False
    statusButton.Visible = True
    On Error Resume Next
    Kill (jarPath & "\" & documentId & "_extr*.*")
    On Error GoTo 0

Macro_UML_exit:

    statusButton.Visible = False
    
    'We show the hidden description text
    Call ShowHiddenText
    DoubleCheckStyle
Exit Function


' This is need when the very last line of the Word document is @enduml
LastParagraph:
    Selection.EndKey Unit:=wdStory
    Selection.TypeParagraph
    Selection.ClearFormatting
    
        imagesDirectory = jarPath & "\" & documentId & "_extr" & Right("000" & currentIndex, 4) & "*.png"
        image = Dir(imagesDirectory)
        While image <> ""
            Set currentparagraph = ActiveDocument.Paragraphs.Add.Range
            Set currentRange = currentparagraph.Paragraphs(1).Next.Range
            currentparagraph.Style = "PlantUMLImg"
            currentparagraph.Collapse
            
            Set image = currentparagraph.InlineShapes.AddPicture _
                (filename:=jarPath & "\" & image _
                , LinkToFile:=False, SaveWithDocument:=True)
            image.AlternativeText = "Generated by PlantUML"
            If image.ScaleHeight > 100 Or image.ScaleWidth > 100 Then
                image.Reset
            End If
            image = Dir()
        Wend
    
    'Resume Next
    GoTo Phase3

End Function

' =========================================================
' Initialize the plantuml ToolBar
Function ToolbarInit()

    On Error GoTo ToolbarCreation
    Set toolBar = ActiveDocument.CommandBars("UML")
    On Error GoTo 0
    toolBar.Visible = True
    
    On Error GoTo ButtonAdd
    Set currentButton = toolBar.Controls(1)
    On Error GoTo 0
    currentButton.OnAction = "Module1.SwitchP"
    currentButton.Style = msoButtonCaption
    currentButton.Caption = Chr(182)
    currentButton.Visible = True
    
    On Error GoTo ButtonAdd
    Set currentButton = toolBar.Controls(2)
    On Error GoTo 0
    currentButton.OnAction = "Module1.ShowPlantuml"
    currentButton.Style = msoButtonCaption
    currentButton.Caption = "Show PlantUML"
    currentButton.Visible = True
    
    On Error GoTo ButtonAdd
    Set currentButton = toolBar.Controls(3)
    On Error GoTo 0
    currentButton.OnAction = "Module1.HidePlantuml"
    currentButton.Style = msoButtonCaption
    currentButton.Caption = "Hide PlantUML"
    currentButton.Visible = True
    
    On Error GoTo ButtonAdd
    Set currentButton = toolBar.Controls(4)
    On Error GoTo 0
    currentButton.OnAction = "Module1.Macro_UML_all"
    currentButton.Style = msoButtonCaption
    currentButton.Caption = "UML.*"
    currentButton.Visible = True
    
    On Error GoTo ButtonAdd
    Set currentButton = toolBar.Controls(5)
    On Error GoTo 0
    currentButton.OnAction = "Module1.Macro_UML_parg"
    currentButton.Style = msoButtonCaption
    currentButton.Caption = "UML.1"
    currentButton.Visible = True
    
    On Error GoTo ButtonAdd
    Set currentButton = toolBar.Controls(6)
    On Error GoTo 0
    currentButton.OnAction = ""
    currentButton.Style = msoButtonCaption
    currentButton.Caption = "Trace"
    currentButton.Visible = True
    Exit Function

ToolbarCreation:
    Set toolBar = ActiveDocument.CommandBars.Add(Name:="UML")
    Resume Next

ButtonAdd:
    Set currentButton = toolBar.Controls.Add(Type:=msoControlButton, Before:=toolBar.Controls.Count + 1)
    Resume Next

End Function

' =========================================================
' We need to double check that the style is present in the document
Function DoubleCheckStyle()
    CreateStyle
    CreateStyleImg
    Set mystyle = ActiveDocument.Styles("PlantUML")
    mystyle.BaseStyle = ActiveDocument.Styles.Item(1).BaseStyle
    
    mystyle.AutomaticallyUpdate = True
    With mystyle.Font
        .Name = "Courier New"
        .Size = 9
        .Hidden = False
        .Hidden = True
        .Color = wdColorGreen
    End With
End Function

' =========================================================
Function CreateStyle()
    On Error GoTo CreateStyleAdding
    Set mystyle = ActiveDocument.Styles("PlantUML")
    Exit Function
CreateStyleAdding:
    Set mystyle = ActiveDocument.Styles.Add(Name:="PlantUML", Type:=wdStyleTypeParagraph)
    mystyle.BaseStyle = ActiveDocument.Styles.Item(1).BaseStyle
    mystyle.AutomaticallyUpdate = True
    With mystyle.Font
        .Name = "Courier New"
        .Size = 9
        .Hidden = False
        .Hidden = True
        .Color = wdColorGreen
    End With
    With mystyle.ParagraphFormat
        With .Shading
            .Texture = wdTextureNone
            .ForegroundPatternColor = wdColorAutomatic
            .BackgroundPatternColor = wdColorLightGreen
        End With
        
        .LeftIndent = CentimetersToPoints(0)
        With .Shading
            .Texture = wdTextureNone
            .ForegroundPatternColor = wdColorAutomatic
            .BackgroundPatternColor = 12254650
        End With
        With .Borders(wdBorderLeft)
            .LineStyle = wdLineStyleDashLargeGap
            .LineWidth = wdLineWidth050pt
            .Color = 3910491
        End With
        With .Borders(wdBorderRight)
            .LineStyle = wdLineStyleDashLargeGap
            .LineWidth = wdLineWidth050pt
            .Color = 3910491
        End With
        With .Borders(wdBorderTop)
            .LineStyle = wdLineStyleDashLargeGap
            .LineWidth = wdLineWidth050pt
            .Color = 3910491
        End With
        With .Borders(wdBorderBottom)
            .LineStyle = wdLineStyleDashLargeGap
            .LineWidth = wdLineWidth050pt
            .Color = 3910491
        End With
        With .Borders
            .DistanceFromTop = 1
            .DistanceFromLeft = 4
            .DistanceFromBottom = 1
            .DistanceFromRight = 4
            .Shadow = False
        End With
    End With
    
    ' ajout des tabulations
    mystyle.NoSpaceBetweenParagraphsOfSameStyle = False
    mystyle.ParagraphFormat.TabStops.ClearAll
    mystyle.ParagraphFormat.TabStops.Add Position:= _
        CentimetersToPoints(1), Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
    mystyle.ParagraphFormat.TabStops.Add Position:= _
        CentimetersToPoints(2), Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
    mystyle.ParagraphFormat.TabStops.Add Position:= _
        CentimetersToPoints(3), Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
    mystyle.ParagraphFormat.TabStops.Add Position:= _
        CentimetersToPoints(4), Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces


End Function

' =========================================================
Function CreateStyleImg()
    On Error GoTo CreateStyleImgAdding
    Set mystyle = ActiveDocument.Styles("PlantUMLImg")
    mystyle.BaseStyle = ActiveDocument.Styles.Item(1).BaseStyle
    On Error GoTo 0
    Exit Function
CreateStyleImgAdding:
    Set mystyle = ActiveDocument.Styles.Add(Name:="PlantUMLImg", Type:=wdStyleTypeParagraph)
    mystyle.AutomaticallyUpdate = True
End Function

' =========================================================
' We show the hidden text
Function ShowPlantuml()
    DoubleCheckStyle

    'WordBasic.ShowComments
    ' We put a bookmark to retrieve position after showing the text
    ActiveDocument.Bookmarks.Add Name:="Position", Range:=Selection.Range
        
    Set mystyle = ActiveDocument.Styles("PlantUML")
    Set toolBar = ActiveDocument.CommandBars("UML")
        
    toolBar.Controls(2).Visible = False
    toolBar.Controls(3).Visible = True
    toolBar.Controls(4).Visible = True
    toolBar.Controls(5).Visible = True
        
    Call ShowHiddenText
        
    'We go back to the bookmark and we delete it
    Selection.GoTo What:=wdGoToBookmark, Name:="Position"
    ActiveDocument.Bookmarks(Index:="Position").Delete
    
End Function


' =========================================================
' MSR - gestion de l'option d'affichage des textes masques du style : "PlantUML"
Function HidePlantuml()
    DoubleCheckStyle
    'WordBasic.ShowComments
    ' We put a bookmark to retrieve position after showing the text
    ActiveDocument.Bookmarks.Add Name:="Position", Range:=Selection.Range
    
    Set mystyle = ActiveDocument.Styles("PlantUML")
    Set toolBar = ActiveDocument.CommandBars("UML")
        
    toolBar.Controls(2).Visible = True
    toolBar.Controls(3).Visible = False
    toolBar.Controls(4).Visible = False
    toolBar.Controls(5).Visible = False
    
    Call HideHiddenText
    
    'We go back to the bookmark and we delete it
    Selection.GoTo What:=wdGoToBookmark, Name:="Position"
    ActiveDocument.Bookmarks(Index:="Position").Delete

End Function

' =========================================================
Function HideHiddenText()
    ActiveDocument.ActiveWindow.View.ShowAll = False
    ActiveDocument.ActiveWindow.View.ShowHiddenText = False
End Function

' =========================================================
Function ShowHiddenText()
    ActiveDocument.ActiveWindow.View.ShowAll = False
    ActiveDocument.ActiveWindow.View.ShowHiddenText = True
End Function

' =========================================================
Function SwitchP()
    flag = Not (ActiveDocument.ActiveWindow.View.ShowTabs)
    ActiveDocument.ActiveWindow.View.ShowParagraphs = flag
    ActiveDocument.ActiveWindow.View.ShowTabs = flag
    ActiveDocument.ActiveWindow.View.ShowSpaces = flag
    ActiveDocument.ActiveWindow.View.ShowHyphens = flag
    ActiveDocument.ActiveWindow.View.ShowAll = False
End Function






