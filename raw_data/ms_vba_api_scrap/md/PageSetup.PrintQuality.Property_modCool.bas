Attribute VB_Name = "modCool"
Option Explicit
Public installPath As String
Public VBInst As VBIDE.VBE
Public VbInstCB As Office.CommandBar
Public VbInstCB1 As Office.CommandBar
Public Const App_Name = "Kimmo - Cool Source"
Public totPanes As Long
Public cntPanes As Long
Public aAll As Integer
Public VbCp As VBIDE.VBComponent
Public StrFunc() As String
Public ArrProcCode() As String
Public Const Quote As String = """"
Public Archive As Collection
Public ArchiveFilename As String
Public Dbase As String
Public WhatToProcess As Integer
Public Const SWP_NOMOVE As Long = 2
Public Const SWP_NOSIZE As Long = 1
Public allCode As String
Public IsOnTop As Long
Public procStart As Long
Public procEnd As Long
Public clearError As Integer
Public Const flags As Long = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_TOPMOST As Long = -1
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function FindWindowByTitle Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetModuleHandle Lib "kernel32.dll" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Public Declare Function LoadLibraryEx Lib "kernel32.dll" Alias "LoadLibraryExA" (ByVal lpFileName As String, ByVal hFile As Long, ByVal dwFlags As Long) As Long
Public Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Public Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Boolean
Public Declare Function AddFile Lib "zipit.dll" (ByVal ZipFileName As String, ByVal Filename As String, ByVal StoreDirInfo As Boolean, ByVal DOS83 As Boolean, ByVal Action As Integer, ByVal CompressionLevel As Integer) As Boolean
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const LVM_FIRST As Long = &H1000
Public Const LVM_SETCOLUMNWIDTH As Long = (LVM_FIRST + 30)
Public Const LVSCW_AUTOSIZE_USEHEADER As Long = -2
' start: browse for folder
Public Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type
Public Declare Function lstrcat Lib "KERNEL32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Public Const BIF_RETURNONLYFSDIRS = 1
Public Declare Function SHBrowseForFolder Lib "Shell32" (lpbi As BrowseInfo) As Long
Public Const MAX_PATH = 260
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Public Resp As Long
Public VM As String
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'Private Const MAX_PATH As Long = 260
Private Const ILD_TRANSPARENT As Long = &H1                     '  Display transparent
Private Const SHGFI_DISPLAYNAME As Long = &H200                 '  get display name
Private Const SHGFI_EXETYPE As Long = &H2000                    '  return exe type
'Private Const SHGFI_LARGEICON = &H0                      '  get large icon
Private Const SHGFI_SHELLICONSIZE As Long = &H4                 '  get shell size icon
Private Const SHGFI_SMALLICON As Long = &H1                     '  get small icon
Private Const SHGFI_SYSICONINDEX As Long = &H4000                '  get system icon index
Private Const SHGFI_TYPENAME As Long = &H400                    '  get type name
Private Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * MAX_PATH
    szTypeName As String * 80
End Type
Private IFileInfo As SHFILEINFO
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwAttributes As Long, psfi As SHFILEINFO, ByVal cbSizeFileInfo As Long, ByVal uFlags As Long) As Long
Private Const IFlags As Long = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE 'Too stuffs, just put it in decs
Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal hIml As Long, ByVal i As Long, ByVal hDCDest As Long, ByVal X As Long, ByVal Y As Long, ByVal flags As Long) As Long

Public Function ReadReg(sApp As String, sSection As String, Optional sKey As String = "Account", Optional sDefault As String = vbNullString) As String
    On Error Resume Next
    ReadReg = GetSetting(sApp, sSection, sKey, sDefault)
    Err.Clear
End Function
Public Sub SaveReg(sApp As String, sSection As String, Optional sValue As String = vbNullString, Optional sKey As String = "Account")
    On Error Resume Next
    SaveSetting sApp, sSection, sKey, sValue
    Err.Clear
End Sub
Public Function IsModuleAppropriate(VbCp As VBIDE.VBComponent) As Boolean
    On Error Resume Next
    If VbCp.Type <> vbext_ct_ResFile And VbCp.Type <> vbext_ct_RelatedDocument Then
        IsModuleAppropriate = True
    Else
        IsModuleAppropriate = False
    End If
    Err.Clear
End Function
Public Function ModuleSelectedLines(VbACp As VBIDE.CodePane) As String
    On Error Resume Next
    Dim iCurrentLine As Long
    Dim b As Long
    Dim c As Long
    Dim d As Long
    Set VbACp = VBInst.ActiveCodePane
    VbACp.GetSelection iCurrentLine, b, c, d
    ModuleSelectedLines = VbACp.CodeModule.lines(iCurrentLine, c)
    Err.Clear
End Function
Public Sub LstViewSwapSort(lstView As Control, lstHeader As Variant)
    On Error Resume Next
    Select Case lstView.SortOrder
    Case 0
        lstView.SortOrder = 1
    Case Else
        lstView.SortOrder = 0
    End Select
    lstView.SortKey = lstHeader.Index - 1
    lstView.Sorted = True
    lstView.Refresh
    Err.Clear
End Sub
Public Function LstViewGetRow(lstView As Object, ByVal idx As Long) As Variant
    On Error Resume Next
    Dim retarray() As String
    Dim clsColTot As Long
    Dim clsColCnt As Long
    clsColTot = lstView.ColumnHeaders.Count
    ReDim retarray(clsColTot)
    retarray(1) = lstView.ListItems(idx).Text
    clsColTot = clsColTot - 1
    For clsColCnt = 1 To clsColTot
        retarray(clsColCnt + 1) = lstView.ListItems(idx).SubItems(clsColCnt)
        Err.Clear
    Next
    LstViewGetRow = retarray
    Err.Clear
End Function
Public Function StringParse(retarray() As String, ByVal strText As String, Optional ByVal Delim As String = vbNullString) As Long
    On Error Resume Next
    Dim varArray() As String
    Dim varCnt As Long
    Dim VarS As Long
    Dim VarE As Long
    Dim varA As Long
    If Len(Delim) = 0 Then
        Delim = Chr$(253)
    End If
    varArray = Split(strText, Delim)
    VarS = LBound(varArray)
    VarE = UBound(varArray)
    varA = VarE + 1
    ReDim retarray(varA)
    For varCnt = VarS To VarE
        varA = varCnt + 1
        retarray(varA) = varArray(varCnt)
        Err.Clear
    Next
    Erase varArray
    StringParse = UBound(retarray)
    Err.Clear
End Function
Public Function MvFromMv(ByVal strOriginalMv As String, ByVal StartPos As Long, Optional ByVal NumOfItems As Long = -1, Optional ByVal Delim As String = vbNullString) As String
    On Error Resume Next
    Dim sporiginal() As String
    Dim spTot As Long
    Dim spCnt As Long
    Dim sLine As String
    Dim endPos As Long
    sLine = vbNullString
    If Len(Delim) = 0 Then
        Delim = Chr$(253)
    End If
    Call StringParse(sporiginal, strOriginalMv, Delim)
    spTot = UBound(sporiginal)
    If NumOfItems = -1 Then
        endPos = spTot
    Else
        endPos = (StartPos + NumOfItems) - 1
    End If
    For spCnt = StartPos To endPos
        If spCnt = endPos Then
            sLine = sLine & sporiginal(spCnt)
        Else
            sLine = sLine & sporiginal(spCnt) & Delim
        End If
        Err.Clear
    Next
    MvFromMv = sLine
    Err.Clear
End Function
Public Function boolIsBlank(ObjectName As Control, ByVal fldName As String) As Boolean
    On Error Resume Next
    Dim StrM As String
    Dim StrT As String
    Dim strO As String
    boolIsBlank = False
    If TypeName(ObjectName) = "TextBox" Then
        If Len(ObjectName.Text) = 0 Then
            strO = "type"
            GoSub CompileError
            Call intError(StrT, StrM)
            boolIsBlank = True
            ObjectName.SetFocus
        End If
    ElseIf TypeName(ObjectName) = "ComboBox" Then
        If Len(ObjectName.Text) = 0 Then
            strO = "select"
            GoSub CompileError
            Call intError(StrT, StrM)
            boolIsBlank = True
            ObjectName.SetFocus
        End If
    ElseIf TypeName(ObjectName) = "ImageCombo" Then
        If Len(ObjectName.Text) = 0 Then
            strO = "select"
            GoSub CompileError
            Call intError(StrT, StrM)
            boolIsBlank = True
            ObjectName.SetFocus
        End If
    ElseIf TypeName(ObjectName) = "CheckBox" Then
        If ObjectName.Value = 0 Then
            strO = "select"
            GoSub CompileError
            Call intError(StrT, StrM)
            boolIsBlank = True
            ObjectName.SetFocus
        End If
    ElseIf TypeName(ObjectName) = "ListBox" Then
        If (ObjectName.ListCount - 1) = -1 Then
            strO = "select"
            GoSub CompileError
            Call intError(StrT, StrM)
            boolIsBlank = True
            ObjectName.SetFocus
        End If
    ElseIf TypeName(ObjectName) = "OptionButton" Then
        If ObjectName.Value = False Then
            strO = "select"
            GoSub CompileError
            Call intError(StrT, StrM)
            boolIsBlank = True
            ObjectName.SetFocus
        End If
    ElseIf TypeName(ObjectName) = "Label" Then
        If Len(ObjectName.Caption) = 0 Then
            strO = "specify"
            GoSub CompileError
            Call intError(StrT, StrM)
            boolIsBlank = True
            ObjectName.SetFocus
        End If
    ElseIf TypeName(ObjectName) = "MaskEdBox" Then
        If ObjectName.Text = "__/__/____" Then
            strO = "enter"
            GoSub CompileError
            Call intError(StrT, StrM)
            boolIsBlank = True
            ObjectName.SetFocus
        End If
    End If
    Err.Clear
    Exit Function
CompileError:
    StrM = "The " & LCase$(fldName) & " cannot be left blank. Please " & strO & " the " & LCase$(fldName) & "."
    StrT = StringProperCase(fldName & " error")
    Err.Clear
    Return
    Err.Clear
End Function
Public Function intError(ByVal StrTitle As String, ByVal strmessage As String) As Integer
    On Error Resume Next
    intError = MsgBox(strmessage, vbOKOnly + vbExclamation + vbApplicationModal, StringProperCase(StrTitle))
    Err.Clear
End Function
Public Function StringProperCase(ByVal StrValue As String) As String
    On Error Resume Next
    StringProperCase = StrConv(StrValue, vbProperCase)
    Err.Clear
End Function
Public Function MvComment(ByVal strData As String, Optional ByVal Delim As String = vbNullString, Optional StartPos As Integer = 1) As String
    On Error Resume Next
    Dim sData() As String
    Dim tCnt As Integer
    Dim wCnt As Integer
    If Len(Delim) = 0 Then
        Delim = Chr$(253)
    End If
    Call StringParse(sData, strData, Delim)
    wCnt = UBound(sData)
    For tCnt = 1 To wCnt
        If tCnt >= StartPos Then sData(tCnt) = "'" & sData(tCnt)
        Err.Clear
    Next
    MvComment = MvFromArray(sData, Delim)
    Erase sData
    Err.Clear
End Function
Public Function MvFromArray(Varray As Variant, Optional ByVal Delim As String = vbNullString) As String
    On Error Resume Next
    If Len(Delim) = 0 Then
        Delim = Chr$(253)
    End If
    Dim i As Long
    Dim BldStr As String
    Dim strL As String
    Dim totArray As Long
    totArray = UBound(Varray)
    For i = 1 To totArray
        strL = Varray(i)
        If i = totArray Then
            BldStr = BldStr & strL
        Else
            BldStr = BldStr & strL & Delim
        End If
        Err.Clear
    Next
    MvFromArray = BldStr
    BldStr = vbNullString
    Err.Clear
End Function
Public Function TreeViewAddPath(TreeV As TreeView, ByVal sPath As String, Optional bProper As Boolean = False, Optional ByVal Image As String = vbNullString, Optional ByVal SelectedImage As String = vbNullString, Optional ByVal Tag As String = vbNullString, Optional Delim As String = "\") As Long
    On Error Resume Next
    Dim prevP As String
    Dim currP As String
    Dim lngC As Long
    Dim lngT As Long
    Dim pStr() As String
    Dim currN As String
    Dim nodeN As Node
    Dim pKey As String
    Dim cLoc As Long
    If bProper = True Then sPath = StringProperCase(sPath)
    Call StringParse(pStr, sPath, Delim)
    lngT = UBound(pStr)
    For lngC = 1 To lngT
        prevP = MvFromMv(sPath, 1, lngC - 1, Delim)
        currP = MvFromMv(sPath, 1, lngC, Delim)
        currN = pStr(lngC)
        If prevP = vbNullString Then
            ' this is the root node, locate it in the treeview
            cLoc = TreeViewPathLocation(TreeV, currP)
            ' if the root node does not exist, add it
            If cLoc = 0 Then
                Set nodeN = TreeV.Nodes.Add(, , currP, currN)
                If Len(Image) > 0 Then nodeN.Image = Image
                If Len(SelectedImage) > 0 Then nodeN.SelectedImage = SelectedImage
                nodeN.Tag = Tag
            Else
                Set nodeN = TreeV.Nodes(cLoc)
            End If
        Else
            ' this is the second, third etc node, locate it from the tree view
            cLoc = TreeViewPathLocation(TreeV, currP)
            If cLoc = 0 Then
                Set nodeN = TreeV.Nodes.Add(pKey, tvwChild, currP, currN)
                If Len(Image) > 0 Then nodeN.Image = Image
                If Len(SelectedImage) > 0 Then nodeN.SelectedImage = SelectedImage
                nodeN.Tag = Tag
            Else
                Set nodeN = TreeV.Nodes(cLoc)
            End If
        End If
        pKey = nodeN.Key
        If lngC = lngT Then
            If Len(Image) > 0 Then nodeN.Image = Image
            If Len(SelectedImage) > 0 Then nodeN.SelectedImage = SelectedImage
            nodeN.Tag = Tag
            TreeViewAddPath = nodeN.Index
            Exit For
        End If
        Err.Clear
    Next
    Err.Clear
End Function
Public Function MvCount(ByVal StringMv As String, Optional ByVal Delim As String = vbNullString) As Long
    On Error Resume Next
    Dim xNew() As String
    If Len(Delim) = 0 Then
        Delim = Chr$(253)
    End If
    Call StringParse(xNew, StringMv, Delim)
    MvCount = UBound(xNew)
    Erase xNew
    Err.Clear
End Function
Public Function TreeViewPathLocation(treeDms As TreeView, ByVal SearchPath As String) As Long
    On Error Resume Next
    Dim nTot As Long
    Dim nCnt As Long
    Dim nStr As String
    TreeViewPathLocation = 0
    SearchPath = LCase$(SearchPath)
    nTot = treeDms.Nodes.Count
    For nCnt = 1 To nTot
        nStr = LCase$(treeDms.Nodes(nCnt).FullPath)
        If nStr = SearchPath Then
            TreeViewPathLocation = nCnt
            Exit For
        End If
        Err.Clear
    Next
    Err.Clear
End Function
Public Function AlphaSortString(input_string As String) As String
    On Error GoTo Err_Handler:
    '-------------------------------------------
    'sort a string of text (paragraph or sentence)
    'alphabetically
    '-------------------------------------------
    Dim lng_counter As Long
    Dim lng_Upper As Long
    Dim lng_compare As Long
    Dim string_SortedCode As String
    Dim parts_string() As String
    Dim str_compare As String
    'split up the text in top textbox into seperate words
    'remove the carriage returns
    parts_string = Split(Replace$(LCase$(input_string), vbCrLf, vbNullString))
    'cache the array ubound
    lng_Upper = UBound(parts_string)
    Do
        'loop through every word starting from last to first
        For lng_counter = lng_Upper To 0 Step -1
            If str_compare = vbNullString Then
                'assign the current word to str_compare
                str_compare = Trim$(parts_string(lng_counter))
                'cache the arrays index
                lng_compare = lng_counter
                'this means were done sorting
                If lng_compare <= 0 Then
                    Exit Do
                End If
                'we dont want to bother comparing to a blank word
            ElseIf parts_string(lng_counter) <> vbNullString Then
                'if the compare word > this word
                If str_compare > parts_string(lng_counter) Then
                    'replace str_compare with this word
                    str_compare = parts_string(lng_counter)
                    'keep track of the current array index
                    lng_compare = lng_counter
                End If
            End If
            Err.Clear
        Next
        'add the next alphabetically sorted proccode
        'from arrProc to string_SortedCode
        string_SortedCode = (string_SortedCode & ArrProcCode(1, lng_compare) & vbCrLf)
        'since we have already added this word to the sorted
        'list, "remove" it from the original
        parts_string(lng_compare) = vbNullString
        'blank str_compare
        str_compare = vbNullString
    Loop
    'the return is all the modules code
    'sorted alphabetically
    AlphaSortString = string_SortedCode
    'free ram
    Erase parts_string
    Err.Clear
    Exit Function
Err_Handler:
    Err.Source = Err.Source & "." & varType(VbCp) & ".AlphaSortString"
    MsgBox Err.Number & vbTab & Err.Source & Err.Description, , App.Title
    Resume Next
    Err.Clear
End Function
Public Function ArraySearch(varArray() As String, ByVal StrSearch As String) As Long
    On Error Resume Next
    ArraySearch = 0
    Dim ArrayTot As Long
    Dim arrayCnt As Long
    Dim strCur As String
    StrSearch = LCase$(Trim$(StrSearch))
    ArrayTot = UBound(varArray)
    If ArrayTot = 0 Then
        Err.Clear
        Exit Function
    End If
    For arrayCnt = 1 To ArrayTot
        strCur = varArray(arrayCnt)
        strCur = LCase$(Trim$(strCur))
        Select Case strCur
        Case StrSearch
            ArraySearch = arrayCnt
            Exit For
        End Select
        Err.Clear
    Next
    Err.Clear
End Function
Public Function arrToSentence() As String
    On Error GoTo Err_Handler:
    Dim temp_string As String
    Dim lng_cnt As Long
    Dim lng_cnt_Tot As Long
    '-------------------------------------------
    'first loop through array and place procnames in
    'sentence format with each procname seperated by " "
    '-------------------------------------------
    lng_cnt_Tot = UBound(ArrProcCode, 2)
    For lng_cnt = 0 To lng_cnt_Tot
        temp_string = (temp_string & ArrProcCode(0, lng_cnt) & " ")
        Err.Clear
    Next
    arrToSentence = Trim$(temp_string)
    Err.Clear
    Exit Function
Err_Handler:
    Err.Source = Err.Source & "." & varType(VbCp) & ".arrToSentence"
    MsgBox Err.Number & vbTab & Err.Source & Err.Description, , App.Title
    Resume Next
    Err.Clear
End Function
Public Function boolFileExists(ByVal Filename As String) As Boolean
    On Error Resume Next
    boolFileExists = False
    If Len(Filename) = 0 Then
        Err.Clear
        Exit Function
    End If
    boolFileExists = IIf(Dir$(Filename) <> vbNullString, True, False)
    Err.Clear
End Function
Public Function AddButton(CB As Office.CommandBar, oButton As Office.CommandBarButton, Caption As String, Optional ToolTip As String = vbNullString, Optional Enabled As Boolean = True, Optional DescriptionText As String = vbNullString, Optional Tag As String = vbNullString, Optional sFaceID As String = vbNullString) As CommandBarEvents
    On Error Resume Next
    Set oButton = CB.Controls.Add(msoControlButton)
    With oButton
        .ToolTipText = ToolTip
        .Style = msoButtonCaption
        .State = msoButtonUp
        .Caption = Caption
        .DescriptionText = DescriptionText
        .Tag = Tag
        .Enabled = Enabled
        If Len(sFaceID) > 0 Then
            .FaceId = Val(sFaceID)
            .Style = msoButtonIcon
            If Len(Caption) > 0 Then
                .Style = msoButtonIconAndCaption
            End If
        End If
    End With
    Set AddButton = VBInst.Events.CommandBarEvents(oButton)
    Err.Clear
End Function
Public Function AddCommandBarControl(CB As Office.CommandBar, oButton As Office.CommandBarControl, Caption As String, Optional ToolTip As String = vbNullString, Optional Enabled As Boolean = True, Optional DescriptionText As String = vbNullString, Optional Tag As String = vbNullString, Optional CommandBarType As Office.MsoControlType = msoControlDropdown, Optional IsGroup As Boolean = False) As CommandBarEvents
    On Error Resume Next
    Set oButton = CB.Controls.Add(CommandBarType)
    With oButton
        .ToolTipText = ToolTip
        .Caption = Caption
        .DescriptionText = DescriptionText
        .Tag = Tag
        .Enabled = Enabled
        .BeginGroup = IsGroup
    End With
    Set AddCommandBarControl = VBInst.Events.CommandBarEvents(oButton)
    Err.Clear
End Function
Public Sub BackupFirst(Optional AskFirst As Boolean = True)
    On Error Resume Next
    Dim bPath As String
    Dim myFiles As String
    Dim memTot As Long
    Dim memCnt As Long
    'Dim memStr As String
    myFiles = vbNullString
    If AskFirst = True Then
        cntPanes = MsgBox("Do you want to backup the project to a compressed file first?", vbYesNo + vbQuestion + vbApplicationModal, "Confirm Backup")
    Else
        cntPanes = vbYes
    End If
    If cntPanes = vbYes Then
        Screen.MousePointer = vbHourglass
        bPath = StringGetFileToken(VBInst.ActiveVBProject.Filename, "p")
        bPath = StringGetFileToken(bPath, "p")
        ArchiveFilename = VBInst.ActiveVBProject.Filename
        ArchiveFilename = ExactPath(bPath) & "\" & StringGetFileToken(ArchiveFilename, "fo") & " " & Format$(Now(), "ddmmyyyy hhmmss") & ".zip"
        'file names for all components
        memTot = VBInst.ActiveVBProject.VBComponents.Count
        For memCnt = 1 To memTot
            myFiles = myFiles & VBInst.ActiveVBProject.VBComponents(memCnt).FileNames(1) & "|"
            Err.Clear
        Next
        myFiles = StringRemoveDelim(myFiles, "|")
        ZipFile frmProcess.ZIPLight1, ArchiveFilename, 6, myFiles
        'Backup_ToCompressedFile ArchiveFilename, ExactPath(StringGetFileToken(VBInst.ActiveVBProject.Filename, "p"))
EndProc:
        Screen.MousePointer = vbDefault
    End If
    Err.Clear
End Sub
Private Function ZipFile(MyZipLight As ZIPLight, ByVal ZipFileName As String, CompressionLevel As Integer, ParamArray FilesToAdd()) As String
    On Error Resume Next
    Dim varFile As Variant
    Dim strFile As String
    strFile = vbNullString
    For Each varFile In FilesToAdd
        strFile = strFile & CStr(varFile) & "|"
        Err.Clear
    Next
    strFile = StringRemoveDelim(strFile, "|")
    If MvCount(strFile, "|") = 1 Then
        ZipFileName = File_Token(strFile, "p") & "\" & File_Token(strFile, "fo") & ".zip"
    End If
    If File_Exists(ZipFileName) = True Then Kill ZipFileName
    ' create an empty zip file
    With MyZipLight
        .ZipFileName = ZipFileName
        '.SourceDirectory = File_Token(strFile, "p")
        .FilesToProcess = strFile
        .AllowErrorReporting = False
        .CompressionLevel = CompressionLevel
        .Overwrite = True
        .Add
    End With
    ZipFile = ZipFileName
    Err.Clear
End Function
Public Function File_Exists(ByVal strFile As String) As Boolean
    On Error Resume Next
    Dim fs As FileSystemObject
    Set fs = New FileSystemObject
    File_Exists = fs.FileExists(strFile)
    Set fs = Nothing
    Err.Clear
End Function
Public Function File_Token(ByVal strFileName As String, Optional ByVal Sretrieve As String = "F", Optional ByVal Delim As String = "\") As String
    On Error Resume Next
    Dim intNum As Long
    Dim sNew As String
    File_Token = strFileName
    Select Case UCase$(Sretrieve)
    Case "D"
        File_Token = Left$(strFileName, 3)
    Case "F"
        intNum = InStrRev(strFileName, Delim)
        If intNum <> 0 Then
            File_Token = Mid$(strFileName, intNum + 1)
        End If
    Case "P"
        If InStr(1, strFileName, Delim, vbTextCompare) > 0 Then
            intNum = InStrRev(strFileName, Delim)
            If intNum <> 0 Then
                File_Token = Mid$(strFileName, 1, intNum - 1)
            End If
        Else
            File_Token = vbNullString
        End If
    Case "E"
        intNum = InStrRev(strFileName, ".")
        If intNum <> 0 Then
            File_Token = Mid$(strFileName, intNum + 1)
        End If
    Case "FO"
        sNew = strFileName
        intNum = InStrRev(sNew, Delim)
        If intNum <> 0 Then
            sNew = Mid$(sNew, intNum + 1)
        End If
        intNum = InStrRev(sNew, ".")
        If intNum <> 0 Then
            sNew = Left$(sNew, intNum - 1)
        End If
        File_Token = sNew
    Case "PF"
        intNum = InStrRev(strFileName, ".")
        If intNum <> 0 Then
            File_Token = Left$(strFileName, intNum - 1)
        End If
    End Select
    Err.Clear
End Function
Public Sub Backup_ToCompressedFile(ByVal ArchiveName As String, ParamArray FoldersToBackup())
    On Error Resume Next
    Dim eachFolder As Variant
    Dim xFolder As String
    Dim vFiles As New Collection
    Dim tAdded As Long
    Screen.MousePointer = vbHourglass
    For Each eachFolder In FoldersToBackup
        xFolder = CStr(eachFolder)
        Call TotalDirFiles(xFolder, vFiles)
        Err.Clear
    Next
    tAdded = Zip_Add2Archive(ArchiveName, vFiles, 1, True, False, 9)
    If tAdded <> 0 Then
        Call MsgBox("Not all files could be added to the archive!", vbOKOnly + vbExclamation, tAdded & " Files Not Added")
    End If
    Screen.MousePointer = vbDefault
    Err.Clear
End Sub
Public Sub CloseProgress()
    On Error Resume Next
    frmPg.Hide
    Err.Clear
End Sub
'Public Function CloseWindow(tpWindow As VBIDE.vbext_WindowType) As Boolean
'    On Error GoTo HandleError
'    With VBInst
'        nLoop = 1
'        While nLoop < .Windows.Count
'            If .Windows.Item(nLoop).Type = tpWindow Then
'                .Windows.Item(nLoop).Close
'            Else
'                nLoop = nLoop + 1
'            End If
'        Wend
'    End With
'    CloseWindow = True
'    Exit Function
'HandleError:
'    CloseWindow = False
'End Function
Public Function CloseWindows(tpWindow() As VBIDE.vbext_WindowType) As Boolean
    On Error GoTo HandleError
    Dim nLoop As Integer
    Dim nTypes As Integer
    Dim nTypes_Tot As Integer
    With VBInst
        nLoop = 1
        While nLoop < .Windows.Count
            nTypes_Tot = UBound(tpWindow())
            For nTypes = 0 To nTypes_Tot
                If .Windows.Item(nLoop).Type = tpWindow(nTypes) Then
                    .Windows.Item(nLoop).Close
                    nLoop = nLoop - 1
                    Exit For
                End If
                Err.Clear
            Next
            nLoop = nLoop + 1
        Wend
    End With
    CloseWindows = True
    Err.Clear
    Exit Function
HandleError:
    CloseWindows = False
    Err.Clear
End Function
Public Sub ApplicationOnTop(ByVal wHandle As Long)
    On Error Resume Next
    IsOnTop = SetWindowPos(wHandle, HWND_TOPMOST, 0, 0, 0, 0, flags)
    Err.Clear
End Sub
Public Function UpdateProgress(ByVal minValue As Long, ByVal maxValue As Long, Optional ByVal Note As String = vbNullString) As Boolean
    On Error Resume Next
    frmPg.ProgressShow minValue, maxValue, Note
    If frmPg.cmdStop.Tag = "s" Then
        UpdateProgress = False
        frmPg.Hide
    Else
        UpdateProgress = True
    End If
    Err.Clear
End Function
Public Function StringGetFileToken(ByVal strFileName As String, Optional ByVal Sretrieve As String = vbNullString, Optional ByVal Delim As String = "\") As String
    On Error Resume Next
    Dim intNum As Long
    Dim sNew As String
    StringGetFileToken = strFileName
    If Len(Sretrieve) = 0 Then
        Sretrieve = "F"
    End If
    If Len(Delim) = 0 Then
        Delim = "\"
    End If
    Select Case UCase$(Sretrieve)
    Case "D"
        StringGetFileToken = Left$(strFileName, 3)
    Case "F"
        intNum = InStrRev(strFileName, Delim)
        If intNum <> 0 Then
            StringGetFileToken = Mid$(strFileName, intNum + 1)
        End If
    Case "P"
        intNum = InStrRev(strFileName, Delim)
        If intNum <> 0 Then
            StringGetFileToken = Mid$(strFileName, 1, intNum - 1)
        End If
    Case "E"
        intNum = InStrRev(strFileName, ".")
        If intNum <> 0 Then
            StringGetFileToken = Mid$(strFileName, intNum + 1)
        End If
    Case "FO"
        sNew = strFileName
        intNum = InStrRev(sNew, Delim)
        If intNum <> 0 Then
            sNew = Mid$(sNew, intNum + 1)
        End If
        intNum = InStrRev(sNew, ".")
        If intNum <> 0 Then
            sNew = Left$(sNew, intNum - 1)
        End If
        StringGetFileToken = sNew
    Case "PF"
        intNum = InStrRev(strFileName, ".")
        If intNum <> 0 Then
            StringGetFileToken = Left$(strFileName, intNum - 1)
        End If
    End Select
    sNew = vbNullString
    Err.Clear
End Function
Public Function ExactPath(ByVal StrValue As String) As String
    On Error Resume Next
    If Right$(StrValue, 1) = "\" Then
        StrValue = Left$(StrValue, Len(StrValue) - 1)
    End If
    ExactPath = StrValue
    Err.Clear
End Function
Public Function TotalDirFiles(ByVal DirPath As String, FilesCollection As Collection, Optional ByVal FilePattern As String = "*.*") As Long
    On Error Resume Next
    Dim sFile As String
    Dim StrF As String
    Dim StrP As String
    Dim lngL As Long
    StrP = DirPath & "\"
    StrF = StrP & FilePattern
    sFile = Dir$(StrF)
    lngL = Len(sFile)
    Do While lngL
        sFile = StrP & sFile
        FilesCollection.Add sFile, sFile
        sFile = Dir$
        lngL = Len(sFile)
    Loop
    TotalDirFiles = FilesCollection.Count
    sFile = vbNullString
    StrF = vbNullString
    StrP = vbNullString
    Err.Clear
End Function
Public Function Zip_Add2Archive(ZipFileName As String, Files As Collection, Action As Integer, StorePathInfo As Boolean, UseDOS83 As Boolean, CompressionLevel As Integer) As Long
    On Error Resume Next
    Dim i As Long
    Dim Result As Long
    Dim nAdded As Long
    Dim FilesToAdd As Collection
    Dim i_Tot As Long
    nAdded = 0
    If Not win_Function_Exist("zipit.dll", "AddFile") Then
        MsgBox "Info zip dll is missing, contact the author about this problem.", vbCritical, "Error": Exit Function
    End If
    If Not win_Function_Exist("zipit.dll", "ExtractFile") Then
        MsgBox "Info zip dll is missing, contact the author about this problem.", vbCritical, "Error": Exit Function
    End If
    If Not win_Function_Exist("zipit.dll", "DeleteFile") Then
        MsgBox "Info zip dll is missing, contact the author about this problem.", vbCritical, "Error": Exit Function
    End If
    If Archive.Count = 0 Then
        If Dir$(ZipFileName, vbHidden Or vbSystem Or vbReadOnly) <> vbNullString Then
            Kill ZipFileName
        End If
    End If
    Set FilesToAdd = FindFiles(Files)
    i_Tot = FilesToAdd.Count
    Call InitProgress(i_Tot)
    For i = 1 To i_Tot
        If UpdateProgress(i, i_Tot, "Backing " & FilesToAdd(i)) = False Then
            Exit For
        End If
        If AddFile(ZipFileName, FilesToAdd(i), StorePathInfo, UseDOS83, Action, CompressionLevel) Then
            Result = Result + 1
        Else
            nAdded = nAdded + 1
        End If
        Err.Clear
    Next
    CloseProgress
    Zip_Add2Archive = nAdded
    Err.Clear
End Function
Private Function win_Function_Exist(sModule As String, sFunction As String) As Boolean
    On Error Resume Next
    Dim hHandle As Long
    hHandle = GetModuleHandle(sModule)
    If hHandle = 0 Then
        hHandle = LoadLibraryEx(sModule, 0&, 0&)
        If GetProcAddress(hHandle, sFunction) = 0 Then
            win_Function_Exist = False
        Else
            win_Function_Exist = True
        End If
        FreeLibrary hHandle
    Else
        If GetProcAddress(hHandle, sFunction) <> 0 Then
            win_Function_Exist = True
        End If
    End If
    Err.Clear
End Function
Private Function FindFiles(Files As Collection)
    On Error Resume Next
    Dim Result As New Collection
    Dim Path As String
    Dim r As String
    Dim i As Long
    Dim i_Tot As Long
    i_Tot = Files.Count
    For i = 1 To i_Tot
        Path = File_ParsePath(Files(i))
        r = Dir$(Files(i), vbHidden Or vbSystem Or vbReadOnly)
        Do Until r = vbNullString
            Result.Add Path & r
            r = Dir$()
        Loop
        Err.Clear
    Next
    Set FindFiles = Result
    Err.Clear
End Function
Public Function InitProgress(ByVal totRecords As Long, Optional OnTop As Boolean = True) As Boolean
    On Error Resume Next
    Dim aHandle As Long
    InitProgress = True
    frmPg.chkPrg.Width = 0
    frmPg.lblTime.Caption = vbNullString
    frmPg.lblPerc.Caption = vbNullString
    frmPg.lblRecs.Caption = vbNullString
    frmPg.lblHeader.Caption = vbNullString
    frmPg.cmdStop.Tag = vbNullString
    frmPg.cmdStop.Enabled = True
    If totRecords <= 0 Then
        InitProgress = False
EndProc:
        Screen.MousePointer = vbDefault
        frmPg.Hide
        Err.Clear
        Exit Function
    End If
    frmPg.Show
    frmPg.Refresh
    If OnTop = True Then
        aHandle = FindWindowByTitle(vbNullString, "Processing")
        If aHandle > 0 Then
            ApplicationOnTop aHandle
        End If
    End If
    Err.Clear
End Function
Private Function File_ParsePath(Path As String) As String
    On Error Resume Next
    Dim a As Long
    Dim A_Cnt As Long
    A_Cnt = Len(Path)
    For a = A_Cnt To 1 Step -1
        If Mid$(Path, a, 1) = "\" Or Mid$(Path, a, 1) = "/" Then
            If Mid$(Path, a, 1) = "\" Then
                File_ParsePath = LCase$(Left$(Path, a - 1) & "\")
            Else
                File_ParsePath = LCase$(Left$(Path, a - 1) & "/")
            End If
            Err.Clear
            Exit Function
        End If
        Err.Clear
    Next
    Err.Clear
End Function
Public Function ComponentType(VbCp As VBIDE.VBComponent) As String
    On Error Resume Next
    Select Case VbCp.Type
    Case vbext_ct_ActiveXDesigner
        ComponentType = "ActiveX Designers"
    Case vbext_ct_ClassModule
        ComponentType = "Class Modules"
    Case vbext_ct_PropPage
        ComponentType = "Property Pages"
    Case vbext_ct_UserControl
        ComponentType = "User Controls"
    Case vbext_ct_DocObject
        ComponentType = "User Documents"
    Case vbext_ct_VBForm
        If VbCp.Properties(35) Then
            ComponentType = "Child Forms"
        Else
            ComponentType = "Forms"
        End If
    Case vbext_ct_VBMDIForm
        ComponentType = "MDI Form"
    Case vbext_ct_StdModule
        ComponentType = "Standard Modules"
    Case vbext_ct_ResFile
        ComponentType = "Resource File"
    Case vbext_ct_RelatedDocument
        ComponentType = "Related Document"
    Case vbext_ct_MSForm
        ComponentType = "MS Form"
    Case Else
        ComponentType = "Unknown Component Type"
    End Select
    Err.Clear
End Function
Public Function MemberType(VbMember As VBIDE.Member) As String
    On Error Resume Next
    'Dim lngLoc As Long
    'Dim strCode As String
    
    Select Case VbMember.Type
    Case vbext_mt_Method
        MemberType = "Method"
    Case vbext_mt_Property
        MemberType = "Property"
    Case vbext_mt_Event
        MemberType = "Event"
    Case vbext_mt_Variable
        MemberType = "Variable"
    Case vbext_mt_Const
        MemberType = "Constant"
    End Select
    Err.Clear
End Function
Public Function MemberScope(VbMember As VBIDE.Member) As String
    On Error Resume Next
    Select Case VbMember.Scope
    Case vbext_Friend
        MemberScope = "Friend"
    Case vbext_Private
        MemberScope = "Private"
    Case vbext_Public
        MemberScope = "Public"
    End Select
    Err.Clear
End Function
Public Sub LstViewAutoResize(lstView As Control)
    On Error Resume Next
    'Size each column based on the maximum of
    'EITHER the columnheader text width, or,
    'if the items below it are wider, the
    'widest list item in the column
    Dim col2adjust As Long
    Dim col2adjust_Tot As Long
    If lstView.ListItems.Count = 0 Then
        Err.Clear
        Exit Sub
    End If
    col2adjust_Tot = lstView.ColumnHeaders.Count - 1
    For col2adjust = 0 To col2adjust_Tot
        Call SendMessage(lstView.hwnd, LVM_SETCOLUMNWIDTH, col2adjust, ByVal LVSCW_AUTOSIZE_USEHEADER)
        Err.Clear
    Next
    LstViewResizeMax lstView
    lstView.Refresh
    Err.Clear
End Sub
Public Sub LstViewResizeMax(lstView As Object)
    On Error Resume Next
    'Because applying the LVSCW_AUTOSIZE_USEHEADER
    'message to the last column in the control always
    'sets its width to the maximum remaining control
    'space, calling SendMessage passing the last column
    'will cause the listview data to utilize the full
    'control width space. For example, if a four-column
    'listview had a total width of 2000, and the first
    'three columns each had individual widths of 250,
    'calling this will cause the last column to widen
    'to cover the remaining 1250.
    'For this message to (visually) work as expected,
    'all columns should be within the viewing rect of the
    'listview control; if the last column is wider than
    'the control the message works, but the columns
    'remain wider than the control.
    Dim col2adjust As Long
    col2adjust = lstView.ColumnHeaders.Count - 1
    Call SendMessage(lstView.hwnd, LVM_SETCOLUMNWIDTH, col2adjust, ByVal LVSCW_AUTOSIZE_USEHEADER)
    Err.Clear
End Sub
Public Function ProjectCode() As String
    On Error Resume Next
    Dim xCode As String
    xCode = vbNullString
    totPanes = VBInst.ActiveVBProject.VBComponents.Count
    For cntPanes = 1 To totPanes
        Set VbCp = VBInst.ActiveVBProject.VBComponents(cntPanes)
        If IsModuleAppropriate(VbCp) = False Then GoTo NextModule
        xCode = xCode & vbCrLf & VbCp.CodeModule.lines(1, VbCp.CodeModule.CountOfLines)
NextModule:
        Err.Clear
    Next
    ProjectCode = xCode
    Err.Clear
End Function
Public Function LstViewUpdate(Arrfields() As String, lstView As Control, Optional ByVal lstIndex As String = vbNullString) As Long
    On Error Resume Next
    Dim itmX As Object
    Dim FldCnt As Integer
    Dim sStr As String
    Dim wCnt As Integer
    sStr = CStr(Val(lstIndex))
    Select Case sStr
    Case "0"
        Set itmX = lstView.ListItems.Add()
    Case Else
        Set itmX = lstView.ListItems(Val(lstIndex))
    End Select
    wCnt = UBound(Arrfields) - 1
    With itmX
        .Text = Arrfields(1)
        For FldCnt = 1 To wCnt
            .SubItems(FldCnt) = Arrfields(FldCnt + 1)
            Err.Clear
        Next
    End With
    LstViewUpdate = itmX.Index
    Set itmX = Nothing
    Err.Clear
End Function
Public Function ScanVariablesUSE(ByVal sCode As String, sSearch As String) As Long
    On Error Resume Next
    Dim clsRegExp As RegExp
    Dim clsMatchCol As MatchCollection
    Set clsRegExp = New RegExp
    'first remove comments
    Dim rsCnt As Long
    Dim rsTot As Long
    Dim spLine() As String
    Dim rsStr As String
    rsTot = StrParse(spLine, sCode, vbNewLine)
    For rsCnt = 1 To rsTot
        rsStr = Trim$(spLine(rsCnt))
        If Left$(rsStr, 1) = "'" Then spLine(rsCnt) = vbNullString
        Err.Clear
    Next
    sCode = MvFromArray(spLine, vbNewLine)
    With clsRegExp
        .Global = True
        .IgnoreCase = False
        .Pattern = "'.*\n"  ' pattern starts with single quote, (match any single character), (match the preceeding character)
        sCode = .Replace(sCode, vbCrLf)
        'next remove new line char
        .Pattern = "\s*_{1}\s*\n"
        sCode = .Replace(sCode, vbNullString)
        'replace more then 1 space to 1
        .Pattern = "( ){2,}"
        sCode = .Replace(sCode, " ")
        'replace other duplicated ascii chars
        .Pattern = "\r{2,}"
        sCode = .Replace(sCode, vbCrLf)
        .Pattern = "\n{2,}"
        sCode = .Replace(sCode, vbCrLf)
    End With
    'i = 0
    'Search for variables global, private, public, static, dim, const
    '  '(\s|\(|\+|\-|\*|\&|\/|\\|^|\=){1}varname(\s|\)|\+|\-|\*|\&|\/|\\|^|\=){0,1}
    clsRegExp.Pattern = "(\s|\(|\+|\-|\*|\&|\/|\\|^|\=){1}" & sSearch & "(\s|\)|\+|\-|\*|\&|\/|\\|^|\=){0,1}"
    Set clsMatchCol = clsRegExp.Execute(sCode)
    ScanVariablesUSE = clsMatchCol.Count
    Set clsMatchCol = Nothing
    Err.Clear
    Exit Function
    Err.Clear
End Function
Public Function StrParse(retarray() As String, ByVal strText As String, Optional ByVal Delim As String = vbNullString) As Long
    On Error Resume Next
    Dim varArray() As String
    Dim varCnt As Long
    Dim VarS As Long
    Dim VarE As Long
    Dim varA As Long
    If Len(Delim) = 0 Then
        Delim = Chr$(253)
    End If
    varArray = Split(strText, Delim)
    VarS = LBound(varArray)
    VarE = UBound(varArray)
    ReDim retarray(VarE + 1)
    For varCnt = VarS To VarE
        varA = varCnt + 1
        retarray(varA) = varArray(varCnt)
        Err.Clear
    Next
    StrParse = UBound(retarray)
    Err.Clear
End Function
Public Function StringProcedureName(ByVal strDeclaration As String) As String
    On Error Resume Next
    Dim fBracket As Long
    Dim sProcedure As String
    fBracket = InStr(1, strDeclaration, "(")
    sProcedure = Left$(strDeclaration, fBracket - 1)
    StringProcedureName = MVLastItem(sProcedure, " ")
    Err.Clear
End Function
Public Function MVLastItem(ByVal StringMv As String, Optional ByVal Delim = vbNullString) As String
    On Error Resume Next
    Dim spValues() As String
    Dim spTotal As Long
    If Len(Delim) = 0 Then
        Delim = Chr$(253)
    End If
    spValues = Split(StringMv, Delim)
    spTotal = UBound(spValues)
    MVLastItem = spValues(spTotal)
    Erase spValues
    Err.Clear
End Function
Public Sub Code_InsertBeforeAfter(VbCp As VBIDE.VBComponent, ByVal sSearch As String, ByVal sInsert As String, Optional ByVal SPos As String = "b")
    On Error Resume Next
    Dim tLines As Long
    Dim tCount As Long
    Dim lBefore As String
    Dim lAfter As String
    Dim currLine As String
    sSearch = Trim$(sSearch)
    sInsert = Trim$(sInsert)
    tLines = VbCp.CodeModule.CountOfLines
    Call InitProgress(tLines)
    For tCount = tLines To 1 Step -1
        If UpdateProgress(tCount, tLines, VbCp.Name & ".Inserting Code Before / After") = False Then
            Exit For
        End If
        currLine = Trim$(VbCp.CodeModule.lines(tCount, 1))
        lBefore = Trim$(VbCp.CodeModule.lines(tCount - 1, 1))
        lAfter = Trim$(VbCp.CodeModule.lines(tCount + 1, 1))
        Select Case currLine
        Case sSearch
            Select Case SPos
            Case "b"
                If lBefore <> sInsert Then VbCp.CodeModule.InsertLines tCount, sInsert
            Case "a"
                If lAfter <> sInsert Then VbCp.CodeModule.InsertLines tCount + 1, sInsert
            End Select
        End Select
        Err.Clear
    Next
    CloseProgress
    Err.Clear
End Sub
Public Function Procedure_Variables(ByVal strCode As String) As Collection
    On Error Resume Next
    Dim xCol As New Collection
    Dim spLines() As String
    Dim spTot As Long
    Dim spCnt As Long
    spLines = Split(strCode, vbNewLine)
    spTot = UBound(spLines)
    For spCnt = 0 To spTot
        If IsVariable(spLines(spCnt)) = True Then
            xCol.Add VariableName(spLines(spCnt))
        End If
        Err.Clear
    Next
    Set Procedure_Variables = xCol
    Err.Clear
End Function
Public Function IsVariable(ByVal currLine As String) As Boolean
    On Error Resume Next
    currLine = Trim$(currLine)
    If Left$(currLine, Len("Dim ")) = "Dim " Then
        IsVariable = True
    ElseIf Left$(currLine, Len("Const ")) = "Const " Then
        IsVariable = True
    Else
        IsVariable = False
    End If
    Err.Clear
End Function
Public Function VariableName(ByVal currLine As String) As String
    On Error Resume Next
    Dim bPos As Long
    Dim bStr As String
    currLine = Trim$(currLine)
    bStr = Split(currLine, " ")(1)
    bPos = InStr(1, bStr, "(")
    If bPos > 0 Then
        bStr = Left$(bStr, bPos - 1)
    End If
    VariableName = Trim$(bStr)
    Err.Clear
End Function
Public Sub LstViewToWorkSheetAsIs(lstView As Object, ByVal strFileName As String, Optional ByVal strHeader As String = vbNullString, _
    Optional ByVal LeftFooter As String = vbNullString, Optional ByVal CenterFooter As String = vbNullString, Optional ByVal strTab As String = vbNullString)
    On Error Resume Next
    Dim XLWkb As Excel.Workbook
    Dim lngNumCols As Long
    Dim lngNumRows As Long
    Dim sprLine() As String
    Dim xPageSetUp As Excel.PageSetup
    Dim xLWksNew As Excel.Worksheet
    Dim lngC As Long
    Dim lngNext As Long
    Dim bFound As Boolean
    Dim xlRange As Excel.Range
    Dim xApp As Excel.Application
    Dim spCnt As Long
    Dim spTot As Long
    Dim strHeads As String
    If boolFileExists(strFileName) = True Then
        Kill strFileName
    End If
    Set xApp = New Excel.Application
    xApp.ActiveWindow.FreezePanes = True
    xApp.ActiveWindow.Visible = False
    xApp.ScreenUpdating = False
    xApp.Workbooks.Add
    Set XLWkb = xApp.ActiveWorkbook
    lngNext = 1
    Do Until bFound = True
        Set xLWksNew = xApp.Worksheets("sheet" & lngNext)
        If xLWksNew Is Nothing Then
            bFound = False
            lngNext = lngNext + 1
        Else
            bFound = True
        End If
    Loop
    xLWksNew.Name = StringProperCase(ExcelCorrectSheetName(strTab))
    xLWksNew.Activate
    xLWksNew.Cells.NumberFormat = "General"
    xLWksNew.PageSetup.PrintQuality = Array(300, 300)
    xLWksNew.Cells.Font.Name = lstView.Font.Name
    xLWksNew.Cells.Font.Size = lstView.Font.Size
    GoSub DoPageSetup
    ' process worksheet
    strHeads = LstViewColNames(lstView)
    ' add first row as headings
    Call StringParse(sprLine, strHeads, ",")
    lngNumCols = UBound(sprLine)
    For lngC = 1 To lngNumCols
        If IsNumeric(sprLine(lngC)) = True Then
            sprLine(lngC) = "'" & sprLine(lngC)
        End If
        Set xlRange = xLWksNew.Cells(1, lngC)
        xlRange.Value = sprLine(lngC)
        xlRange.Borders.Weight = Excel.XlBorderWeight.xlThin
        xlRange.Interior.ColorIndex = 15
        xlRange.Interior.Pattern = Excel.xlSolid
        Err.Clear
    Next
    ' format first row to be bold
    xLWksNew.Rows(1).Font.Bold = True
    ' add the remaining lines
    spTot = lstView.ListItems.Count
    lngNumRows = 1
    If InitProgress(spTot) = False Then
        Err.Clear
        Exit Sub
    End If
    For spCnt = 1 To spTot
        If UpdateProgress(spCnt, spTot, "Converting list view to a worksheet...") = False Then
            Exit For
        End If
        sprLine = LstViewGetRow(lstView, spCnt)
        lngNumCols = UBound(sprLine)
        lngNumRows = lngNumRows + 1
        For lngC = 1 To lngNumCols
            sprLine(lngC) = MvToText(sprLine(lngC))
            If IsNumeric(sprLine(lngC)) = True Then
                sprLine(lngC) = "'" & sprLine(lngC)
            End If
            With xLWksNew.Cells(lngNumRows, lngC)
                .Value = sprLine(lngC)
                .Borders.Weight = Excel.XlBorderWeight.xlThin
                .Font.Italic = lstView.Font.Italic
                .Font.Bold = LstViewRowColumnIsBold(lstView, lngNumRows - 1, lngC)
                .Font.color = LstViewRowColumnIsColor(lstView, lngNumRows - 1, lngC)
            End With
            Err.Clear
        Next
        Err.Clear
    Next
    ' ensure columns are auto fixed
    For lngC = 1 To lngNumCols
        Set xlRange = xLWksNew.Columns(lngC)
        xlRange.EntireColumn.AutoFit
        Err.Clear
    Next
    xApp.ScreenUpdating = True
    XLWkb.SaveAs strFileName
    xApp.Quit
    Set xPageSetUp = Nothing
    Set xLWksNew = Nothing
    Set xlRange = Nothing
    Set xApp = Nothing
    CloseProgress
    Err.Clear
    Exit Sub
DoPageSetup:
    Set xPageSetUp = xLWksNew.PageSetup
    With xPageSetUp
        .PrintTitleRows = vbNullString
        .PrintTitleColumns = vbNullString
        .PrintArea = vbNullString
        .CenterHeader = vbNullString
        .LeftHeader = "&B&I&" & Quote & "Tahoma" & Quote & "&12" & strHeader
        .RightHeader = vbNullString
        .LeftFooter = "&B&" & Quote & "Tahoma" & Quote & "&8" & LeftFooter
        .CenterFooter = "&B&" & Quote & "Tahoma" & Quote & "&8" & CenterFooter
        .RightFooter = "&" & Quote & "Tahoma" & Quote & "&8" & "Page &P of &N  &D"
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintNotes = False
        .CenterHorizontally = False
        .CenterVertically = False
        .PrintTitleRows = "$1:$1"
        .Orientation = Excel.xlLandscape
        .Draft = False
        .PaperSize = Excel.xlPaperA4
        .FirstPageNumber = Excel.xlAutomatic
        .Order = Excel.xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 100
    End With
    Err.Clear
    Return
    Err.Clear
End Sub
Public Function ExcelCorrectSheetName(ByVal StrValue As String) As String
    On Error Resume Next
    StrValue = Replace$(StrValue, ":", vbNullString)
    StrValue = Replace$(StrValue, ",", vbNullString)
    StrValue = Replace$(StrValue, "/", vbNullString)
    StrValue = Replace$(StrValue, "\", vbNullString)
    StrValue = Replace$(StrValue, "?", vbNullString)
    StrValue = Replace$(StrValue, "*", vbNullString)
    StrValue = Replace$(StrValue, "]", vbNullString)
    StrValue = Replace$(StrValue, "[", vbNullString)
    ExcelCorrectSheetName = Left$(Trim$(StrValue), 31)
    Err.Clear
End Function
Public Function LstViewColNames(lstView As Object) As String
    On Error Resume Next
    Dim strHead As String
    Dim StrName As String
    Dim clsColTot As Long
    Dim clsColCnt As Long
    strHead = vbNullString
    clsColTot = lstView.ColumnHeaders.Count
    For clsColCnt = 1 To clsColTot
        StrName = lstView.ColumnHeaders(clsColCnt).Text
        Select Case clsColCnt
        Case clsColTot
            strHead = strHead & StrName
        Case Else
            strHead = strHead & StrName & ","
        End Select
        Err.Clear
    Next
    LstViewColNames = strHead
    Err.Clear
End Function
Public Function MvToText(ByVal Mvtext As String, Optional ByVal Delim As String = vbNullString) As String
    On Error Resume Next
    If Len(Delim) = 0 Then
        Delim = Chr$(253)
    End If
    MvToText = Replace$(Mvtext, Delim, vbNewLine)
    Err.Clear
End Function
Public Function LstViewRowColumnIsBold(lstView As Object, rowPos As Long, ColumnPos As Long) As Boolean
    On Error Resume Next
    Dim numColumn As Long
    numColumn = ColumnPos - 1
    If numColumn = 0 Then
        LstViewRowColumnIsBold = lstView.ListItems(rowPos).Bold
    Else
        LstViewRowColumnIsBold = lstView.ListItems(rowPos).ListSubItems(numColumn).Bold
    End If
    Err.Clear
End Function
Public Function LstViewRowColumnIsColor(lstView As Object, rowPos As Long, ColumnPos As Long) As Integer
    On Error Resume Next
    Dim numColumn As Long
    numColumn = ColumnPos - 1
    If numColumn = 0 Then
        LstViewRowColumnIsColor = lstView.ListItems(rowPos).ForeColor
    Else
        LstViewRowColumnIsColor = lstView.ListItems(rowPos).ListSubItems(numColumn).ForeColor
    End If
    Err.Clear
End Function
Public Sub LstViewFilterNew(lstReport As Object, ByVal ColumnName As String, ByVal ColumnValue As String, Optional Remove As Integer = 0, Optional ByVal Delimiter As String = ",")
    On Error Resume Next
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim xCols As String
    Dim xPos As Long
    Dim spLine() As String
    Dim curValue As String
    xCols = LstViewColNames(lstReport)
    xPos = MvSearch(xCols, ColumnName, ",")
    If xPos = 0 Then Exit Sub
    ColumnValue = LCase$(ColumnValue)
    If ColumnValue = "(none)" Then ColumnValue = vbNullString
    rsTot = lstReport.ListItems.Count
    Call InitProgress(rsTot)
    For rsCnt = rsTot To 1 Step -1
        If UpdateProgress(rsCnt, rsTot, "Filtering report...") = False Then Exit For
        spLine = LstViewGetRow(lstReport, rsCnt)
        curValue = LCase$(Trim$(spLine(xPos)))
        If Remove = 0 Then
            If MvSearch(ColumnValue, curValue, Delimiter) = 0 Then
                lstReport.ListItems.Remove rsCnt
            End If
        Else
            If MvSearch(ColumnValue, curValue, Delimiter) > 0 Then
                lstReport.ListItems.Remove rsCnt
            End If
        End If
        Err.Clear
    Next
    LstViewAutoResize lstReport
    CloseProgress
    Err.Clear
End Sub
Public Function MvSearch(ByVal StringMv As String, ByVal StrLookFor As String, Optional ByVal Delim As String = vbNullString) As Long
    On Error Resume Next
    Dim TheFields() As String
    MvSearch = 0
    If Len(StringMv) = 0 Then
        MvSearch = 0
        Err.Clear
        Exit Function
    End If
    If Len(Delim) = 0 Then
        Delim = Chr$(253)
    End If
    Call StringParse(TheFields, StringMv, Delim)
    MvSearch = ArraySearch(TheFields, StrLookFor)
    Erase TheFields
    Err.Clear
End Function
Public Function MvField(ByVal strData As String, Optional ByVal fldPos As Long = 1, Optional ByVal Delim As String = vbNullString) As String
    On Error Resume Next
    Dim spData() As String
    Dim spCnt As Long
    MvField = vbNullString
    If Len(Delim) = 0 Then
        Delim = Chr$(253)
    End If
    If Len(strData) = 0 Then
        Err.Clear
        Exit Function
    End If
    Call StringParse(spData, strData, Delim)
    spCnt = UBound(spData)
    Select Case fldPos
    Case -1
        MvField = Trim$(spData(spCnt))
    Case Else
        If fldPos <= spCnt Then
            MvField = Trim$(spData(fldPos))
        End If
    End Select
    Erase spData
    Err.Clear
End Function
Public Function LstViewGetRowCode(ByVal StrValue As String) As String
    On Error Resume Next
    Dim arg1 As String
    Dim arg2 As String
    Dim arg3 As String
    Dim xCom As String
    If InStr(1, StrValue, "LstViewGetRow", vbTextCompare) > 0 Then
        If InStr(1, StrValue, "=") = 0 Then
            If InStr(1, StrValue, "Call ") > 0 Then
                StrValue = Replace$(StrValue, "Call", vbNullString)
                StrValue = Replace$(StrValue, "(", " ")
                StrValue = Replace$(StrValue, ")", vbNullString)
                StrValue = Trim$(StrValue)
            End If
            If InStr(1, StrValue, "Function ", vbTextCompare) > 0 Then
                GoTo IsFunction
                Err.Clear
                Exit Function
            End If
            arg1 = Trim$(MvField(StrValue, 1, ","))
            xCom = Trim$(MvField(arg1, 1, " "))
            arg1 = Trim$(MvField(arg1, 2, " "))
            arg2 = Trim$(MvField(StrValue, 2, ","))
            arg3 = Trim$(MvField(StrValue, 3, ","))
            StrValue = arg1 & " = " & xCom & "(" & arg2 & ", " & arg3 & ")"
        End If
    End If
IsFunction:
    LstViewGetRowCode = StrValue
    Err.Clear
End Function
Public Function StringBrowseForFolder(hWndOwner As Long, sPrompt As String) As String
    On Error Resume Next
    Dim iNull As Integer
    Dim lpIDList As Long
    Dim lResult As Long
    Dim sPath As String
    Dim udtBI As BrowseInfo
    With udtBI
        .hWndOwner = hWndOwner
        .lpszTitle = lstrcat(sPrompt, vbNullString)
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With
    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        lResult = SHGetPathFromIDList(lpIDList, sPath)
        Call CoTaskMemFree(lpIDList)
        iNull = InStr(sPath, vbNullChar)
        If iNull Then
            sPath = Left$(sPath, iNull - 1)
        End If
    End If
    StringBrowseForFolder = sPath
    Err.Clear
End Function
Public Sub CollectionOfFolders(ByVal StartDirectory As String, colPaths As Collection)
    On Error Resume Next
    Dim fso As New Scripting.FileSystemObject
    Dim objFolder As Scripting.Folder
    Dim objFolders As Scripting.Folders
    Dim objFiles As Scripting.Files
    Dim objEachFolder As Scripting.Folder
    Dim objEachFile As Scripting.File
    Set objFolder = fso.GetFolder(StartDirectory)
    Set objFolders = objFolder.SubFolders
    Set objFiles = objFolder.Files
    'For each subfolder in the Folder
    For Each objEachFolder In objFolders
        colPaths.Add objEachFolder
        'Do something with the Folder Name
        'Then recurse this function with the sub folder to get any sub-folders
        CollectionOfFolders objEachFolder, colPaths
        Err.Clear
    Next
    'For each folder check for any files
    For Each objEachFile In objFiles
        colPaths.Add objEachFile
        Err.Clear
    Next
    Set fso = Nothing
    Set objFolder = Nothing
    Set objFolders = Nothing
    Set objFiles = Nothing
    Set objEachFolder = Nothing
    Set objEachFile = Nothing
    Err.Clear
End Sub
Public Sub CleanAllControls(Thisform As Variant)
    On Error Resume Next
    Dim ctlControl As Variant
    For Each ctlControl In Thisform.Controls
        Select Case ctlControl.Name
        Case "cmbNames", "cmbLists", "cmbSort", "cmbWork"
            GoTo NextSection
        Case "cmbType", "cmbFormat", "cmbJustify", "cmbDepth"
            GoTo NextSection
        End Select
        ctlControl.Text = vbNullString
        ctlControl.Clear
        ctlControl.ListIndex = -1
        ctlControl.Value = 0
        ctlControl.Enabled = True
NextSection:
        Err.Clear
    Next
    Err.Clear
End Sub
Public Sub LstBoxFromMV(lstObj As Variant, ByVal StringMv As String, Optional ByVal Delim As String = vbNullString, Optional ByVal Sclear As String = vbNullString, Optional ByVal RemoveDups As String = vbNullString)
    On Error Resume Next
    Dim spDel() As String
    Dim spCnt As Long
    Dim wCnt As Long
    Dim xItm As String
    If Len(Delim) = 0 Then
        Delim = Chr$(253)
    End If
    If Len(Sclear) = 0 Then
        lstObj.Clear
    End If
    Call StringParse(spDel, StringMv, Delim)
    wCnt = UBound(spDel)
    For spCnt = 1 To wCnt
        xItm = StringProperCase(spDel(spCnt))
        If Len(xItm) = 0 Then
            GoTo NextLine
        End If
        Select Case StringFixText(RemoveDups)
        Case vbNullString, "Y"
            LstBoxUpdate lstObj, xItm
        Case Else
            lstObj.AddItem xItm
        End Select
NextLine:
        Err.Clear
    Next
    Erase spDel
    Err.Clear
End Sub
Public Function StringFixText(ByVal sString As String) As String
    On Error Resume Next
    StringFixText = UCase$(Trim$(sString))
    Err.Clear
End Function
Public Sub LstBoxUpdate(lstBox As Variant, ParamArray items())
    On Error Resume Next
    Dim Item As Variant
    For Each Item In items
        If LstBoxFindExactItem(lstBox, CStr(Item)) = -1 Then
            lstBox.AddItem CStr(Item)
        End If
        Err.Clear
    Next
    Set Item = Nothing
    Err.Clear
End Sub
Public Function LstBoxFindExactItem(lstBox As Variant, ByVal StrSearch As String) As Long
    On Error Resume Next
    Dim lstTot As Long
    Dim lstCnt As Long
    Dim lstStr As String
    LstBoxFindExactItem = -1
    If Len(StrSearch) = 0 Then
        Err.Clear
        Exit Function
    End If
    StrSearch = LCase$(StrSearch)
    lstTot = lstBox.ListCount - 1
    For lstCnt = 0 To lstTot
        lstStr = LCase$(lstBox.list(lstCnt))
        Select Case lstStr
        Case StrSearch
            LstBoxFindExactItem = lstCnt
            Exit For
        End Select
        Err.Clear
    Next
    Err.Clear
End Function
Public Function DialogOpen(ByVal Filter As String, Optional ByVal Title As String = vbNullString, Optional ByVal InitDir As String = vbNullString, Optional ByVal DefaultExt As String = vbNullString) As String
    On Error GoTo ErrHandler
    Dim filName As String
    Dim filCnt As Integer
    Dim spFilter() As String
    filCnt = 0
    If Len(DefaultExt) = 0 Then
        DefaultExt = "*.*"
    End If
    If Len(Title) = 0 Then
        Title = "Open an existing file"
    End If
    If Len(InitDir) = 0 Then
        InitDir = ReadReg(App.Title, "lastpath")
    End If
    If Len(InitDir) = 0 Then
        InitDir = "..."
    End If
    Call StringParse(spFilter, Filter, "|")
    filCnt = ArraySearch(spFilter, DefaultExt)
    If filCnt <> 0 Then
        filCnt = filCnt / 2
    End If
    filName = vbNullString
    With SetupConverter.CD
        .CancelError = True
        '.flags = OFN_HIDEREADONLY Or OFN_FILEMUSTEXIST
        .Filter = Filter
        .DialogTitle = Title
        .InitDir = InitDir
        .Filename = vbNullString
        .DefaultExt = DefaultExt
        .FilterIndex = filCnt
        .ShowOpen
        filName = .Filename
        If Len(filName) = 0 Then
            Err.Clear
            Exit Function
        End If
    End With
    DialogOpen = filName
    SaveReg "lastpath", StringGetFileToken(filName, "p")
    Err.Clear
    Exit Function
ErrHandler:
    Err.Clear
    Exit Function
    Err.Clear
End Function
Public Sub LstViewFromFile(lstView As Control, ByVal strFile As String, Optional ByVal lstClear As String = vbNullString)
    On Error Resume Next
    Dim sFileLin As String
    Dim xDetails() As String
    Dim nLine As String
    Dim xTag As String
    Dim xIcon As String
    Dim xSmallIcon As String
    Dim xKey As String
    Dim xLine() As String
    Dim fPos As Long
    Dim tRecs As Long
    Dim cRecs As Long
    Dim sRecs As String
    Dim fData() As String
    If Len(lstClear) = 0 Then
        lstView.ListItems.Clear
    End If
    sRecs = FileData(strFile)
    sRecs = StringRemAllNL(sRecs)
    If Len(sRecs) = 0 Then
        Err.Clear
        Exit Sub
    End If
    tRecs = StringParse(fData, sRecs, vbNewLine)
    If InitProgress(tRecs) = False Then
        Err.Clear
        Exit Sub
    End If
    For cRecs = 1 To tRecs
        sFileLin = fData(cRecs)
        If UpdateProgress(cRecs, tRecs, "List View From File: " & StringGetFileToken(strFile, "fo")) = False Then
            Exit For
        End If
        sFileLin = Trim$(sFileLin)
        Call StringParse(xDetails, sFileLin, Chr$(193))
        ReDim Preserve xDetails(5)
        nLine = xDetails(1)
        xTag = xDetails(2)
        xIcon = xDetails(3)
        xSmallIcon = xDetails(4)
        xKey = xDetails(5)
        Call StringParse(xLine, nLine, Chr$(254))
        fPos = LstViewUpdate(xLine, lstView, vbNullString)
        lstView.ListItems(fPos).EnsureVisible
        If Len(xTag) > 0 Then
            lstView.ListItems(fPos).Tag = xTag
        End If
        If Len(xIcon) > 0 Then
            lstView.ListItems(fPos).Icon = xIcon
        End If
        If Len(xSmallIcon) > 0 Then
            lstView.ListItems(fPos).SmallIcon = xSmallIcon
        End If
        If Len(xKey) > 0 Then
            lstView.ListItems(fPos).Key = xKey
        End If
        Err.Clear
    Next
    CloseProgress
    sFileLin = vbNullString
    Err.Clear
End Sub
Public Function FileData(ByVal Filename As String) As String
    On Error Resume Next
    Dim sLen As Long
    Dim myBuf As String
    Dim FileNum As Long
    Dim Size As Long
    FileNum = FreeFile
    Size = FileLen(Filename)
    myBuf = String$(Size, "*")
    Open Filename For Input Access Read As #FileNum
    sLen = LOF(FileNum)
    FileData = Input(sLen, #FileNum)
    Close #FileNum
    Err.Clear
End Function
Public Function StringRemAllNL(ByVal StrString As String) As String
    On Error Resume Next
    Dim StrSize As Long
    Dim LAST2 As String
    Dim TmpString As String
    TmpString = StrString
    LAST2 = Right$(TmpString, 2)
    Do While LAST2 = vbNewLine
        StrSize = Len(TmpString) - 2
        TmpString = Left$(TmpString, StrSize)
        LAST2 = Right$(TmpString, 2)
    Loop
    StringRemAllNL = TmpString
    LAST2 = vbNullString
    TmpString = vbNullString
    Err.Clear
End Function
Public Function boolViewFile(ByVal Filename As String, Optional ByVal Operation As String = "Open", Optional ByVal WindowState As Long = 1) As Boolean
    On Error Resume Next
    Dim r As Long
    r = lngStartDoc(Filename, Operation, WindowState)
    If r <= 32 Then
        ' there was an error
        Beep
        Resp = MyPrompt("An error occurred while opening your document." & vbCr & "The possibility is that the selected entry does not have" & vbCr & "a link in the registry to open it with.", "o", "w", "Viewer Error")
        boolViewFile = False
    Else
        boolViewFile = True
        Pause 1
    End If
    Err.Clear
End Function
Private Function lngStartDoc(ByVal Docname As String, Optional ByVal Operation As String = "Open", Optional ByVal WindowState As Long = 1) As Long
    On Error Resume Next
    Dim Scr_hDC As Long
    Dim sDir As String
    sDir = StringGetFileToken(Docname, "d")
    Scr_hDC = GetDesktopWindow()
    lngStartDoc = ShellExecute(Scr_hDC, Operation, Docname, vbNullString, sDir, WindowState)
    Err.Clear
End Function
Public Function MyPrompt(ByVal StrMsg As String, Optional ByVal strButton As String = "o", Optional ByVal StrIcon As String = "e", Optional ByVal StrHeading As String = vbNullString) As Variant
    On Error Resume Next
    ' button can be any of
    ' ync - yesnocancel, c - cancel, o - ok, oc - okcancel, rc - retrycancel and yn - yesno
    ' and ari - abortretryignore, bc - backclose, bnc - backnextclose
    ' bns - backnextsnooze, nc - nextclose, sc - searchclose, toc - tipsoptionsclose, yanc - yesallnocancel
    ' icon can be any of
    ' i - information, w - warning, c - critical, t - tip, q - query
    ' mode can be any of
    ' ad - autodown, ma - modal, me - modeless
    Dim isCheck As Long
    Dim Mode As Long
    Dim Button As Long
    Dim Icon As Long
    ' see if excel is already running
    If Len(StrHeading) = 0 Then
        StrHeading = App.Title
    End If
    isCheck = 0
    Mode = vbApplicationModal
    Select Case LCase$(strButton)
    Case "ync"
        Button = vbYesNoCancel
    Case "c"
        Button = vbCancel
    Case "o"
        Button = vbOKOnly
    Case "oc"
        Button = vbOKCancel
    Case "rc"
        Button = vbRetryCancel
    Case "yn"
        Button = vbYesNo
    Case "ari"
        Button = vbAbortRetryIgnore
    End Select
    Select Case LCase$(StrIcon)
    Case "i", "t"
        Icon = vbInformation
    Case "w", "e"
        Icon = vbExclamation
    Case "c"
        Icon = vbCritical
    Case "q"
        Icon = vbQuestion
    End Select
    MyPrompt = MsgBox(StrMsg, Button + Icon + Mode, StrHeading)
    Err.Clear
End Function
Public Sub Pause(ByVal nSecond As Double)
    On Error Resume Next
    ' call pause(2)      delay for 2 seconds
    Dim t0 As Double
    'Dim dummy As Integer
    t0 = Timer
    Do While Timer - t0 < nSecond
        DoEvents
        ' if we cross midnight, back up one day
        If Timer < t0 Then
            t0 = t0 - CLng(24) * CLng(60) * CLng(60)
        End If
    Loop
    Err.Clear
End Sub
Public Function boolDirExists(ByVal Sdirname As String) As Boolean
    On Error Resume Next
    Dim sDir As String
    boolDirExists = False
    sDir = Dir$(Sdirname, vbDirectory)
    If (Len(sDir) > 0) And (Err = 0) Then
        boolDirExists = True
    End If
    sDir = vbNullString
    Err.Clear
End Function
Public Function MvFromCollection(objCollection As Collection, Optional ByVal Delim As String = vbNullString) As String
    On Error Resume Next
    Dim xTot As Long
    Dim xCnt As Long
    Dim sRet As String
    sRet = vbNullString
    If Delim = vbNullString Then
        Delim = Chr$(253)
    End If
    xTot = objCollection.Count
    For xCnt = 1 To xTot
        sRet = sRet & objCollection.Item(xCnt) & Delim
        Err.Clear
    Next
    MvFromCollection = StringRemoveDelim(sRet, Delim)
    Err.Clear
End Function
Public Function StringRemoveDelim(ByVal strData As String, Optional ByVal Delim As String = vbNullString) As String
    On Error Resume Next
    Dim intDataSize As Long
    Dim intDelimSize As Long
    Dim strLast As String
    If Len(Delim) = 0 Then
        Delim = VM
    End If
    intDataSize = Len(strData)
    intDelimSize = Len(Delim)
    strLast = Right$(strData, intDelimSize)
    Select Case strLast
    Case Delim
        StringRemoveDelim = Left$(strData, (intDataSize - intDelimSize))
    Case Else
        StringRemoveDelim = strData
    End Select
    strLast = vbNullString
    Err.Clear
End Function
Public Function StringClean(ByVal StrSource As String) As String
    On Error Resume Next
    Dim strRslt As String
    strRslt = StrSource
    strRslt = Replace$(strRslt, vbTab, " ")
    strRslt = Replace$(strRslt, vbCrLf, " ")
    Do While (InStr(strRslt, "  "))
        strRslt = Replace$(strRslt, "  ", " ")
    Loop
    StringClean = Trim$(strRslt)
    Err.Clear
End Function
Public Sub FileUpdate(ByVal filName As String, ByVal Fillines As String, Optional ByVal Wora As String = vbNullString)
    On Error Resume Next
    Dim iFileNum As Integer
    Wora = StringFixText(Wora)
    If Len(Wora) = 0 Then
        Wora = "W"
    End If
    iFileNum = FreeFile
    Select Case Wora
    Case "W"
        Open filName For Output As #iFileNum
    Case "A"
        Open filName For Append As #iFileNum
    End Select
    Print #iFileNum, Fillines
    Close #iFileNum
    Err.Clear
End Sub
Private Sub ZipFile_CreateEmpty(ByVal sPath As String)
    On Error Resume Next
    'Create empty Zip File
    Dim oFSO As Scripting.FileSystemObject
    Dim arrHex As Variant
    Dim sBin As Variant
    Dim i As Long
    Set oFSO = New Scripting.FileSystemObject
    arrHex = Array(80, 75, 5, 6, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    For i = 0 To UBound(arrHex)
        sBin = sBin & Chr$(arrHex(i))
        Err.Clear
    Next
    With oFSO.CreateTextFile(sPath, True)
        .Write sBin
        .Close
    End With
    Err.Clear
End Sub
Public Function ZipFileXP(ByVal strFile As String, ByVal strFiles As String) As String
    On Error Resume Next
    Dim strPath As String
    Dim rsCnt As Long
    Dim rsTot As Long
    Dim spFiles() As String
    Dim oApp As Shell32.Shell
    strPath = StringGetFileToken(strFile, "p") & "\" & StringGetFileToken(strFile, "fo") & ".zip"
    ' create an empty zip file
    Do Until boolFileExists(strPath) = True
        ZipFile_CreateEmpty strPath
        DoEvents
    Loop
    Set oApp = New Shell32.Shell
    spFiles = Split(strFiles, ";")
    rsTot = UBound(spFiles) + 1
    Call InitProgress(rsTot)
    For rsCnt = 0 To rsTot
        If UpdateProgress(rsCnt + 1, rsTot, spFiles(rsCnt) & ".Compressing") = False Then
            Exit For
        End If
        If boolFileExists(spFiles(rsCnt)) = True Then oApp.Namespace(strPath).CopyHere spFiles(rsCnt)
        DoEvents
        Err.Clear
    Next
    Set oApp = Nothing
    If boolFileExists(strPath) = True Then
        ZipFileXP = strPath
    Else
        ZipFileXP = vbNullString
    End If
    CloseProgress
    Err.Clear
End Function
'--- Appends "\" to the file path ---
Public Function ToPath(ByVal sPath As String) As String
    On Error Resume Next
    If sPath <> vbNullString Then If Right$(sPath, 1) <> "\" Then sPath = sPath + "\"
    ToPath = sPath
    Err.Clear
End Function
'--- Extracts the file name from a full path ---
Public Function GetFileName(sPath As String) As String
    On Error Resume Next
    Dim X
    X = InStrRev(sPath, "/")
    If X = 0 Then
        X = InStrRev(sPath, "\")
        If X = 0 Then GetFileName = sPath: Exit Function
    End If
    GetFileName = Mid$(sPath, X + 1)
    Err.Clear
End Function
'--- Extracts directory name from a path ---
Public Function GetDirName(sPath As String) As String
    On Error Resume Next
    Dim X
    X = InStrRev(sPath, "/")
    If X = 0 Then
        X = InStrRev(sPath, "\")
        If X = 0 Then GetDirName = sPath: Exit Function
    End If
    GetDirName = Left$(sPath, X - 1)
    Err.Clear
End Function
Public Sub AddInErr(obErr As ErrObject, Optional sCaption As String = vbNullString)
    On Error Resume Next
    'write to error log
    Open App.Path + "\errors.log" For Append As #1
    Print #1, CStr(Now); " // "; sCaption
    Print #1, "    Error: " + obErr.Description
    Print #1, "====================================="
    Close #1
    'comment the next line if you don't wish any error messages displayed
    MsgBox obErr.Description, vbCritical, "Add-in Error"
    Err.Clear
End Sub

Private Sub tvExtractIcon(Filename As String, picIcon As PictureBox)
    On Error Resume Next
    Dim Icon As Long
    Icon = SHGetFileInfo(Filename, 0&, IFileInfo, Len(IFileInfo), IFlags Or SHGFI_SMALLICON)
    If Icon <> 0 Then
        Set picIcon.Picture = LoadPicture()
        Icon = ImageList_Draw(Icon, IFileInfo.iIcon, picIcon.hDC, 0, 0, ILD_TRANSPARENT)
    End If
    Err.Clear
End Sub
Public Function tvAddIconToIML(ByVal Filename As String, ByVal FType As String, imgList As ImageList, picIcon As PictureBox) As Long
    On Error Resume Next
    ' add an image of a file to an image list
    ' the file type is the extension of the file
    Dim i As Long
    Dim i_Tot As Long
    If IsNumeric(FType) Then
        FType = "XXX"
    End If
    If LCase$(FType) = "exe" Or LCase$(FType) = "ico" Then
        Call tvExtractIcon(Filename, picIcon)
        tvAddIconToIML = imgList.ListImages.Add(, , picIcon.Image).Index
    Else
        i_Tot = imgList.ListImages.Count
        For i = 1 To i_Tot
            If imgList.ListImages(i).Key = FType Then
                tvAddIconToIML = i
                Err.Clear
                Exit Function
            End If
            Err.Clear
        Next
        Call tvExtractIcon(Filename, picIcon)
        tvAddIconToIML = imgList.ListImages.Add(, FType, picIcon.Image).Index
    End If
    Err.Clear
End Function

