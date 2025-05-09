Attribute VB_Name = "Mdl_Stadmis"
'****************************************************************************************************
'*  Author:             Masika .S. Elvas                                                            *
'*  Gender:             Male                                                                        *
'*  Postal Address:     P.O Box 137, BUNGOMA 50200, KENYA                                           *
'*  Phone No:           (254) 724 688 172 / (254) 751 041 184                                       *
'*  E-mail Address:     maselv_e@yahoo.co.uk / elvasmasika@lexeme-kenya.com / masika_elvas@live.com *
'*  Location            BUNGOMA, KENYA                                                              *
'****************************************************************************************************

Option Explicit

Public Const def_BackupDatabaseExt = "bkd"
Public Const def_BackupSettingsExt = "bks"

                              '    1           2          3      4      5     6      7       8      9    10    11         12
Public Const def_Privileges = "111111111|1111111111111|1111111|11111|111111|11111|1111111|1111111|11111|11111|11111|11111111111111"
Public Const def_SecurityQuestions1 = "What is the first name of your favourite aunt?|What is the first name of your favourite uncle?|Where did you meet your spouse?|What is your eldest cousin's name|What is your youngest child's nickname?|What is your eldest child's nickname?|What is the first name of your eldest niece?|What is the first name of your eldest nephew?|Where did you spend your honeymoon?"
Public Const def_SecurityQuestions2 = "What was your favourite food as a child?|What was the surname of your first boss?|What is the name of the hospital in which you were born?|What is your first pet's name?|What is the first name of your favourite musician?|What is the name of your favourite movie?|What was the make of your first car?"

'Play Sound
Private Const SND_SYNC = &H0
Private Const SND_ASYNC = &H1
Private Const SND_NODEFAULT = &H2
Private Const SND_LOOP = &H8
Private Const SND_NOSTOP = &H10

Private Const LB_FINDSTRING = &H18F
Private Const CB_FINDSTRING = &H14C

Private Const WM_COMMAND = &H111
Private Const MIN_ALL = &H1A3
Private Const MIN_ALL_UNDO = &H1A0

Private Const Lvm_First = &H1000
Private Const Lvm_SetColumnWidth = Lvm_First + 30
Private Const Lvm_AutoSize = -&H1
Private Const Lvm_AutoSize_UseHeader = -&H2

Private Const LvNumber = &H0
Private Const LvText = &H1
Private Const LvDateTime = &H2
Private Const LvCurrency = &H3

Private Const EM_GETLINECOUNT = 186

Public Enum LvwSortOrder
    
    Ascending = &H0 'If Error, Add this Component 'Microsoft Windows Common Controls 6.0 (SP6)'
    Descending = &H1
    
End Enum

Public Enum vMsgBoxButtons
    
    vbOKOnly = &H0              '0
    vbOKCancel = &H1            '1
    vbAbortRetryIgnore = &H2    '2
    vbYesNoCancel = &H3         '3
    vbYesNo = &H4               '4
    vbRetryCancel = &H5         '5
    vbCustomButtons = &H6       '6
    
    vbBuzzer = -&H1         '-1
    vbCritical = &H10       '16
    vbQuestion = &H20       '32
    vbExclamation = &H30    '48
    vbInformation = &H40    '64
    
    vbDefaultButton1 = &H0      '0
    vbDefaultButton2 = &H100    '256
    vbDefaultButton3 = &H200    '512
    vbDefaultButton4 = &H300    '768
    
End Enum

Public Enum vMsgBoxReturnValues
    
    vbOK = &H1      '1
    vbCancel = &H2  '2
    vbAbort = &H3   '3
    vbRetry = &H4   '4
    vbIgnore = &H5  '5
    vbYes = &H6     '6
    vbNo = &H7      '7
    
End Enum

Public Type SoftwareLicences
    
    License_Code As String
    License_Encrypted As String
    Expiry_Date As Date
    Max_Users As Long
    Key As String
    
End Type

Public Type SoftwareSettings
    
    Term As Integer
    Min_Year As Long
    Max_App_Logs As Long
    Academic_Year As Date
    Min_Subjects As Integer
    SoftwareSound As Boolean
    Max_Login_Records As Long
    MsgBoxSoundFldr As String
    MyRecentFiles(&H1) As Long
    Licences As SoftwareLicences
    Show_Splash_Screen As Boolean
    Min_Username_Characters As Long
    RequestAlternativeLogin As Boolean
    Min_User_Password_Characters As Long
    
End Type

Public Type SchoolData
    
    ID As Long
    Code As Long
    Name As String
    Type As Integer
    Modules As String
    Gender As Integer
    PhoneNo As String
    Website As String
    Location As String
    Motto As String
    NHIFNo As String
    NSSFNo As String
    Logo As StdPicture
    EmailAddress As String
    PostalAddress As String
    
End Type

Public Type StudentData
    
    iCAT As Long
    iTerm As Long
    iAdm_No As Long
    iStudentName As String
    iClass As String
    iClassID As String
    iStream As String
    iStreamID As String
    iExam_ID As Long
    iExam_Year As Long
    iExam_Name As String
    iOpeningDate As Date
    idDaysAbsent As Long
    idDaysPresent As Long
    iFeeArrears As Double
    iFeeThisTerm As Double
    iFeePaid As Double
    iFeeNextTerm As Double
    iNoofCATs As Integer
    
End Type

Public Type DeviceInfo
    
    Account_Name As String
    Name As String
    Serial_No As String
    
End Type

Public Type UserInfo
    
    Login_Name As String
    Login_Date As Date
    Login_ID As Long
    User_ID As Long
    User_Name As String
    Full_Name As String
    Account_Name As String
    Hierarchy As Long
    Privileges As String
    
    Parental_Control_ON As Boolean
    Start_Time As Date
    End_Time As Date
    
    Device_Name As String
    Device_Account_Name As String
    Device_Serial_No As String
    
End Type

Public Type tThemes
    
    tBackColor As OLE_COLOR
    tForeColor As OLE_COLOR
    tImagePicture As Picture
    tLvColorOne As OLE_COLOR
    tLvColorTwo As OLE_COLOR
    tButtonBackColor As OLE_COLOR
    tButtonForeColor As OLE_COLOR
    tButtonGradientColor As OLE_COLOR
    tWarningForeColor As OLE_COLOR
    tEntryColor As OLE_COLOR
    tDarkColor As OLE_COLOR
    tProgBarColor As OLE_COLOR
    
End Type

Public Type sAdditionalPhotos
    
    'Picture in bytes
    vDataBytes() As Byte
    
End Type

'Picture byte arrays
Public sAdditionalPhoto(&H4) As sAdditionalPhotos

Public vFrm(&H1) As Form

Public User As UserInfo
Public VirtualUser As UserInfo
Public Device As DeviceInfo
Public tTheme As tThemes
Public School As SchoolData
Public iData As StudentData
Public SoftwareSetting As SoftwareSettings

'Bounded Arrays
Public vIndex(&H4) As Long
Public vBuffer(&H4) As String
Public vSelectedButton(&H3) As String
Public DontShowMsgAgain(&HA, &H1) As Long

'Unbounded Arrays
Public vArrayList() As String
Public vArrayListTmp() As String
Public vMultiSelectedData As String

'Longs
Public kTrialPeriod&, vFrmWidth&, iSchoolID&

'Strings
Public iTimeStamp$
Public vSearchCriteria$
Public vEditRecordID$, vSelTbl$
Public def_BackupLocation$, def_LogFileLocation$

Public vShowMsgBox, vStartupComplete, vRegistering As Boolean

'ADODB
Public vRs As New ADODB.Recordset
Public vRsTmp As New ADODB.Recordset
Public vAdoCNN As New ADODB.Connection

Public vFso As New FileSystemObject

Public vCancelOperation%
Public vDatabaseAltered, iNotFullyLoaded, vSilentClosure, vWait, vRegistered, iMsgBoxDisplayed, iReportGenerated As Boolean

Public Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)
Private Declare Sub ReleaseCapture Lib "user32" ()

Public Declare Function SleepEx Lib "Kernel32" (ByVal dwMilliseconds As Long, ByVal bAlertable As Long) As Long

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Any) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowsDirectory Lib "Kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function SHChangeNotify Lib "shell32.dll" (ByVal wEventID As Long, ByVal uFlags As Long, ByVal dwItem1 As Long, ByVal dwItem2 As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Declare Function SendMessageB Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageAsLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'Play Sound
Private Declare Function waveOutGetNumDevs Lib "winmm" () As Long
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

'Determine whether a file is already open or not
Private Declare Function lOpen Lib "Kernel32" Alias "_lopen" (ByVal lpPathName As String, ByVal iReadWrite As Long) As Long
Private Declare Function lClose Lib "Kernel32" Alias "_lclose" (ByVal hFile As Long) As Long

Public Function vMsgBox(ByVal mPrompt$, Optional ByVal mButtons As vMsgBoxButtons = vbOKOnly, Optional ByVal mAppTitle$ = VBA.vbNullString, Optional iFrm As Object, Optional ByVal vMsgBoxPromptSize% = &H8, Optional ByVal vMsgBoxPromptFontName$ = "Tahoma", Optional ByVal vAutoCloseDuration& = &H0, Optional ByVal mWarning As Boolean = False, Optional DontShowMsgStr_Visible_Key_Setting As String = VBA.vbNullString, Optional CustomButtons As String = VBA.vbNullString, Optional vUseDefault As Boolean = True) As vMsgBoxReturnValues
On Local Error Resume Next
    
    Dim mArrayList() As String
    
    Dim UsingCustomButtons As Boolean
    Dim MyMsgType&, MyButtons&, MyDefButton&
    
    Frm_MsgBox.ChkDontShowMsg.Visible = False
    
    mArrayList = VBA.Split(DontShowMsgStr_Visible_Key_Setting, "|")
    
    If UBound(mArrayList) >= &H1 Then
        
        Frm_MsgBox.ChkDontShowMsg.Visible = VBA.CBool(VBA.CLng(mArrayList(&H0)))
        
        'If an option has been specified then assign it
        If UBound(mArrayList) >= &H2 Then DontShowMsgAgain(mArrayList(&H1), &H0) = mArrayList(&H2)
        
        'If the User had selected not to display this type of message again then...
        If VBA.CBool(DontShowMsgAgain(mArrayList(&H1), &H0)) Then
            
            'Get the button initially selected by User
            vSelectedButton(&H0) = DontShowMsgAgain(mArrayList(&H1), &H1)
            
            'Assign automatically the selected button
            vMsgBox = vSelectedButton(&H0)
            
            'Change Mouse Pointer to show end of Processing state
            Screen.MousePointer = vbDefault
            
            Exit Function 'Quit this Function
            
        End If 'Close respective IF..THEN block statement
        
        Frm_MsgBox.ChkDontShowMsg.DataField = mArrayList(&H1)
        
    End If 'Close respective IF..THEN block statement
    
    'Set Mouse pointer to indicate beginning of process or operation
    Screen.MousePointer = vbHourglass
    
    MyButtons = mButtons
    
    Select Case mButtons
        
        Case &H0, &H10, &H20, &H30, &H40, &H100, &H110, &H120, &H130, &H140, &H200, &H210, &H220, &H230, &H240, &H300, &H310, &H320, &H330, &H340: MyButtons = vbOKOnly
        Case &H1, &H11, &H21, &H31, &H41, &H101, &H111, &H121, &H131, &H141, &H201, &H211, &H221, &H231, &H241, &H301, &H311, &H321, &H331, &H341: MyButtons = vbOKCancel
        Case &H2, &H12, &H22, &H32, &H42, &H102, &H112, &H122, &H132, &H142, &H202, &H212, &H222, &H232, &H242, &H302, &H312, &H322, &H332, &H342: MyButtons = vbAbortRetryIgnore
        Case &H3, &H13, &H23, &H33, &H43, &H103, &H113, &H123, &H133, &H143, &H203, &H213, &H223, &H233, &H243, &H303, &H313, &H323, &H333, &H343: MyButtons = vbYesNoCancel
        Case &H4, &H14, &H24, &H34, &H44, &H104, &H114, &H124, &H134, &H144, &H204, &H214, &H224, &H234, &H244, &H304, &H314, &H324, &H334, &H344: MyButtons = vbYesNo
        Case &H5, &H15, &H25, &H35, &H45, &H105, &H115, &H125, &H135, &H145, &H205, &H215, &H225, &H235, &H245, &H305, &H315, &H325, &H335, &H345: MyButtons = vbRetryCancel
        Case Is > &H0: UsingCustomButtons = True
        
    End Select 'Close SELECT..CASE block statement
    
    Select Case mButtons
        
        Case &H10 To &H16, &H110 To &H116, &H210 To &H216, &H310 To &H316: MyMsgType = vbCritical
        Case &H20 To &H26, &H120 To &H126, &H220 To &H226, &H320 To &H326: MyMsgType = vbQuestion
        Case &H30 To &H36, &H130 To &H136, &H230 To &H236, &H330 To &H336: MyMsgType = vbExclamation
        Case &H6, &H40 To &H46, &H140 To &H146, &H240 To &H246, &H340 To &H346: MyMsgType = vbInformation
        
    End Select 'Close SELECT..CASE block statement
    
    Select Case mButtons
        
        Case Is >= vbDefaultButton4: MyDefButton = vbDefaultButton4
        Case Is >= vbDefaultButton3: MyDefButton = vbDefaultButton3
        Case Is >= vbDefaultButton2: MyDefButton = vbDefaultButton2
        Case Is >= vbDefaultButton1: MyDefButton = vbDefaultButton1
        Case Else: MyDefButton = -&H1
        
    End Select 'Close SELECT..CASE block statement
    
    Frm_MsgBox.DisplayWarning = mWarning
    Frm_MsgBox.iWarningType = MyMsgType
    
    'Set Mouse pointer to indicate beginning of process or operation
    Screen.MousePointer = vbHourglass
    
    With Frm_MsgBox
        
        vShowMsgBox = True
        
        .Caption = VBA.IIf(VBA.Len(VBA.Trim$(mAppTitle)) = &H0, App.Title, mAppTitle) 'Assign vMsgBox title
        
        .TxtPrompt.FontName = vMsgBoxPromptFontName
        .TxtPrompt.FontSize = vMsgBoxPromptSize
        
        'Assign vMsgBox Image
        '---------------------------------------------------------------
        If MyMsgType = &H0 Then .MsgBoxImg.Picture = Nothing Else .MsgBoxImg.Picture = .MsgBoxImgLst.ListImages((MyMsgType / 16)).Picture
        
        .TxtPrompt.Alignment = &H0
        
        .TxtPrompt.Text = mPrompt 'Assign vMsgBox Information
        .iMsg = mPrompt
        
        If .TextWidth(.TxtPrompt.Text) < 1000 Then
            .TxtPrompt.Width = 4000
        ElseIf .TextWidth(.TxtPrompt.Text) > 9500 Then
            .TxtPrompt.Width = 8000
        Else
            .TxtPrompt.Width = .TextWidth(.TxtPrompt.Text) - 200
        End If
        
        Dim i#
        
        i = SendMessageAsLong(.TxtPrompt.hWnd, EM_GETLINECOUNT, 0, 0)
        .TxtPrompt.Height = 210 * VBA.Val(i)
        
        If .TxtPrompt.Height > &H3E8 Then .TxtPrompt.Height = &H3E8
        
        'Resize the vMsgBox Form to fit its contents
        '---------------------------------------------------------------
        
        If .MsgBoxImg.Picture = &H0 Then .TxtPrompt.Left = &H64 Else .TxtPrompt.Left = .MsgBoxImg.Left + .MsgBoxImg.Width + 100
        
        .Fra_Details.Width = .TxtPrompt.Left + .TxtPrompt.Width + 100
        
        If .TxtPrompt.Height < .MsgBoxImg.Height And .MsgBoxImg.Picture <> &H0 Then
            .Fra_Details.Height = .MsgBoxImg.Top + .MsgBoxImg.Height + &HFA
        Else
            .Fra_Details.Height = .TxtPrompt.Top + .TxtPrompt.Height + &HFA
        End If 'Close respective IF..THEN block statement
        
        .ImgFooter.Top = .Fra_Details.Top + .Fra_Details.Height + 150
                
        '---------------------------------------------------------------
        
        .ShpBttnButton(&H0).Visible = True
        
        For vIndex(&H0) = &H1 To .ShpBttnButton.UBound
            Unload .ShpBttnButton(vIndex(&H0))
        Next vIndex(&H0)
        
        If vAutoCloseDuration < &H1 Then
            
            'Assign vMsgBox Buttons
            '---------------------------------------------------------------
            
            .ShpBttnButton(&H0).Visible = True
            .ShpBttnButton(&H0).Left = (.Fra_Details.Left + .Fra_Details.Width) - .ShpBttnButton(&H0).Width
            .ShpBttnButton(&H0).Top = .ImgFooter.Top + ((.ImgFooter.Height / 2) - (.ShpBttnButton(&H0).Height / 2)) - 50
            
            Dim nButtons&
            
            nButtons = &H1
            
            If UsingCustomButtons Then GoTo CreateButtons
            
            Select Case MyButtons
                
                Case &H0:
                    .ShpBttnButton(&H0).Caption = "&OK": .ShpBttnButton(&H0).Tag = vbOK & ":O"
                    
                Case vbOKCancel:
                    
                    .ShpBttnButton(&H0).Caption = "&Cancel": .ShpBttnButton(&H0).Tag = vbCancel & ":C:Cancel"
                    Load .ShpBttnButton(&H1): .ShpBttnButton(&H1).Caption = "&OK": .ShpBttnButton(&H1).Top = .ShpBttnButton(&H0).Top
                    .ShpBttnButton(&H1).Left = .ShpBttnButton(&H0).Left - .ShpBttnButton(&H1).Width: .ShpBttnButton(&H1).Visible = True
                    .ShpBttnButton(&H1).Tag = vbOK & ":O:OK": nButtons = &H2
                    
                Case vbYesNo:
                    
                    .ShpBttnButton(&H0).Caption = "&No"
                    Load .ShpBttnButton(&H1): .ShpBttnButton(&H1).Caption = "&Yes": .ShpBttnButton(&H1).Top = .ShpBttnButton(&H0).Top
                    .ShpBttnButton(&H0).Tag = vbNo & ":N:No": .ShpBttnButton(&H1).Tag = vbYes & ":Y:Yes": .ShpBttnButton(&H1).Visible = True
                    .ShpBttnButton(&H1).Left = .ShpBttnButton(&H0).Left - .ShpBttnButton(&H1).Width: nButtons = &H2
                     
                Case vbRetryCancel:
                    
                    .ShpBttnButton(&H0).Caption = "&Cancel": .ShpBttnButton(&H0).Tag = vbCancel & ":C:Cancel"
                    Load .ShpBttnButton(&H1): .ShpBttnButton(&H1).Caption = "&Retry": .ShpBttnButton(&H1).Top = .ShpBttnButton(&H0).Top
                    .ShpBttnButton(&H1).Left = .ShpBttnButton(&H0).Left - .ShpBttnButton(&H1).Width: .ShpBttnButton(&H1).Tag = vbRetry & ":R:Retry"
                    .ShpBttnButton(&H1).Visible = True: nButtons = &H2
                     
                Case vbYesNoCancel:
                    
                    .ShpBttnButton(&H0).Caption = "&Cancel": nButtons = &H3
                    Load .ShpBttnButton(&H1): .ShpBttnButton(&H1).Caption = "&No": .ShpBttnButton(&H1).Top = .ShpBttnButton(&H0).Top
                    Load .ShpBttnButton(&H2): .ShpBttnButton(&H2).Caption = "&Yes": .ShpBttnButton(&H2).Top = .ShpBttnButton(&H0).Top
                    .ShpBttnButton(&H1).Left = .ShpBttnButton(&H0).Left - .ShpBttnButton(&H1).Width: .ShpBttnButton(&H1).Visible = True
                    .ShpBttnButton(&H2).Left = .ShpBttnButton(&H1).Left - .ShpBttnButton(&H2).Width: .ShpBttnButton(&H2).Visible = True
                    .ShpBttnButton(&H0).Tag = vbCancel & ":C:Cancel": .ShpBttnButton(&H1).Tag = vbNo & ":N:No": .ShpBttnButton(&H2).Tag = vbYes & ":Y:Yes"
                     
                Case vbAbortRetryIgnore:
                    
                    .ShpBttnButton(&H0).Caption = "&Ignore": nButtons = &H3
                    Load .ShpBttnButton(&H1): .ShpBttnButton(&H1).Caption = "&Retry": .ShpBttnButton(&H1).Top = .ShpBttnButton(&H0).Top
                    Load .ShpBttnButton(&H2): .ShpBttnButton(&H2).Caption = "&Abort": .ShpBttnButton(&H2).Top = .ShpBttnButton(&H0).Top
                    .ShpBttnButton(&H1).Left = .ShpBttnButton(&H0).Left - .ShpBttnButton(&H1).Width: .ShpBttnButton(&H1).Visible = True
                    .ShpBttnButton(&H2).Left = .ShpBttnButton(&H1).Left - .ShpBttnButton(&H2).Width: .ShpBttnButton(&H2).Visible = True
                    .ShpBttnButton(&H0).Tag = vbIgnore & ":I:Ignore": .ShpBttnButton(&H1).Tag = vbRetry & ":R:Retry": .ShpBttnButton(&H2).Tag = vbAbort & ":A:Abort"
                    
                Case vbCustomButtons:
CreateButtons:
                    Dim iButtons&, iButton&
                    Dim vButtons() As String
                    
                    Dim iButtonCaption$
                    Dim CurrentIndex&, iLeftPos&, MaxButton&, MinButton&
                    
                    MinButton = -&H1
                    vButtons = VBA.Split(CustomButtons, "|")
                    nButtons = VBA.IIf(UBound(vButtons) > &H3, &H3, UBound(vButtons)) + &H1
                    
                    .ShpBttnButton(&H0).Caption = VBA.Replace(VBA.Left$(vButtons(nButtons - &H1), &H14), "-", "")
                    .ShpBttnButton(&H0).Width = .TextWidth(VBA.Replace(VBA.Left$(vButtons(nButtons - &H1), &H14), "-", "")) + 100
                    .ShpBttnButton(&H0).Left = (.Fra_Details.Left + .Fra_Details.Width) - .ShpBttnButton(&H0).Width
                    .ShpBttnButton(&H0).Tag = "0" & VBA.IIf(VBA.InStr(vButtons(&H0), "&") <> &H0, ":" & VBA.Mid$(vButtons(&H0), VBA.InStr(vButtons(&H0), "&") + &H1, &H1), VBA.vbNullString) & ":" & VBA.Replace(.ShpBttnButton(&H0).Caption, "&", VBA.vbNullString)
                    .ShpBttnButton(&H0).Visible = (VBA.InStr(VBA.Left$(vButtons(nButtons - &H1), &HF), "-") = &H0)
                    If (VBA.InStr(VBA.Left$(vButtons(nButtons - &H1), &HF), "-") = &H0) Then MaxButton = &H0: MinButton = &H0: iLeftPos = .ShpBttnButton(&H0).Left - 50 Else iLeftPos = .ShpBttnButton(&H0).Left + .ShpBttnButton(&H0).Width
                    
                    For iButton = nButtons - &H1 To &H1 Step -&H1
                        
                        CurrentIndex = nButtons - iButton
                        iButtonCaption = VBA.Left$(vButtons(nButtons - CurrentIndex - &H1), &H14)
                        
                        Load .ShpBttnButton(CurrentIndex): .ShpBttnButton(CurrentIndex).Caption = VBA.Replace(iButtonCaption, "-", ""): .ShpBttnButton(CurrentIndex).Top = .ShpBttnButton(&H0).Top
                        .ShpBttnButton(CurrentIndex).Width = VBA.IIf(.TextWidth(iButtonCaption) < 600, .TextWidth(iButtonCaption) + (1000 - .TextWidth(iButtonCaption)), .TextWidth(iButtonCaption) + 100)
                        .ShpBttnButton(CurrentIndex).Left = iLeftPos - .ShpBttnButton(CurrentIndex).Width: .ShpBttnButton(CurrentIndex).Visible = True
                        .ShpBttnButton(CurrentIndex).Tag = (nButtons - iButton) & VBA.IIf(VBA.InStr(vButtons(CurrentIndex), "&") <> &H0, ":" & VBA.Mid$(vButtons(CurrentIndex), VBA.InStr(vButtons(CurrentIndex), "&") + &H1, &H1), VBA.vbNullString) & ":" & VBA.Replace(.ShpBttnButton(&H0).Caption, "&", VBA.vbNullString)
                        .ShpBttnButton(CurrentIndex).Visible = (VBA.InStr(iButtonCaption, "-") = &H0)
                        If (VBA.InStr(iButtonCaption, "-") = &H0) Then iLeftPos = .ShpBttnButton(CurrentIndex).Left - 50: If MaxButton = &H0 Then MaxButton = CurrentIndex: If MinButton < &H0 Then MinButton = CurrentIndex
                        
                    Next iButton
                    
                    .TxtPrompt.Width = (.ShpBttnButton(MinButton).Left + .ShpBttnButton(MinButton).Width) - .TxtPrompt.Left - .Fra_Details.Left - 150
                    
            End Select 'Close SELECT..CASE block statement
            
            '---------------------------------------------------------------
            
            If (.ShpBttnButton(MaxButton).Left < .Fra_Details.Left + &H5A0) Then
                
                .ShpBttnButton(.ShpBttnButton.UBound).Left = .Fra_Details.Left + &H604
                
                Select Case .ShpBttnButton.UBound
                    
                    Case &H2:
                            .ShpBttnButton(&H1).Left = .ShpBttnButton(&H2).Left + .ShpBttnButton(&H2).Width
                            .ShpBttnButton(&H0).Left = .ShpBttnButton(&H1).Left + .ShpBttnButton(&H1).Width
                            
                    Case &H1:
                            .ShpBttnButton(&H0).Left = .ShpBttnButton(&H1).Left + .ShpBttnButton(&H1).Width
                    
                End Select 'Close SELECT..CASE block statement
                
                .Fra_Details.Width = .ShpBttnButton(&H0).Left + .ShpBttnButton(&H0).Width - &H64
                
            End If 'Close respective IF..THEN block statement
            
        Else
            
            .ShpBttnButton(&H0).Left = .Fra_Details.Left + .Fra_Details.Width - .ShpBttnButton(&H0).Width
            .ShpBttnButton(&H0).Top = .Fra_Details.Top + .Fra_Details.Height + &H64
            
            .DisplayDuration = vAutoCloseDuration
            .TimerAutoClose.Enabled = True
            
        End If 'Close respective IF..THEN block statement
        
        For vIndex(&H0) = &H0 To .ShpBttnButton.UBound Step &H1
            .ShpBttnButton(vIndex(&H0)).TabIndex = vIndex(&H0) + (VBA.Val((.ChkDontShowMsg.Visible)) * VBA.Val((.ChkDontShowMsg.Visible)))
        Next vIndex(&H0)
        
        .Width = .Fra_Details.Left + .Fra_Details.Width + &HC8
        .Height = .ImgFooter.Top + .ImgFooter.Height + 350
        .ImgHeader.Width = .Width: .ImgFooter.Width = .Width
        
        vSelectedButton(&H0) = -&H1 'Initialize selected button
        
        'Assign vMsgBox Default Button
        '---------------------------------------------------------------
        .Tag = VBA.IIf(vUseDefault, VBA.IIf(MyDefButton < &H0, MyDefButton, nButtons - ((256 * nButtons) / 256)), -&H1)
        '---------------------------------------------------------------
        
        Set .ParentFrm = iFrm
        
        On Local Error Resume Next
        If Nothing Is iFrm Then CenterForm Frm_MsgBox, Nothing Else CenterForm Frm_MsgBox, iFrm
        
        VBA.DoEvents
        vMsgBox = vSelectedButton(&H0)
        
        'Indicate that a process or operation is complete.
        Screen.MousePointer = vbDefault
        
    End With
    
End Function

Public Sub AltLvBackground(vLv As ListView, ByVal BackColorOne As OLE_COLOR, ByVal BackColorTwo As OLE_COLOR, Optional StartAtOddRow As Boolean = False)
On Local Error GoTo Handle_AltLvBackground_Error
    
    '---------------------------------------------------------------------------------
    ' Purpose   : Alternates row colors in a ListView control
    ' Method    : Creates a picture box and draws the desired color scheme in it, then
    '             loads the drawn image as the listviews picture.
    '---------------------------------------------------------------------------------
    
    Dim lH      As Long
    Dim lSM     As Byte
    Dim picAlt  As PictureBox
    
    BackColorOne = VBA.IIf(BackColorOne = &H0, 16777215, BackColorOne)
    BackColorTwo = VBA.IIf(BackColorTwo = &H0, &HFDE6E8, BackColorTwo)
    
    With vLv
        
        If .View = lvwReport And .ListItems.Count Then
            
            Set picAlt = vLv.Parent.Controls.Add("VB.PictureBox", "picAlt")
            
            lSM = .Parent.ScaleMode
            .Parent.ScaleMode = vbTwips
            .PictureAlignment = lvwTile
            lH = .ListItems(&H1).Height
            
            With picAlt
                
                .BackColor = VBA.IIf(StartAtOddRow = False, BackColorOne, BackColorTwo)
                .AutoRedraw = True
                .Height = lH * &H2
                .BorderStyle = &H0
                .Width = 10 * Screen.TwipsPerPixelX
                picAlt.Line (0, lH)-(.ScaleWidth, lH * &H2), VBA.IIf(StartAtOddRow = False, BackColorTwo, BackColorOne), BF
                
                Set vLv.Picture = .Image
                
            End With
            
            Set picAlt = Nothing
            
            vLv.Parent.Controls.Remove "picAlt"
            vLv.Parent.ScaleMode = lSM
            
        End If
        
    End With
    
Exit_AltLvBackground:
    
Handle_AltLvBackground_Error:
    
End Sub

Public Function AttachPhoto(Frm As Form, PhotoHolder As Image, Optional Dlg As CommonDialog, Optional FileExt$ = "All Supported Files|*.jpeg;*.jpg;*.gif;*.bmp|Jpeg Files(*.jpeg)|*.jpeg|Jpg Files(*.jpg)|*.jpg" & "|Gif Files(*.gif)|*.gif|Bitmap Files(*.bmp)|*.bmp", Optional ImgHeight& = 400, Optional ImgWidth& = 400, Optional PhotoIndex& = &H0) As Boolean
On Local Error GoTo Handle_AttachPhoto_Error
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'If the CommonDialog control has not been specified then Assign the one on the target Form
    If Nothing Is Dlg Then Set Dlg = Frm.Dlg
    
AttachPhoto:
    
    'Execute a series of statements on the specified object
    With Dlg
        
        'Set Dialog to only show Picture Files
        .Filter = FileExt
        
        .FLAGS = &H4 'Hide Read-Only checkbox
        
        'Generate error when the user chooses the Cancel button.
        .CancelError = True
        
        'Set Mouse pointer to indicate end of this process or operation to finish.
        Screen.MousePointer = vbDefault
        
        .ShowOpen 'Display the CommonDialog control's Open dialog box.
        
        'Set Mouse pointer to indicate beginning of process or operation
        Screen.MousePointer = vbHourglass
        
        'If a Picture with a valid path has been Selected
        If VBA.LenB(VBA.Trim$(.FileName)) <> &H0 Then
            
            'Remove the initial Picture Displayed in the Photo holder
            PhotoHolder.Picture = Nothing
            
            'Display the Picture in the Photo holder
            PhotoHolder.Picture = VB.LoadPicture(.FileName)
            
            'Images with large dimensions take long to load therefore..
            'if the image dimension is greater than 400 then...
            If (ImgHeight <> &H0 And ImgWidth <> &H0) And (VBA.Int(PhotoHolder.Height / Screen.TwipsPerPixelY) > ImgHeight Or VBA.Int(PhotoHolder.Width / Screen.TwipsPerPixelX) > ImgWidth) Then
                
                'Set Mouse pointer to indicate end of this process or operation to finish.
                Screen.MousePointer = vbDefault
                
                'Warn User
                vMsgBox "The selected image has dimensions {" & VBA.Int(PhotoHolder.Width / Screen.TwipsPerPixelX) & " x " & VBA.Int(PhotoHolder.Height / Screen.TwipsPerPixelY) & "}. Please select an image with dimensions not greater than {400 x 400}" & VBA.vbCrLf & "{Images with greater dimensions take long to load}", vbExclamation, App.Title & " : Large Image Size", Frm
                
                PhotoHolder.Picture = Nothing
                
                GoTo AttachPhoto 'Branch to the specified Label
                
            End If 'End respective IF..THEN block statement
            
            'Assign byte length of the picture to vDataBytes variable
            ReDim sAdditionalPhoto(PhotoIndex).vDataBytes(VBA.FileLen(.FileName))
            
            'Enable Input/Output to the Selected picture file
            Open .FileName For Binary As #1
                
                'Read data from the image file into vDataBytes variable
                Get #1, , sAdditionalPhoto(PhotoIndex).vDataBytes
                
            Close #1 'Conclude Input/Ouput to the opened file
            
        End If 'End respective IF..THEN block statement
        
    End With 'End the WITH statement
    
    AttachPhoto = True 'Denote that the Image was successfully attached
    
Exit_AttachPhoto:
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Function 'Quit this Function
    
Handle_AttachPhoto_Error:
    
    'If it is a Cancel error then resume execution at the specified Label
    If Err.Number = 32755 Then Resume Exit_AttachPhoto
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error Attaching Photo - " & Err.Number, Frm
    
    'Resume execution at the specified Label
    Resume Exit_AttachPhoto
    
End Function

Public Function AutoComplete(Obj As TextBox, auTblName$, auFldName$, Optional BackSpacePressed As Boolean = False) As String
On Local Error GoTo Handle_AutoComplete_Error
    
    'If no entry has been made then resume execution at the specified Label
    If VBA.LenB(VBA.Trim$(Obj.Text)) = &H0 Then GoTo Exit_AutoComplete
    
    Dim SelStart&
    Dim MousePointerState%
    Dim tRs As New ADODB.Recordset
    
    SelStart = Obj.SelStart
    If BackSpacePressed Then Obj.LinkItem = VBA.Left$(Obj.LinkItem, VBA.Len(Obj.LinkItem) - &H1): SelStart = SelStart - &H1
    
    If Obj.LinkItem = Obj.Text Then If Not BackSpacePressed Then GoTo Exit_AutoComplete Else SelStart = SelStart + &H1
    
    Obj.LinkItem = Obj.Text
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Create Auto-complete feature for Employee positions
    If VBA.LenB(VBA.Trim$(vAdoCNN.ConnectionString)) = &H0 Then Call ConnectDB(, False)
    If vAdoCNN.State = adStateClosed Then vAdoCNN.Open
    
    Set tRs = New ADODB.Recordset
    tRs.Open "SELECT DISTINCT [" & auFldName & "] FROM " & auTblName & VBA.IIf(VBA.InStr(VBA.LCase(auTblName), VBA.LCase(" WHERE ")) <> &H0, " AND", " WHERE") & " [" & auFldName & "] LIKE '" & VBA.Replace(Obj.LinkItem, "'", "''") & "%' ORDER BY [" & auFldName & "] ASC", vAdoCNN, adOpenKeyset, adLockReadOnly
    
    'If matching values exist then...
    If Not tRs.EOF Then
        
        'Auto-complete and..
        Obj.Text = tRs(auFldName)
        
        '..highlight the remaining text
        Obj.SelStart = SelStart: Obj.SelLength = VBA.Len(Obj.Text) - SelStart
        
    End If 'Close respective IF..THEN block statement
    
    Obj.LinkItem = VBA.Left$(Obj.Text, Obj.SelStart)
    tRs.Close: Set tRs = Nothing
    
Exit_AutoComplete:
    
    AutoComplete = Obj.Text
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    Exit Function 'Quit this Procedure
    
Handle_AutoComplete_Error:
    
    If Err.Number = &H5 Then Resume Exit_AutoComplete
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Auto-complete Error - " & Err.Number, Obj.Parent
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Resume execution at the specified Label
    Resume Exit_AutoComplete
    
End Function

Public Function BackupData(iFrm As Form) As Boolean
On Local Error GoTo Handle_BackupData_Error
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    
    
    BackupData = True
    
Exit_BackupData:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    Exit Function 'Quit this Procedure
    
Handle_BackupData_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error Backing up Data - " & Err.Number, iFrm
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Resume execution at the specified Label
    Resume Exit_BackupData
    
End Function

'This Function was created to replace VB's VBA.StrConv(x, vbProperCase) since it doesn't take care of certain
'characters that require capitalizing the next character to them
Public Function CapAllWords(sString$, Optional sDelimiter = " ") As String
On Local Error GoTo Handle_CapAllWords_Error
    
    'If no string has been defined then quit this Function
    If VBA.LenB(VBA.Trim$(sString)) = &H0 Then GoTo Exit_CapAllWords
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    Dim cCnt1&, cCnt2&
    Dim cArray() As String
    Dim cArray2() As String
    Dim cExemptions, cExemptionsSmall$, cNewWord$, cChar$
    
    'Words that should not be altered
    cExemptions = "|II|III|IV|VI|VII|VIII|IX|"
    cExemptionsSmall = "|and|of|for|a|an|in|e.t.c|i.e.|"
    
    'Split according to the specified delimiter
    cArray = VBA.Split(sString, sDelimiter)
    
    'For each set of words...
    For cCnt1 = &H0 To UBound(cArray) Step &H1
        
        'If the current set is not among the exempted ones then...
        If (VBA.InStr(VBA.LCase$(cExemptionsSmall), "|" & VBA.LCase$(cArray(cCnt1)) & "|") = &H0) And (VBA.InStr(VBA.LCase$(cExemptions), "|" & VBA.LCase$(cArray(cCnt1)) & "|") = &H0) Then
            
            'For each character in the set...
            For cCnt2 = &H1 To VBA.Len(cArray(cCnt1)) Step &H1
                
                'If it is the first character in the set, Capitalize it, else change it to lowercase
                cChar = VBA.IIf(cCnt2 = &H1, VBA.UCase$(VBA.Mid$(cArray(cCnt1), cCnt2, &H1)), VBA.LCase$(VBA.Mid$(cArray(cCnt1), cCnt2, &H1)))
                
                If cCnt2 > &H1 Then
                    
                    'Also capitalize the first characters after the following...
                    
                    '10 & 13    => Enter/New Line
                    '32         => Space
                    '40      (  => Opening Brackets
                    '44      ,  => Comma
                    '45      -  => Hyphen/Dash
                    '46      .  => Fullstop
                    '47      /  => Fullstop
                    
                    Select Case VBA.Asc(VBA.Mid$(cArray(cCnt1), cCnt2 - &H1, &H1))
                        Case 10, 13, 32, 40, 44, 45, 46, 47: cChar = VBA.UCase$(VBA.Mid$(cArray(cCnt1), cCnt2, &H1))
                    End Select 'End SELECT..CASE block statement
                    
                End If 'End respective IF..THEN block statement
                
                cNewWord = cNewWord & cChar
                
            Next cCnt2 'Move to the next character
            
        Else 'If the current set is among the exempted ones then...
            
            'Get the set as it is
            cNewWord = cNewWord & VBA.IIf(VBA.InStr(VBA.LCase$(cExemptions), "|" & VBA.LCase$(cArray(cCnt1)) & "|") = &H0, VBA.IIf(cCnt1 = &H0, VBA.StrConv(cArray(cCnt1), vbProperCase), VBA.LCase(cArray(cCnt1))), cArray(cCnt1))
            
        End If 'End respective IF..THEN block statement
        
        'Generate a full formatted string as the process continues
        CapAllWords = CapAllWords & cNewWord & sDelimiter
        
        cNewWord = VBA.vbNullString 'Initialize variable
        
    Next cCnt1 'Move to the next set of word
    
    'If a delimiter is at the end of the formatted string, remove it
    If VBA.Right$(CapAllWords, &H1) = sDelimiter Then CapAllWords = VBA.Left$(CapAllWords, VBA.Len(CapAllWords) - &H1)
    
Exit_CapAllWords:
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Function 'Quit this Function
    
Handle_CapAllWords_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Format Entry Error - " & Err.Number
    
    'Resume execution at the specified Label
    Resume Exit_CapAllWords
    
End Function

Public Function CenterForm(Frm As Form, Optional vParent As Object, Optional vModal As Boolean = True) As Boolean
On Local Error GoTo Handle_CenterForm_Error
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    If Nothing Is vParent Then
        
        Frm.Left = (Screen.Width / &H2) - (Frm.Width / &H2)
        Frm.Top = (Screen.Height / &H2) - (Frm.Height / &H2)
        
        Frm.Height = Frm.Height + &H1: Frm.Height = Frm.Height - &H1
        
        'Change Mouse Pointer to show end of Processing state
        Screen.MousePointer = vbDefault
        
        If vModal Then Frm.Show vbModal Else Frm.Show
        
    ElseIf TypeOf vParent Is Form Then
        
        Frm.Left = (vParent.Left + (vParent.Width / &H2)) - (Frm.Width / &H2)
        Frm.Top = (vParent.Top + (vParent.Height / &H2)) - (Frm.Height / &H2)
        
        'Change Mouse Pointer to show end of Processing state
        Screen.MousePointer = vbDefault
        
        If vModal Then Frm.Show vbModal, vParent Else Frm.Show , vParent
        
    Else
        
        Dim oObj As Object
        Dim iLeft&, iTop&
        
        Set oObj = vParent
        
        Do While Not TypeOf oObj Is Form
            iLeft = iLeft + oObj.Left: iTop = iTop + oObj.Top
            Set oObj = oObj.Container
        Loop
        
        iLeft = iLeft + oObj.Left: iTop = iTop + oObj.Top
        
        Frm.Left = iLeft + ((vParent.Width / &H2) - (Frm.Width / &H2)) + 50
        Frm.Top = iTop + ((vParent.Height / &H2) - (Frm.Height / &H2)) + 400
        
        Frm.Height = Frm.Height + &H1: Frm.Height = Frm.Height - &H1
        
        'Change Mouse Pointer to show end of Processing state
        Screen.MousePointer = vbDefault
        
        If vModal Then Frm.Show vbModal Else Frm.Show
        
    End If 'End respective IF..THEN block statement
    
Exit_CenterForm:
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Function 'Quit this Function
    
Handle_CenterForm_Error:
    
    If Err.Number = 364 Or Err.Number = 400 Or Err.Number = 401 Then Resume Exit_CenterForm
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error Centering Form - " & Err.Number, Frm
    
    'Resume execution at the specified Label
    Resume Exit_CenterForm
    
End Function

Public Function CheckDBPwdValidity(strDBName$, Pwd$) As Boolean
On Local Error GoTo Handle_CheckDBPwdValidity_Error
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    Dim dbCNN As New ADODB.Connection
    
    'Supply password and open. If it opens then the password is correct else wrong
    dbCNN.Open "Provider=" & vAdoCNN.Provider & ";Data Source=" & strDBName & ";Persist Security Info=False;Jet OLEDB:Database Password=" & Pwd
    dbCNN.Close 'Close thedatabase
    Set dbCNN = Nothing 'disassociate the database variable from actual database
    
    CheckDBPwdValidity = True 'Denote that the password is correct
    
Exit_CheckDBPwdValidity:
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Function 'Quit this Function
    
Handle_CheckDBPwdValidity_Error:
    
    'Resume execution at the specified Label
    Resume Exit_CheckDBPwdValidity
    
End Function

Public Function CheckForRecordDependants(MyFrm As Form, iTableName$, Optional iPerformCleanUp As Boolean = False) As Boolean
On Local Error GoTo Handle_CheckForRecordDependants_Error
    
    Dim MousePointerState%
    Dim nRs As New ADODB.Recordset
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    If VBA.LenB(VBA.Trim$(vAdoCNN.ConnectionString)) = &H0 Then Call ConnectDB 'Call Function in this Module to connect to Software's Database
    Set nRs = New ADODB.Recordset
    With nRs 'Execute a series of statements on vRs recordset
        
        .Open "SELECT * FROM " & iTableName, vAdoCNN, adOpenKeyset, adLockReadOnly
        CheckForRecordDependants = Not (.BOF And .EOF)
        .Close 'Close the opened object and any dependent objects
        
    End With 'Close the WITH block statements
    
Exit_CheckForRecordDependants:
    
    'Call Procedure in this Module to free memory and system resources
    If iPerformCleanUp Then Call PerformMemoryCleanup
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Function 'Quit this Function
    
Handle_CheckForRecordDependants_Error:
    
    'If the specified table does not exist then resume execution at the specified Label
    If Err.Number = -2147217865 Then Resume Exit_CheckForRecordDependants
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error Checking Dependency - " & Err.Number
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Resume execution at the specified Label
    Resume Exit_CheckForRecordDependants
    
End Function

Public Function CloseFrm(MyFrm As Object, Optional AskUser As Boolean = True, Optional InputForm As Boolean = True, Optional CustomMsg$ = "") As Boolean
On Local Error GoTo Handle_CloseFrm_Error
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    CloseFrm = False 'Prevent the Form from Terminating
    
    'If the Form's closure should be questioned then...
    If AskUser = True And Not vSilentClosure And Not Nothing Is MyFrm Then
        
        'Indicate that a process or operation is complete.
        Screen.MousePointer = vbDefault
        
        'Confirm if User wants to Close the Form, if No then quit this procedure
        If vMsgBox(VBA.IIf(CustomMsg <> "", CustomMsg, "Are you sure you want to Close this " & VBA.IIf(TypeOf MyFrm Is Form, "Form", VBA.IIf(TypeOf MyFrm Is DataReport, "Report", "Object"))) & "?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title & " : Confirmation", MyFrm) = vbNo Then Exit Function
        
        'Indicate that a process or operation is in progress.
        Screen.MousePointer = vbHourglass
        
    End If 'Close respective IF..THEN block statement
    
    'If a record is being modified and the user chooses to proceed with modifying it then Quit this Procedure
    If Not Nothing Is MyFrm Then If TypeOf MyFrm Is Form And Not vSilentClosure Then If ProceedEditting(MyFrm, InputForm) = True Then GoTo Exit_CloseFrm
    
    CloseFrm = True 'Allow the Form to Terminate
    
Exit_CloseFrm:
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Function 'Quit this Function
    
Handle_CloseFrm_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error Closing " & VBA.IIf(TypeOf MyFrm Is Form, "Form", VBA.IIf(TypeOf MyFrm Is DataReport, "Report", "Object")) & " - " & Err.Number, MyFrm
    
    'Resume execution at the specified Label
    Resume Exit_CloseFrm
    
End Function

Public Function ConnectDB(Optional ResetAdoCNN As Boolean = True, Optional ResetRS As Boolean = True, Optional ResetDEnv As Boolean = False) As Boolean
On Local Error GoTo Handle_ConnectDB_Error
    
    Static dbPwd$
    Static DbPath$
    
    Dim MousePointerState%
    Dim DEnvConnectionError, vOffice2003 As Boolean
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
ConnectToDB:
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'If vAdoCNN has not been opened then...
    If vAdoCNN.State = adStateClosed Or ResetAdoCNN Then
        
        'Create a new instance of database connector
        Set vAdoCNN = New ADODB.Connection
        
        vOffice2003 = vFso.FileExists(App.Path & "\" & App.Title & " Database Connector.udl")
        
        'vAdoCNN.ConnectionString = "Driver={MySQL ODBC 5.1 Driver};Server=localhost;Port=3306;Database=xprosdb;system=root;Password=p@ssword;Option=3;"
        DbPath = App.Path & "\" & App.Title & " Database Connector.udl"
        
        If vOffice2003 Then
            
            vAdoCNN.Provider = "Microsoft.Jet.OLEDB.4.0"
            
            'Specify connection path to Universal Data Link file which has been linked to the Software Database
            vAdoCNN.ConnectionString = "FILE NAME=" & DbPath
            
            vAdoCNN.CursorLocation = adUseClient
            
            'Assign currently set database password
            vAdoCNN.Properties("Jet OLEDB:Database Password") = dbPwd
            
            vAdoCNN.Open 'Open Connection
            
        Else 'For Office 2007 & later
            
            vAdoCNN.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & App.Path & "\" & App.Title & " Database.accdb;Persist Security Info=False;Jet OLEDB:Database Password=" & dbPwd
            
        End If
        
ConnectToDENV:
        
'        If Not DEnvConnectionError And ResetDEnv Then
'
'            With DEnv_Xpros.RptCNN
'
'                If .State = adStateOpen Then .Close 'Close Report Connection
'
'                'Set the Report Connector to the valid database path connected by vAdoCNN
'                .ConnectionString = vAdoCNN.ConnectionString
'
'                'Assign the database password to it
'                '.Properties("Password") = vAdoCNN.Properties("Jet OLEDB:Database Password")
'
'            End With 'Close WITH block statement
'
'        End If 'Close respective IF..THEN block statement
        
        'If vAdoCNN.State = adStateOpen Then Call GetSoftwareSettings(False)
        
    End If 'Close respective IF..THEN block statement
    
    'If Recordsets have to be reset then...
    If ResetRS Then
        
        'Reset Recordsets
        Set vRs = New ADODB.Recordset
        Set vRsTmp = New ADODB.Recordset
        
    End If 'Close respective IF..THEN block statement
    
Exit_ConnectDB:
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Function 'Quit this Function
    
Handle_ConnectDB_Error: 'In case an error occurs while connecting to the Database then...
    
    Dim iErr(&H1) As String
    
    iErr(&H0) = Err.Number
    iErr(&H1) = Err.Description
    
    'If vAdoCNN has not been set then branch to OpenvAdoCNN Label
    If iErr(&H0) = 91 Then Resume Next
    
    'If connection error then...
    If iErr(&H0) = 432 Then
        
        'If Db Path does not exist then...
        If Not vFso.FileExists(DbPath) Then
            
            'Indicate that a process or operation is complete.
            Screen.MousePointer = vbDefault
            
            'Warn User about the problem
            vMsgBox "The Software has failed to connect to the Server since it does not exist at the specified location. Terminating in £ seconds..", vbCritical, App.Title & " : Server Connection Error", , , , &HA
            
            'Quit this Application
            GoTo QuitApp
            
        End If 'Close respective IF..THEN block statement
        
    End If 'Close respective IF..THEN block statement
    
    'Automation error for Reports
    If iErr(&H0) = -2147024770 Then DEnvConnectionError = True: Resume ConnectToDENV
    
    'If the software does not have the correct database password then...
    If iErr(&H0) = -2147217843 Then
        
        'Check the saved Password in the registry if it is the correct one
        dbPwd = SmartDecrypt(VBA.GetSetting(App.Title, "Settings", "DB Password", VBA.vbNullString))
        
        'If the supplied database password is wrong then...Proceed with the database connection process
        If CheckDBPwdValidity(vAdoCNN.Properties("Data Source"), dbPwd) Then GoTo ConnectToDB
        
        'Warn User and confirm Password verification
        If vMsgBox("The Software's database is password protected. Please enter the Password for the Software to function normally. Proceed?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title & " : Confirmation") = vbYes Then
            
            Dim Trials%
            
RequestPwd:
            
            Load Frm_DataEntry
            Frm_DataEntry.IsPassword = True
            Frm_DataEntry.LblInput.Caption = "Database Password"
            
            'Indicate that a process or operation is complete.
            Screen.MousePointer = vbDefault
            
            Frm_DataEntry.Show vbModal
            
            'Indicate that a process or operation is in progress.
            Screen.MousePointer = vbHourglass
            
            'Request User to supply the Database's New Password.
            dbPwd = vBuffer(&H0): Erase vBuffer 'Initialize variable
            
            'If the supplied database password is wrong then...
            If Not CheckDBPwdValidity(vAdoCNN.Properties("Data Source"), dbPwd) Then
                
                Trials = Trials + &H1 'Increment trials counter by 1
                
                'Indicate that a process or operation is complete.
                Screen.MousePointer = vbDefault
                
                'If the specified max no of incorrect entries has not been reached then
                If Trials < &H3 Then
                    
                    'Warn User and request for password re-entry
                    If vMsgBox("The entered Password is incorrect. Retry?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title & " : Invalid Db Password") = vbYes Then GoTo RequestPwd
                    
                Else 'If the specified max no of incorrect entries has been reached then
                    
                    'Warn User
                    vMsgBox "The entered Password is incorrect. Terminating...", vbCritical, App.Title & " : Invalid Db Password"
                    
                    'Quit this Application
                    GoTo QuitApp
                    
                End If 'Close respective IF..THEN block statement
                
            Else 'If the supplied database password is correct then...
                
                'Encrypt and store the new password in the System's Registry
                VBA.SaveSetting App.Title, "Settings", "DB Password", SmartEncrypt(dbPwd)
                
                Resume ConnectToDB 'Proceed with the database connection process
                
            End If 'Close respective IF..THEN block statement
            
        End If 'Close respective IF..THEN block statement
        
    End If 'Close respective IF..THEN block statement
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox "A Connection Local Error has been encountered {" & iErr(&H1) & "}. Terminating in £ seconds...", vbExclamation, App.Title & " : Connection Error - " & iErr(&H0), , , , &HA
    
QuitApp:
    
    'Call Procedure in this Module to free memory and system resources
    Call PerformMemoryCleanup
    
    End 'Halt the Application
    
End Function

'Builds specified path by creating Folders that don't exist
Public Function CreatePath(ByVal vPath$) As String
On Local Error GoTo Handle_CreatePath_Error
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    Dim iCnt&
    Dim iDrive
    Dim iPathFolders() As String
    Dim iFso As New FileSystemObject
    
    If VBA.Trim$(vPath) = VBA.vbNullString Then Exit Function
    
    'Get the Drive to Create the folders into
    Set iDrive = iFso.GetDrive(iFso.GetDriveName(vPath))
    
    iPathFolders = VBA.Split(VBA.Replace(vPath, iDrive & "\", VBA.vbNullString), "\")
    CreatePath = iDrive
    
    For iCnt = &H0 To UBound(iPathFolders) Step &H1
        
        If Not iFso.FolderExists(CreatePath & "\" & iPathFolders(iCnt)) Then iFso.CreateFolder CreatePath & "\" & iPathFolders(iCnt)
        CreatePath = CreatePath & "\" & iPathFolders(iCnt)
        
    Next iCnt
    
Exit_CreatePath:
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Function 'Quit this Function
    
Handle_CreatePath_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error Building Folder Path - " & Err.Number
    
    'Resume execution at the specified Label
    Resume Exit_CreatePath
    
End Function

Public Function FindLvItemByFirstItemChar(vLv As ListView, KeyAscii As Integer, Optional ColumnPos& = &H1, Optional iShiftKey% = &H0) As ListItem
On Local Error GoTo Handle_FindLvItemByFirstItemChar_Error
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    Dim Lst
    Static StartPos&
    
    'Ensure the column position entered is valid
    ColumnPos = VBA.IIf(ColumnPos < &H1, &H1, ColumnPos)
    
    If Not Nothing Is vLv.SelectedItem Then StartPos = vLv.SelectedItem.Index: StartPos = StartPos + VBA.IIf(iShiftKey, &H0, &H1)
    StartPos = VBA.IIf(StartPos = &H0 Or StartPos > vLv.ListItems.Count, &H1, StartPos)
    
    vIndex(&H1) = &H0 'Initialize variable
    
    'If the Shift Key has not been pressed then...
    If iShiftKey = &H0 Then
        
        'Start Search from last search's Position to the last item
        For vIndex(&H0) = StartPos To vLv.ListItems.Count Step &H1
            
            If ColumnPos = &H1 Then Set Lst = vLv.ListItems(vIndex(&H0)) Else Set Lst = vLv.ListItems(vIndex(&H0)).ListSubItems(ColumnPos - &H1)
            If Not Nothing Is Lst Then If Lst.Text <> VBA.vbNullString Then If VBA.Asc(VBA.LCase$(VBA.Left$(Lst.Text, &H1))) = VBA.Asc(VBA.LCase$(VBA.Chr$(KeyAscii))) Then vIndex(&H1) = vIndex(&H0): Exit For
            
        Next vIndex(&H0) 'Move to the next item
        
    Else 'If the Shift Key has been pressed then...
        
        'Start Search from last search's Position to the last item
        For vIndex(&H0) = StartPos - &H1 To &H1 Step -&H1
            
            If ColumnPos = &H1 Then Set Lst = vLv.ListItems(vIndex(&H0)) Else Set Lst = vLv.ListItems(vIndex(&H0)).ListSubItems(ColumnPos - &H1)
            If Not Nothing Is Lst Then If Lst.Text <> VBA.vbNullString Then If VBA.Asc(VBA.LCase$(VBA.Left$(Lst.Text, &H1))) = VBA.Asc(VBA.LCase$(VBA.Chr$(KeyAscii))) Then vIndex(&H1) = vIndex(&H0): Exit For
            
        Next vIndex(&H0) 'Move to the next item
        
    End If 'Close respective IF..THEN block statement
    
    'If a matching item has been found the select it, make it visible to the User and then Quit this Procedure
    If vIndex(&H1) > &H0 Then Set FindLvItemByFirstItemChar = vLv.ListItems(vIndex(&H1)): StartPos = vIndex(&H1): vLv.ListItems(vIndex(&H1)).Selected = True: vLv.ListItems(vIndex(&H1)).EnsureVisible: KeyAscii = Empty: GoTo Exit_FindLvItemByFirstItemChar
    
    'If the last Search's position was not at the first item then...
    If StartPos <> &H1 Then
        
        'If the Shift Key has not been pressed then...
        If iShiftKey = &H0 Then
            
            'Start Search from the first item to the last search's Position
            For vIndex(&H0) = &H1 To StartPos - &H1 Step &H1
                
                If ColumnPos = &H1 Then Set Lst = vLv.ListItems(vIndex(&H0)) Else Set Lst = vLv.ListItems(vIndex(&H0)).ListSubItems(ColumnPos - &H1)
                If Not Nothing Is Lst Then If Lst.Text <> VBA.vbNullString Then If VBA.Asc(VBA.LCase$(VBA.Left$(Lst.Text, &H1))) = VBA.Asc(VBA.LCase$(VBA.Chr$(KeyAscii))) Then vIndex(&H1) = vIndex(&H0): Exit For
                
            Next vIndex(&H0) 'Move to the next item
            
        Else 'If the Shift Key has been pressed then...
            
            'Start Search from the Last item to the last search's Position
            For vIndex(&H0) = vLv.ListItems.Count To StartPos + &H1 Step -&H1
                
                If ColumnPos = &H1 Then Set Lst = vLv.ListItems(vIndex(&H0)) Else Set Lst = vLv.ListItems(vIndex(&H0)).ListSubItems(ColumnPos - &H1)
                If Not Nothing Is Lst Then If Lst.Text <> VBA.vbNullString Then If VBA.Asc(VBA.LCase$(VBA.Left$(Lst.Text, &H1))) = VBA.Asc(VBA.LCase$(VBA.Chr$(KeyAscii))) Then vIndex(&H1) = vIndex(&H0): Exit For
                
            Next vIndex(&H0) 'Move to the next item
            
        End If 'Close respective IF..THEN block statement
        
        'If a matching item has been found the select it, make it visible to the User and then Quit this Procedure
        If vIndex(&H1) > &H0 Then Set FindLvItemByFirstItemChar = vLv.ListItems(vIndex(&H1)): StartPos = vIndex(&H1): vLv.ListItems(vIndex(&H1)).Selected = True: vLv.ListItems(vIndex(&H1)).EnsureVisible: KeyAscii = Empty: GoTo Exit_FindLvItemByFirstItemChar
        
    End If 'Close respective IF..THEN block statement
    
Exit_FindLvItemByFirstItemChar:
     
    iShiftKey = &H0 'Initialize variable
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Function 'Quit this Procedure
    
Handle_FindLvItemByFirstItemChar_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error Finding Lv Item By First Item Char - " & Err.Number, vLv.Parent
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Resume execution at the specified Label
    Resume Exit_FindLvItemByFirstItemChar
    
End Function

Public Function EventLog(Optional iFrm$, Optional iType& = &H40, Optional iSource$, Optional iNumber&, Optional iDesc$, Optional iSelectedOption$ = "OK") As Boolean
On Local Error GoTo Handle_EventLog_Error
    
    If User.User_Name = VBA.vbNullString Then Exit Function
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    Dim Fl
    Dim LogFilePath$, LogFileData$
    
    If VBA.LenB(def_LogFileLocation) = &H0 Then def_LogFileLocation = App.Path & "\Application Data"

    LogFilePath = def_LogFileLocation & "\" & App.Title & " Event Log.txt"
    
    'If the Log File Folder does not exist then create it
    If Not vFso.FolderExists(vFso.GetParentFolderName(LogFilePath)) Then Call CreatePath(vFso.GetParentFolderName(LogFilePath))
    
    Set Fl = vFso.OpenTextFile(LogFilePath, ForReading, True)
    If Not Fl.AtEndOfStream Then LogFileData = Fl.ReadAll
    Fl.Close
    
    Dim iMsgType$
    Dim MaxAppLogs&
    Dim iArray() As String
    
    MaxAppLogs = SoftwareSetting.Max_App_Logs
    
    iArray = VBA.Split(LogFileData, VBA.vbCrLf)
    
    If UBound(iArray) > MaxAppLogs - &H1 Then
        
        ReDim Preserve iArray(MaxAppLogs - &H2) As String
        LogFileData = VBA.Join(iArray, VBA.vbCrLf)
        
    End If
    
    'Check the type of message being displayed
    Select Case iType
        
        Case vbCritical: iMsgType = "Critical"
        Case vbExclamation: iMsgType = "Warning"
        Case vbQuestion: iMsgType = "Question"
        Case Else: iMsgType = "Information"
        
    End Select 'Close SELECT..CASE block statement
    
    LogFileData = VBA.Replace(VBA.Replace(VBA.Format$(VBA.Date, "dd/mm/yyyy") & "|" & VBA.Format$(VBA.Time, "HH:nn:ss") & "|" & User.Device_Serial_No & "|" & User.Device_Name & "|" & User.Device_Account_Name & "|" & "|" & User.User_ID & "|" & User.User_Name & "|" & iFrm & "|" & iSource & "|" & iNumber & "|" & iMsgType & "|" & iDesc, VBA.vbCrLf & VBA.vbCrLf, VBA.vbCrLf), VBA.vbCrLf, ". ") & "|" & iSelectedOption & VBA.IIf(LogFileData <> VBA.vbNullString, VBA.vbCrLf, VBA.vbNullString) & LogFileData
    
    Set Fl = vFso.OpenTextFile(LogFilePath, ForWriting, True)
    Fl.Write LogFileData
    Fl.Close
    
    EventLog = True
    
Exit_EventLog:
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Function 'Quit this Function
    
Handle_EventLog_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    MsgBox Err.Description, vbExclamation, App.Title & " : Error Writing Log - " & Err.Number
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Resume execution at the specified Label
    Resume Exit_EventLog
    
End Function

Public Function ExportLvToExcel(xLv As ListView, Optional xFileName As String, Optional xFreezeCol& = &H1, Optional xFontSize% = &H8, Optional xLandScape As Boolean = False) As Boolean
On Local Error GoTo Handle_ExportLvToExcel_Error
    
    'If the specified Listview has no data for transfer then...
    If xLv.ListItems.Count = &H0 Then
        
        'Warn User
        vMsgBox "There are no items to be exported to Excel", vbExclamation, App.Title & " : Operation Aborted", xLv.Parent
        Exit Function 'Quit this Function
        
    End If 'Close respective IF..THEN block statement
    
    'Confirm if the User really wants to Export the Data to Excel. If not then Quit this Function
    If vMsgBox("Are you sure you want to transfer the data to Excel?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title & " : Confirmation", xLv.Parent) = vbNo Then Exit Function
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    Dim xlsAborted As Boolean
    Dim xRow&, xCol&, xMaxCols&, xStartRow&
    Dim xQry$, xStatement$, xValidExportCols$
    
    Dim MsExcelApp As Object
    Dim MsExcelSheet As Object
    
    Set MsExcelApp = CreateObject("Excel.Application")
    
    MsExcelApp.Workbooks.Add
    
    'If the Excel version is 8.0 or greater then...
    If Val(MsExcelApp.Application.Version) >= &H8 Then
        Set MsExcelSheet = MsExcelApp.ActiveSheet
    Else 'If the Excel version is less than 8.0 then...
        Set MsExcelSheet = MsExcelApp
    End If 'Close respective IF..THEN block statement
    
    '-------------------------------------------------------------------------------------------
    '           CALCULATE THE MAXIMUM VALUE FOR THE 'PLEASE WAIT' PROGRESS BAR
    '-------------------------------------------------------------------------------------------
    
    Dim MaxPBarVal&, iCnt#
    
    'For Headers Transfer
    MaxPBarVal = xLv.ColumnHeaders.Count
    
    'For Items Transfer
    MaxPBarVal = MaxPBarVal + (xLv.ColumnHeaders.Count * xLv.ListItems.Count)
    
    Frm_PleaseWait.ImgProgressBar.Width = &H0
    Frm_PleaseWait.lblInfo.Caption = "Transfering columns to Worksheet..."
    
    '-------------------------------------------------------------------------------------------
    
    'Show the 'Please Wait' Form to the User
    CenterForm Frm_PleaseWait, xLv.Parent, False
    
    '--------------------------------------------------------------------------------------------
    '                                   EXPORT LISTVIEW HEADERS
    '--------------------------------------------------------------------------------------------
    
    'Export the heading to Excel
    MsExcelSheet.Cells(&H1, &H1) = VBA.UCase$(School.Name)
    
    xRow = &H2 'Set default row position
    
    'Export the heading to Excel
    MsExcelSheet.Cells(xRow, &H1) = xFileName
    
    xValidExportCols = "|" 'Initialize variable
    
    xRow = xRow + &H1 'Set default row position
    xStartRow = xRow 'Set Starting row position
    
    'For each column in the specified Listview...
    For xCol = &H1 To xLv.ColumnHeaders.Count Step &H1
        
        'If the User has cancelled the process then branch to the specified Label
        If vCancelOperation = &H2 Then xlsAborted = True: GoTo Exit_ExportLvToExcel
        
        'If the User has paused the process then call function to wait
        If vCancelOperation = &H1 Then Call Wait
        
        'If the column header is visible then...
        If (xLv.ColumnHeaders(xCol).Width > 300 And ((Nothing Is xLv.SmallIcons And xLv.Checkboxes) Or (Not Nothing Is xLv.SmallIcons And Not xLv.Checkboxes))) Or (xLv.ColumnHeaders(xCol).Width > 530 And xLv.Checkboxes And Not Nothing Is xLv.SmallIcons) Then
            
            xMaxCols = xMaxCols + &H1
            
            'Export the heading to Excel
            MsExcelSheet.Cells(xRow, xMaxCols) = xLv.ColumnHeaders(xCol).Text
            
            'Apply Borders in each Cell
            MsExcelSheet.Range(MsExcelSheet.Cells(xRow, &H1), MsExcelSheet.Cells(xRow, xMaxCols)).BorderAround &H1, &H2, -4105, &H80&
            
            'Assign the column position
            xValidExportCols = xValidExportCols & xCol & "|"
            
            'Get the DataType of the specified Column
            
            Dim LvColType&
            Dim sArrayDataTypes() As String
            
            sArrayDataTypes = VBA.Split(xLv.Parent.iLvwItemDataType, "|")
            
            Select Case sArrayDataTypes(xCol - &H1)
                
                Case "D": LvColType = 2
                Case "N", "C": LvColType = -4145
                Case Else: LvColType = -4158
                
            End Select 'Close SELECT..CASE block statement
            
        End If 'Close respective IF..THEN block statement
        
        iCnt = iCnt + (Frm_PleaseWait.lblProgressBar.Width / MaxPBarVal)
        Frm_PleaseWait.ImgProgressBar.Width = iCnt
        Frm_PleaseWait.lblStatus.Caption = VBA.Int(((VBA.CLng(iCnt) * 100) / Frm_PleaseWait.lblProgressBar.Width)) & "% Complete"
        
        VBA.DoEvents 'Yield execution so that the operating system can process other events
        
    Next xCol 'Increment counter by 1 to move to the next column in the Lv
    
    '---------------------------------FORMAT EXPORTED HEADINGS----------------------------------
    
    'Merge the Cells of the first row
    MsExcelSheet.Range(MsExcelSheet.Cells(&H1, &H1), MsExcelSheet.Cells(&H1, xMaxCols)).Merge
    
    'Apply Borders in each Cell of the first row
    MsExcelSheet.Range(MsExcelSheet.Cells(&H1, &H1), MsExcelSheet.Cells(&H1, xMaxCols)).BorderAround &H1, &H2, -4105, &H80&
    
    'Merge the Cells of the Second row
    MsExcelSheet.Range(MsExcelSheet.Cells(&H2, &H1), MsExcelSheet.Cells(&H2, xMaxCols)).Merge
    
    'Apply Borders in each Cell of the Second row
    MsExcelSheet.Range(MsExcelSheet.Cells(&H2, &H1), MsExcelSheet.Cells(&H2, xMaxCols)).BorderAround &H1, &H2, -4105, &H80&
    
    'Center Headings in Cells
    MsExcelSheet.Range(MsExcelSheet.Cells(&H1, &H1), MsExcelSheet.Cells(xRow, xMaxCols)).HorizontalAlignment = -4108
    
    'Make the Heading Bold
    MsExcelSheet.Range(MsExcelSheet.Cells(&H1, &H1), MsExcelSheet.Cells(xRow, xMaxCols)).Font.Bold = True
    
    'Set ForeColor to Brown
    MsExcelSheet.Range(MsExcelSheet.Cells(&H1, &H1), MsExcelSheet.Cells(xRow, xMaxCols)).Font.Color = &H80&
    
    '-------------------------------------------------------------------------------------------
    '                                  EXPORT LISTVIEW ITEMS
    '-------------------------------------------------------------------------------------------
    
    Dim Lst
    Dim xlsArray() As String
    Dim vColNo&, vFirstVisibleCol&
    
    Frm_PleaseWait.lblInfo.Caption = "Transfering items to Worksheet..."
    
    xlsArray = VBA.Split(xLv.Parent.iLvwItemDataType, "|")
    
    xRow = xRow + &H1
    
    'For each item in the specified Listview...
    For xRow = xRow To xLv.ListItems.Count + xStartRow Step &H1
        
        'If the User has cancelled the process then branch to the specified Label
        If vCancelOperation = &H2 Then xlsAborted = True: GoTo FormattingWorksheet
        
        'If the User has paused the process then call function to wait
        If vCancelOperation = &H1 Then Call Wait
        
        vColNo = &H0 'Initialize variable
        
        'For each column in the specified Listview...
        For xCol = &H1 To xLv.ColumnHeaders.Count Step &H1
            
            'If the Column value has to be exported then...
            If VBA.InStr(xValidExportCols, "|" & xCol & "|") <> &H0 Then
                
                vColNo = vColNo + &H1
                
                'Get the value of the column item
                If xCol = &H1 Then Set Lst = xLv.ListItems(xRow - xStartRow) Else If xLv.ListItems(xRow - xStartRow).ListSubItems.Count >= xCol - &H1 Then Set Lst = xLv.ListItems(xRow - xStartRow).ListSubItems(xCol - &H1) Else xLv.ListItems(xRow - xStartRow).ListSubItems.Add (xCol - &H1), , VBA.vbNullString: Set Lst = xLv.ListItems(xRow - xStartRow).ListSubItems(xCol - &H1)
                
                vBuffer(&H0) = Lst.Text
                
                'If the current data is currency then...
                If xlsArray(xCol - &H1) = "C" Then
                    
                    'Format the data to 2 decimal places
                    vBuffer(&H0) = VBA.IIf(Lst.Text = VBA.vbNullString, Lst.Text, VBA.FormatNumber(VBA.CLng(VBA.Replace(Lst.Text, ",", VBA.vbNullString)), &H2))
                    
                    'Set the data to 2 decimal places in the Cell (Denote negative values in Red)
                    MsExcelSheet.Range(MsExcelSheet.Cells(xRow, vColNo), MsExcelSheet.Cells(xRow, vColNo)).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
                    
                End If 'Close respective IF..THEN block statement
                
                'Assign the value to the Excel WorkSheet
                MsExcelSheet.Cells(xRow, vColNo) = vBuffer(&H0) & VBA.IIf(VBA.IsNumeric(VBA.Replace(Lst.Text, ",", VBA.vbNullString)) And VBA.LenB(VBA.Int(VBA.Val(VBA.Replace(Lst.Text, ",", VBA.vbNullString)))) > &HB, ";", VBA.vbNullString)
                
                If xRow Mod &H2 = &H0 Then
                    
                    MsExcelSheet.Range(MsExcelSheet.Cells(xRow, vColNo), MsExcelSheet.Cells(xRow, vColNo)).Interior.ColorIndex = 19 'Light Grey
                    MsExcelSheet.Range(MsExcelSheet.Cells(xRow, vColNo), MsExcelSheet.Cells(xRow, vColNo)).Interior.Pattern = &H1 'xlSolid
                    MsExcelSheet.Range(MsExcelSheet.Cells(xRow, vColNo), MsExcelSheet.Cells(xRow, vColNo)).Interior.PatternColorIndex = -4105 'xlAutomatic
                    
                End If 'Close respective IF..THEN block statement
                
                'Apply Borders in each Cell
                MsExcelSheet.Range(MsExcelSheet.Cells(xRow, vColNo), MsExcelSheet.Cells(xRow, vColNo)).BorderAround &H1, &H2, -4105, &H80&
                
                MsExcelSheet.Range(MsExcelSheet.Cells(xRow, vColNo), MsExcelSheet.Cells(xRow, vColNo)).Font.Bold = (xLv.ColumnHeaders(xCol).Text = "Kr" Or xLv.ColumnHeaders(xCol).Text = "Kshs" Or VBA.InStr(xLv.ColumnHeaders(xCol).Text, "Net ") <> &H0)
                
                If vFirstVisibleCol = &H0 Then vFirstVisibleCol = xCol
                
            Else
                xFreezeCol = xFreezeCol - &H1
            End If 'Close respective IF..THEN block statement
            
            iCnt = iCnt + (Frm_PleaseWait.lblProgressBar.Width / MaxPBarVal)
            Frm_PleaseWait.ImgProgressBar.Width = iCnt
            Frm_PleaseWait.lblStatus.Caption = VBA.Int(((VBA.CLng(iCnt) * 100) / Frm_PleaseWait.lblProgressBar.Width)) & "% Complete"
            
        Next xCol 'Increment counter by 1 to move to the next column in the Lv
        
        VBA.DoEvents 'Yield execution so that the operating system can process other events
        
    Next xRow 'Increment counter by 1 to move to the next row in the Lv
    
FormattingWorksheet:
    
    '---------------------------------FORMAT EXPORTED ITEMS----------------------------------
    
    Frm_PleaseWait.lblInfo.Caption = "Formatting Worksheet..."
    
    If Not xlsAborted Then
        
        Frm_PleaseWait.Visible = False
        
        MsExcelSheet.PageSetup.PaperSize = VBA.IIf(vMsgBox("Which type of printing paper do you want to use?", vbQuestion + vbCustomButtons + vbDefaultButton2, App.Title & " : Excel Export", xLv.Parent, , , , , "1|1", "&Legal Paper|&A4 Paper") = &H0, &H9, &H5)
        
        'If exported Landscape paper orientation then...
        If xLandScape Then
            MsExcelSheet.PageSetup.Orientation = &H2 'xlLandscape
        Else
            MsExcelSheet.PageSetup.Orientation = VBA.IIf(vMsgBox("Which type of paper orientation do you want to use?", vbQuestion + vbCustomButtons + vbDefaultButton2, App.Title & " : Excel Export", xLv.Parent, , , , , "1|2", "&Landscape|&Portrait") = &H0, &H1, &H2)
        End If 'Close respective IF..THEN block statement
        
        Frm_PleaseWait.Visible = True
        
    End If
    
    MsExcelSheet.PageSetup.TopMargin = MsExcelApp.CentimetersToPoints(0.8)
    MsExcelSheet.PageSetup.LeftMargin = MsExcelApp.CentimetersToPoints(0.5)
    MsExcelSheet.PageSetup.BottomMargin = MsExcelApp.CentimetersToPoints(0.5)
    MsExcelSheet.PageSetup.RightMargin = MsExcelApp.CentimetersToPoints(0)
    MsExcelSheet.PageSetup.HeaderMargin = MsExcelApp.CentimetersToPoints(0)
    MsExcelSheet.PageSetup.FooterMargin = MsExcelApp.CentimetersToPoints(0)
    
    'Set Font Name to 'Tahoma'
    MsExcelSheet.Range(MsExcelSheet.Cells(&H1, &H1), MsExcelSheet.Cells(xLv.ListItems.Count + xStartRow, xMaxCols)).Font.Name = "Tahoma"
    
    'Set Font Size to the specified value
    MsExcelSheet.Range(MsExcelSheet.Cells(&H1, &H1), MsExcelSheet.Cells(xLv.ListItems.Count + xStartRow, xMaxCols)).Font.Size = xFontSize
    
    'Set ForeColor to Blue
    MsExcelSheet.Range(MsExcelSheet.Cells(xStartRow + &H1, &H1), MsExcelSheet.Cells(xLv.ListItems.Count + &H3, xMaxCols)).Font.Color = &HC00000
    
    If xLv.ListItems(xLv.ListItems.Count).Tag = "Function" Then
        
        'Set ForeColor to Brown
        MsExcelSheet.Range(MsExcelSheet.Cells(xLv.ListItems.Count + xStartRow, &H1), MsExcelSheet.Cells(xLv.ListItems.Count + xStartRow, xMaxCols)).Font.Color = &H80&
        MsExcelSheet.Range(MsExcelSheet.Cells(xLv.ListItems.Count + xStartRow, &H1), MsExcelSheet.Cells(xLv.ListItems.Count + xStartRow, xMaxCols)).Font.Bold = True
        
    End If 'Close respective IF..THEN block statement
    
    'Auto-Resize column widths to fit their contents
    MsExcelSheet.Range(MsExcelSheet.Cells(&H1, &H1), MsExcelSheet.Cells(xLv.ListItems.Count + &H3, xMaxCols)).Columns.AutoFit
    
    'If there are more than one column then...
    If xMaxCols > &H1 And xFreezeCol >= vFirstVisibleCol And xFreezeCol <> &H0 Then
        
        'Freeze the panes
        MsExcelSheet.Cells(xStartRow + &H1, xFreezeCol + &H1).Select
        MsExcelSheet.Cells(xStartRow + &H1, xFreezeCol + &H1).Activate
        MsExcelApp.ActiveWindow.FreezePanes = True
        
    End If 'Close respective IF..THEN block statement
    
    MsExcelSheet.Cells(xStartRow + &H1, &H1).Select
    
    ExportLvToExcel = True 'Denote that the Export was successful
    
Exit_ExportLvToExcel:
    
    vCancelOperation = &H0 'Initialize variable
    
    Unload Frm_PleaseWait 'Unload Form from the memory
    
    'COMMENTED TO LEAVE THE WORKBOOK VISIBLE TO THE USER
'    MsExcelApp.ActiveWorkbook.Close 'Close the Excel WorkBook
'    MsExcelApp.Quit 'Close the Excel Application
    
    'If the Export was successful then...
    If ExportLvToExcel Then
        
        'Indicate that a process or operation is complete.
        Screen.MousePointer = vbDefault
        
        If xlsAborted Then
            
            'Inform User
            vMsgBox "The Data export to Excel has been aborted by User.", vbExclamation, App.Title & " : Data Transfer Aborted", xLv.Parent
            
        Else
            
            'Inform User
            vMsgBox "The Data has successfully been exported to Excel.", vbInformation, App.Title & " : Data Transfer", xLv.Parent
            
        End If 'Close respective IF..THEN block statement
        
    End If 'Close respective IF..THEN block statement
    
    MsExcelApp.Visible = True: MsExcelApp.WindowState = &HFFFFEFD7 'xlMaximized
    
    'Disassociate object variables from their actual objects
    Set MsExcelSheet = Nothing: Set MsExcelApp = Nothing
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Function 'Quit this Procedure
    
Handle_ExportLvToExcel_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error Exporting Lv Data To Excel - " & Err.Number, xLv.Parent
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Resume execution at the specified Label
    Resume Exit_ExportLvToExcel
    
End Function

Public Function ExportTvToExcel(vTv As TreeView, Optional xlsStartRow& = &H2, Optional xlsMaxCols& = &H0, Optional xlsHeading$, Optional ImgProgressBar As Object = Nothing, Optional ShowPleaseWait As Boolean = False) As Boolean
On Local Error GoTo Handle_ExportTvToExcel_Error
    
    'If the specified Treeview has no data for transfer then...
    If vTv.Nodes.Count = &H0 Then
        
        'Warn User
        vMsgBox "There are no items to be exported to Ms Excel", vbExclamation, App.Title & " : Operation Aborted", vTv.Parent
        Exit Function 'Quit this Function
        
    End If 'Close respective IF..THEN block statement
    
    'Confirm if the User really wants to Export the Data to Ms Excel. If not then Quit this Function
    If vMsgBox("Are you sure you want to transfer the data to Ms Excel?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title & " : Confirmation", vTv.Parent) = vbNo Then Exit Function
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    Dim xRow&, xCol&, xMaxCols&
    Dim xQry$, xStatement$, xValidExportCols$
    
    Dim MsExcelApp As Object
    Dim MsExcelSheet As Object
    
    Set MsExcelApp = CreateObject("Excel.Application")
    
    MsExcelApp.Workbooks.Add
    
    'If the Excel version is 8.0 or greater then...
    If Val(MsExcelApp.Application.Version) >= &H8 Then
        Set MsExcelSheet = MsExcelApp.ActiveSheet
    Else 'If the Excel version is less than 8.0 then...
        Set MsExcelSheet = MsExcelApp
    End If 'Close respective IF..THEN block statement
    
    Dim MaxPBarVal&, iCnt#
    
    If ShowPleaseWait Then
        
        Frm_PleaseWait.ImgProgressBar.Width = &H0
        Frm_PleaseWait.lblInfo.Caption = "Transfering data to Worksheet..."
        
        'For Items Transfer
        MaxPBarVal = vTv.Nodes.Count
        
        'Show the 'Please Wait' Form to the User
        CenterForm Frm_PleaseWait, vTv.Parent, False
        
    End If 'Close respective IF..THEN block statement
    
    xMaxCols = VBA.IIf(xlsMaxCols > &H0, xlsMaxCols, &H0)
    
    '--------------------------------------------------------------------------------------------
    '                                   EXPORT DATA TO MS EXCEL
    '--------------------------------------------------------------------------------------------
    
    Dim mNode As Node
    Dim xlsCnt&, xlsRowCnt&
    
    'Assign value to cell
    MsExcelSheet.Cells(&H1, &H1) = School.Name
    
    'Set Cell's Font Size
    MsExcelSheet.Cells(&H1, &H1).Font.Size = &HC
    
    'Set Cell's Font Bold
    MsExcelSheet.Cells(&H1, &H1).Font.Bold = True
    
    'Set Cell's Font Underline
    MsExcelSheet.Cells(&H1, &H1).Font.Underline = True
    
    'Set Cell's contents to center of the cell
    MsExcelSheet.Cells(&H1, &H1).HorizontalAlignment = &H1 'wdHorizontalLineAlignCenter
    
    'Assign value to cell
    MsExcelSheet.Cells(&H2, &H1) = App.Title & " " & xlsHeading
    
    'Set Cell's Font Size
    MsExcelSheet.Cells(&H2, &H1).Font.Size = &HC
    
    'Set Cell's Font Bold
    MsExcelSheet.Cells(&H2, &H1).Font.Bold = True
    
    'Set Cell's Font Underline
    MsExcelSheet.Cells(&H2, &H1).Font.Underline = True
    
    'Set Cell's contents to center of the cell
    MsExcelSheet.Cells(&H2, &H1).HorizontalAlignment = &H1 'wdHorizontalLineAlignCenter
    
    xlsStartRow = &H4
    
    'For each item in the Treeview...
    For xlsCnt = xlsStartRow To vTv.Nodes.Count + (xlsStartRow - &H1) Step &H1
        
        'If the User has cancelled the process then branch to the specified Label
        If vCancelOperation = &H2 Then GoTo Exit_ExportTvToExcel
        
        'If the User has paused the process then call function to wait
        If vCancelOperation = &H1 Then Call Wait
        
        Set mNode = vTv.Nodes(xlsCnt - (xlsStartRow - &H1))
        
        vArrayList = VBA.Split(mNode.FullPath, vTv.PathSeparator)
        
        'Assign value to cell
        MsExcelSheet.Cells(xlsCnt, UBound(vArrayList) + &H1) = mNode.Text
        If vTv.Checkboxes Then MsExcelSheet.Cells(xlsCnt, xMaxCols) = VBA.IIf(mNode.Checked, "{Allowed}", "{Limited}")
        
        'Set Cell's Font Size
        MsExcelSheet.Cells(xlsCnt, UBound(vArrayList) + &H1).Font.Size = VBA.IIf((&HB - UBound(vArrayList)) < &H9, &H9, &HB - UBound(vArrayList))
        
        'Set Cell's Font Bold
        MsExcelSheet.Cells(xlsCnt, UBound(vArrayList) + &H1).Font.Bold = (UBound(vArrayList) < &H2 And xMaxCols > &H1)
        
        'Set Cell's Font Underline
        MsExcelSheet.Cells(xlsCnt, UBound(vArrayList) + &H1).Font.Underline = (UBound(vArrayList) = &H0)
        
        If UBound(vArrayList) < &H2 Then
            
            'Merge Headers
            'MsExcelSheet.Range(MsExcelSheet.Cells(xlsCnt, UBound(vArrayList) + &H1), MsExcelSheet.Cells(xlsCnt, &H3)).Merge
            
        End If 'Close respective IF..THEN block statement
        
        If ShowPleaseWait Then
                        
            iCnt = iCnt + (Frm_PleaseWait.lblProgressBar.Width / MaxPBarVal)
            Frm_PleaseWait.ImgProgressBar.Width = iCnt
            Frm_PleaseWait.lblStatus.Caption = VBA.Int(((VBA.CLng(iCnt) * 100) / Frm_PleaseWait.lblProgressBar.Width)) & "% Complete"
            
        End If 'Close respective IF..THEN block statement
        
        VBA.DoEvents 'Yield execution so that the operating system can process other events
        
    Next xlsCnt 'Increment counter by value in the step option
    
    'Merge Company Name Cells
    MsExcelSheet.Range(MsExcelSheet.Cells(&H1, &H1), MsExcelSheet.Cells(&H1, xMaxCols)).Merge
    
    'Merge Header
    MsExcelSheet.Range(MsExcelSheet.Cells(&H2, &H1), MsExcelSheet.Cells(&H2, xMaxCols)).Merge
    
    'Set Font Name to 'Tahoma'
    MsExcelSheet.Range(MsExcelSheet.Cells(&H1, &H1), MsExcelSheet.Cells(vTv.Nodes.Count + (xlsStartRow - &H1), xMaxCols)).Font.Name = "Tahoma"
    
    'Apply Thick Border around the exported contents                                                 xlContinuous, xlThick, xlColorIndexAutomatic, Brown
    'MsExcelSheet.Range(MsExcelSheet.Cells(xlsStartRow, &H1), MsExcelSheet.Cells(vTv.Nodes.Count + (xlsStartRow - &H1), xMaxCols)).BorderAround &H1, &H4, -4105, &H80&
    
    'Set ForeColor to Blue
    MsExcelSheet.Range(MsExcelSheet.Cells(&H1, &H1), MsExcelSheet.Cells(vTv.Nodes.Count + (xlsStartRow - &H1), xMaxCols)).Font.Color = &HC00000
    
    'Auto-Resize column widths to fit their contents
    MsExcelSheet.Range(MsExcelSheet.Cells(&H1, &H1), MsExcelSheet.Cells(vTv.Nodes.Count + (xlsStartRow - &H1), xMaxCols)).Columns.AutoFit
    
    MsExcelSheet.Cells(xlsStartRow, &H1).Select
    
    ExportTvToExcel = True 'Denote successful data transfer to Excel
    
Exit_ExportTvToExcel:
    
    'COMMENTED TO LEAVE THE WORKBOOK VISIBLE TO THE USER
'    MsExcelApp.ActiveWorkbook.Close 'Close the Excel WorkBook
'    MsExcelApp.Quit 'Close the Excel Application
    
    If ShowPleaseWait Then Unload Frm_PleaseWait
    
    'If the Export was successful then...
    If ExportTvToExcel Then
        
        'Indicate that a process or operation is complete.
        Screen.MousePointer = vbDefault
        
        'Inform User
        vMsgBox "The Data has successfully been exported to Excel.", vbInformation, App.Title & " : Data Transfer", vTv.Parent
        
    End If 'Close respective IF..THEN block statement
    
    MsExcelApp.Visible = True: MsExcelApp.WindowState = &HFFFFEFD7 'xlMaximized
    
    'Disassociate object variables from their actual objects
    Set MsExcelSheet = Nothing: Set MsExcelApp = Nothing
    
    vTv.Visible = True 'Show Treeview control
    
    vCancelOperation = &H0 'Initialize variable
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Function 'Quit this Function
    
Handle_ExportTvToExcel_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error Exporting Tv Data To Excel - " & Err.Number, vTv.Parent
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Resume execution at the specified Label
    Resume Exit_ExportTvToExcel
    
End Function

Public Function ExportTvToWord(vTv As TreeView, Optional wdHeading$, Optional ImgProgressBar As Object = Nothing, Optional ShowPleaseWait As Boolean = False, Optional xMaxNodes& = 2) As Boolean
On Local Error GoTo Handle_ExportTvToWord_Error
    
    'If the specified Treeview has no data for transfer then...
    If vTv.Nodes.Count = &H0 Then
        
        'Warn User
        vMsgBox "There are no items to be exported to Ms Word", vbExclamation, App.Title & " : Operation Aborted", vTv.Parent
        Exit Function 'Quit this Function
        
    End If 'Close respective IF..THEN block statement
    
    'Confirm if the User really wants to Export the Data to Ms Word. If not then Quit this Function
    If vMsgBox("Are you sure you want to transfer the data to Ms Word?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title & " : Confirmation", vTv.Parent) = vbNo Then Exit Function
    
    Dim MaxPBarVal&, iCnt#
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    If ShowPleaseWait Then
        
        Frm_PleaseWait.ImgProgressBar.Width = &H0
        Frm_PleaseWait.lblInfo.Caption = "Transfering data to Worksheet..."
        
        'For Items Transfer
        MaxPBarVal = vTv.Nodes.Count
        
        'Show the 'Please Wait' Form to the User
        CenterForm Frm_PleaseWait, vTv.Parent, False
        
    End If 'Close respective IF..THEN block statement
    
    Dim MsWordApp As Object
    Dim MsWordDoc As Object
    
    Set MsWordApp = CreateObject("Word.Application")
    Set MsWordDoc = MsWordApp.Documents.Add(, , , True)
    
    MsWordDoc.Activate
    MsWordApp.ActiveWindow.DisplayRulers = True
    
    'Write summary Report heading to the created Word Document
    MsWordApp.Selection.style = -&H2 'wdStyleHeading1
    MsWordApp.Selection.Font.Underline = &H1 'wdUnderlineSingle
    MsWordApp.Selection.Paragraphs.Alignment = &H1 'wdAlignParagraphCenter
    MsWordApp.Selection.TypeText VBA.UCase$(School.Name)
    MsWordApp.Selection.TypeParagraph
    MsWordApp.Selection.style = -&H3 'wdStyleHeading2
    MsWordApp.Selection.Font.Underline = &H1 'wdUnderlineSingle
    MsWordApp.Selection.TypeText wdHeading
    MsWordApp.Selection.Font.Underline = &H0 'wdUnderlineNone
    MsWordApp.Selection.TypeParagraph
    
    MsWordApp.Selection.Paragraphs.Alignment = 2 ' xlHAlignLeft
    MsWordApp.Selection.Paragraphs.SpaceAFTER = &H1
    
    'For each item in the Treeview
    For vIndex(&H0) = &H1 To vTv.Nodes.Count Step &H1
        
        'If the User has cancelled the process then branch to the specified Label
        If vCancelOperation = &H2 Then GoTo Exit_ExportTvToWord
        
        'If the User has paused the process then call function to wait
        If vCancelOperation = &H1 Then Call Wait
        
        vArrayList = VBA.Split(vTv.Nodes(vIndex(&H0)).FullPath, vTv.PathSeparator)
        MsWordApp.Selection.style = VBA.IIf(UBound(vArrayList) = &H0, -&H4, VBA.IIf(UBound(vArrayList) = &H1, -&H5, -&H1)) 'wdStyleHeading3, wdStyleHeading4, wdStyleNormal
        
        vBuffer(&H0) = vTv.Nodes(vIndex(&H0)).Text
        
        If UBound(vArrayList) >= &H2 Then
            
            If UBound(vArrayList) = &H2 And xMaxNodes > &H1 Then MsWordApp.Selection.Font.Bold = True Else MsWordApp.Selection.Font.Bold = False
            
            If vTv.Parent.Name = "Frm_SummaryInfo" Then
                
                vArrayListTmp = VBA.Split(vTv.Nodes(vIndex(&H0)).Text, ":")
                vBuffer(&H0) = vArrayListTmp(&H0) & VBA.String$(100 - VBA.LenB(vArrayListTmp(&H0)), ".") & vArrayListTmp(&H1)
                
            End If 'Close respective IF..THEN block statement
            
        End If 'Close respective IF..THEN block statement
        
        If vTv.Nodes(vIndex(&H0)).Children = &H0 Then MsWordApp.Selection.Paragraphs.Space1
        If vTv.Nodes(vIndex(&H0)).Children = &H0 Then MsWordApp.Selection.Font.Size = 10
        
                                                                            'wdUnderlineSingle, wdUnderlineNone
        MsWordApp.Selection.Font.Underline = VBA.IIf(UBound(vArrayList) = &H0 And xMaxNodes > &H1, &H1, &H0)
        MsWordApp.Selection.TypeText VBA.String(UBound(vArrayList), vbTab) & vBuffer(&H0) & VBA.IIf(vTv.Checkboxes, VBA.IIf(vTv.Nodes(vIndex(&H0)).Checked, " - {Allowed}", " - {Limited}"), VBA.vbNullString)
        MsWordApp.Selection.TypeParagraph
        
        If ShowPleaseWait Then
                        
            iCnt = iCnt + (Frm_PleaseWait.lblProgressBar.Width / MaxPBarVal)
            Frm_PleaseWait.ImgProgressBar.Width = iCnt
            Frm_PleaseWait.lblStatus.Caption = VBA.Int(((VBA.CLng(iCnt) * 100) / Frm_PleaseWait.lblProgressBar.Width)) & "% Complete"
            
        End If 'Close respective IF..THEN block statement
        
        VBA.DoEvents 'Yield execution so that the operating system can process other events
        
    Next vIndex(&H0) 'Move to the next item in the Treeview
    
    ExportTvToWord = True 'Denote successful data transfer to Word
    
Exit_ExportTvToWord:
    
    vTv.Visible = True 'Show Treeview control
    
    If ShowPleaseWait Then Unload Frm_PleaseWait
    
    'If the Export was successful then...
    If ExportTvToWord Then
        
        'Indicate that a process or operation is complete.
        Screen.MousePointer = vbDefault
        
        'Inform User
        vMsgBox "The Data has successfully been exported to Word.", vbInformation, App.Title & " : Data Transfer", vTv.Parent
        
    End If 'Close respective IF..THEN block statement
    
    MsWordApp.Visible = True
    
    Set MsWordApp = Nothing
    
    vCancelOperation = &H0 'Initialize variable
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Function 'Quit this Function
    
Handle_ExportTvToWord_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error Exporting Tv Data To Word - " & Err.Number, vTv.Parent
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Resume execution at the specified Label
    Resume Exit_ExportTvToWord
    
End Function

Public Function FillListView(vLv As ListView, iSQL$, Optional iHideFlds$, Optional iExcludeFlds$, Optional iTagField$, Optional iFieldCondition$, Optional iFieldConditionColor$, Optional ShowPleaseWait As Boolean, Optional iLastRowTotal$ = VBA.vbNullString, Optional iAfterFillLv As Boolean = False, Optional iImgIndex& = &H1, Optional iCheckAll As Boolean = True, Optional iExcludeRecord$ = VBA.vbNullString, Optional iToolTipText$ = VBA.vbNullString) As Long
On Local Error GoTo Handle_FillListView_Error
    
    Dim iCounter&, iCnt#
    Dim strTxtWidth() As String
    Dim LvItemDataType() As String
    Dim vRow&, vCol&, vFirstVisibleCol&
    Dim LvState, iWaitMsgDisplayed As Boolean
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    LvState = vLv.Visible
    vLv.Visible = False 'Hide Listview for faster items addition
    
    'Remove all Listview Rows and Columns
    vLv.ListItems.Clear
    vLv.ColumnHeaders.Clear
    vLv.Parent.iSearchIDs = VBA.vbNullString
    
    If VBA.LenB(VBA.Trim$(vAdoCNN.ConnectionString)) = &H0 Then Call ConnectDB 'Call Function in this Module to connect to Software's Database
    Set vRs = New ADODB.Recordset
    With vRs 'Execute a series of statements on vRs recordset
        
        .Open iSQL, vAdoCNN, adOpenKeyset, adLockReadOnly
        
        'Display the 'Please Wait' message box only when Records to be fetched are greater than 100
        ShowPleaseWait = (.RecordCount > 100)
        iCounter = .Fields.Count + .RecordCount
        
        If ShowPleaseWait Then
            
            Frm_PleaseWait.ImgProgressBar.Width = &H0
            Frm_PleaseWait.ImgProgressBar.Left = Frm_PleaseWait.lblProgressBar.Left
            Frm_PleaseWait.ImgProgressBar.Visible = True
            
            'Show the 'Please Wait' Form to the User
            CenterForm Frm_PleaseWait, vLv.Parent, False
            
        End If 'Close respective IF..THEN block statement
        
        Dim iExcluded As Boolean
        Dim iLastRowSumFields$, sStr$
        
        For vCol = &H0 To .Fields.Count - &H1 Step &H1
            sStr = sStr & "|" & .Fields(vCol).Name
        Next vCol
        
        If VBA.InStr(sStr & "|", "|A|A-|B|B-|B+|C|C-|C+|D|D-|D+|E|X|Y|Z|") <> &H0 Then sStr = VBA.Replace(sStr & "|", "|A|A-|B|B-|B+|C|C-|C+|D|D-|D+|E|X|Y|Z|", "|A|A-|B+|B|B-|C+|C|C-|D+|D|D-|E|X|Y|Z|")
        
        'If the data starts with '|' character then remove it
        If VBA.Left$(sStr, &H1) = "|" Then sStr = VBA.Right$(sStr, VBA.Len(sStr) - &H1)
        
        'If the data ends with '|' character then remove it
        If VBA.Right$(sStr, &H1) = "|" Then sStr = VBA.Left$(sStr, VBA.Len(sStr) - &H1)
        
        'VB.Clipboard.Clear: VB.Clipboard.SetText iSQL
        
        Dim nCols() As String
        
        nCols = VBA.Split(sStr, "|"): sStr = VBA.vbNullString
        
        If iLastRowTotal <> VBA.vbNullString Then iLastRowTotal = VBA.Replace(iLastRowTotal, ";", "|")
        
        'Load each table field into the specified Listview control as column header
        
        For vCol = &H0 To UBound(nCols) Step &H1
            
            'If the User has cancelled the process then branch to the specified Label
            If vCancelOperation = &H2 And ShowPleaseWait Then GoTo Exit_FillListView
            
            'If the User has paused the process then call function to wait
            If vCancelOperation = &H1 And ShowPleaseWait Then Call Wait
            
            'If the criteria specified is a field position then...
            If VBA.IsNumeric(VBA.Replace(VBA.Replace(iExcludeFlds, ";", "|"), "|", VBA.vbNullString)) Then
                iExcluded = (VBA.InStr("|" & VBA.Replace(iExcludeFlds, ";", "|") & "|", "|" & vCol + &H1 & "|") <> &H0)
            Else 'If the criteria specified is a field name then...
                iExcluded = (VBA.InStr("|" & VBA.Replace(iExcludeFlds, ";", "|") & "|", "|" & .Fields(vCol).Name & "|") <> &H0)
            End If 'Close respective IF..THEN block statement
            
            'If the Field should not be excluded during retrieval then...
            If Not iExcluded Then
                
                If VBA.IsDate(nCols(vCol)) Then sStr = VBA.Format$(nCols(vCol), "MMMM") Else sStr = nCols(vCol)
                
                'Add the Field name to Listview as column header
                vLv.ColumnHeaders.Add vLv.ColumnHeaders.Count + &H1, , sStr
                vLv.ColumnHeaders(vLv.ColumnHeaders.Count).Tag = nCols(vCol)
                
                'Reallocate storage space to fit the available columns
                ReDim Preserve LvItemDataType(vCol) As String
                
                'Get column datatypes
                
                Select Case .Fields(nCols(vCol)).Type
                    
                    Case adLongVarBinary: LvItemDataType(vLv.ColumnHeaders.Count - &H1) = "O"
                    Case adDate: LvItemDataType(vLv.ColumnHeaders.Count - &H1) = "D"
                    Case adBoolean: LvItemDataType(vLv.ColumnHeaders.Count - &H1) = "B"
                    Case adCurrency: LvItemDataType(vLv.ColumnHeaders.Count - &H1) = "C"
                    Case adNumeric, adInteger, adDouble: LvItemDataType(vLv.ColumnHeaders.Count - &H1) = "N"
                    Case Else: LvItemDataType(vLv.ColumnHeaders.Count - &H1) = "T"
                    
                End Select 'Close SELECT..CASE block statement 'End Select 'Close SELECT..CASE block statement..CASE statement
                
                ReDim Preserve strTxtWidth(vLv.ColumnHeaders.Count - &H1) As String
                
                'If the Field should be hidden then...
                If VBA.InStr("|" & VBA.Replace(iHideFlds, ";", "|") & "|", "|" & vCol + &H1 & "|") <> &H0 Then
                    
                    'Set its Width value array to zero
                    strTxtWidth(vLv.ColumnHeaders.Count - &H1) = &H0
                    
                Else 'If the Field should not be hidden then...
                    
                    'If the no of characters in the field name is greater than 40 then...
                    If VBA.LenB(sStr) < 80 Then
                        If VBA.Val(strTxtWidth(vLv.ColumnHeaders.Count - &H1)) < VBA.Val(vLv.Parent.TextWidth(sStr)) Then strTxtWidth(vLv.ColumnHeaders.Count - &H1) = vLv.Parent.TextWidth(sStr) + 200
                    Else
                        strTxtWidth(vLv.ColumnHeaders.Count - &H1) = 1000
                    End If 'Close respective IF..THEN block statement
                    
                End If 'Close respective IF..THEN block statement
                
                'Align to right if the column involves finances
                If (LvItemDataType(vLv.ColumnHeaders.Count - &H1) = "C" Or LvItemDataType(vLv.ColumnHeaders.Count - &H1) = "N") And vLv.ColumnHeaders.Count > &H1 Then vLv.ColumnHeaders(vLv.ColumnHeaders.Count).Alignment = lvwColumnRight
                
                'Get the position of the first visible column
                vFirstVisibleCol = VBA.IIf(vFirstVisibleCol = &H0 And strTxtWidth(vLv.ColumnHeaders.Count - &H1) <> &H0, vLv.ColumnHeaders.Count, vFirstVisibleCol)
                
            End If 'Close respective IF..THEN block statement
            
            If ShowPleaseWait Then
                
                iCnt = iCnt + (Frm_PleaseWait.lblProgressBar.Width / iCounter)
                Frm_PleaseWait.ImgProgressBar.Width = iCnt
                Frm_PleaseWait.lblStatus.Caption = VBA.Int(((VBA.CLng(iCnt) * 100) / Frm_PleaseWait.lblProgressBar.Width)) & "% Complete"
                
            End If 'Close respective IF..THEN block statement
            
            VBA.DoEvents 'Yield execution so that the operating system can process other events
            
            If VBA.LenB(VBA.Trim$(iLastRowTotal)) <> &H0 Then If VBA.InStr("|" & iLastRowTotal & "|", "|" & vCol + &H1 & "|") <> &H0 Then iLastRowSumFields = iLastRowSumFields & "|" & sStr
            
        Next vCol 'Move to the next field
        
        'If the data starts with '|' character then remove it
        If VBA.Left$(iLastRowSumFields, &H1) = "|" Then iLastRowSumFields = VBA.Right$(iLastRowSumFields, VBA.Len(iLastRowSumFields) - &H1)
        
        'Assign column datatypes
        vLv.Parent.iLvwItemDataType = VBA.Join(LvItemDataType, "|")
        
        'If there are no visible columns then...
        If vFirstVisibleCol = &H0 Then
            
            vMsgBox "There are no Fields defined to be viewed by User.", vbExclamation, App.Title & " : No Visible Fields", vLv.Parent
            GoTo ResizeColumns 'Branch to resize columns
            
        End If 'Close respective IF..THEN block statement
        
        'Load each Record into the specified Listview control as list items
        
        Dim Lst
        Dim iSmallIcons
        Dim RecFldValue$, iStatus$
        Dim vArrayListTmp1() As String
        
        'If Listview Icon has been specified then associate it with the Icon index specified
        If Not Nothing Is vLv.SmallIcons And iImgIndex > &H0 Then iSmallIcons = iImgIndex
        
        Dim dtFldDate, dtDefinedDate As DTPicker
        
        'If coloring condition has been specified then...
        If VBA.LenB(VBA.Trim$(iFieldCondition)) <> &H0 Then
            
            'Add virtual date fields on the Source Form
            Set dtFldDate = vLv.Parent.Controls.Add("MSCOMCTL2.DTPicker.2", "DTP1")
            Set dtDefinedDate = vLv.Parent.Controls.Add("MSCOMCTL2.DTPicker.2", "DTP2")
            
        End If 'Close respective IF..THEN block statement
        
        Dim nArray() As String
        Dim nArrayTmp() As String
        Dim nTagField() As String
        Dim iLastRowTotals&, iVal&
        Dim nRs As New ADODB.Recordset
        Dim iLastRowSumTotals() As Long
        
        Set nRs = New ADODB.Recordset
        
        'If coloring condition has been specified then...
        If VBA.LenB(VBA.Trim$(iFieldCondition)) <> &H0 Or VBA.LenB(VBA.Trim$(iTagField)) <> &H0 Or VBA.LenB(VBA.Trim$(iExcludeRecord)) <> &H0 Then
            
            Dim nStr$
            
            nStr = VBA.vbNullString
            
            If VBA.InStr(iSQL, "TRANSFORM ") = &H0 Then
                
                nStr = VBA.Mid$(iSQL, VBA.InStr(iSQL, "SELECT "))
                nStr = VBA.Replace(iSQL, VBA.Mid$(nStr, &H1, VBA.InStr(nStr, " FROM ") + &H4), "SELECT * FROM")
                
            Else
                nStr = iSQL
            End If
            
            If VBA.LenB(VBA.Trim$(vAdoCNN.ConnectionString)) = &H0 Then Call ConnectDB(, False) 'Call Function in this Module to connect to Software's Database
            
            'Retrieve Records with all fields
            nRs.Open nStr, vAdoCNN, adOpenKeyset, adLockReadOnly
            
        End If 'Close respective IF..THEN block statement
        
        ReDim iLastRowSumTotals(VBA.CLng(vLv.ColumnHeaders.Count - &H1)) As Long
        
        'For each Record retrieved from the database table...
        For vRow = &H1 To .RecordCount Step &H1
            
            If .EOF Then Exit For
            
            'If the User has cancelled the process then branch to the specified Label
            If vCancelOperation And ShowPleaseWait Then GoTo Exit_FillListView
            
            iStatus = VBA.vbNullString 'Initialize variable
            
            'For each field displayed in the Listview columns...
            For vCol = &H1 To vLv.ColumnHeaders.Count Step &H1
                
                'Add an item for a new List Item else add a sub-item
                If vCol = &H1 Then Set Lst = vLv.ListItems.Add(vLv.ListItems.Count + &H1, , , , iSmallIcons) Else Set Lst = vLv.ListItems(vLv.ListItems.Count).ListSubItems.Add(vLv.ListItems(vLv.ListItems.Count).ListSubItems.Count + &H1)
                
                If VBA.IsNull(vRs(vLv.ColumnHeaders(vCol).Tag)) Then RecFldValue = VBA.IIf(VBA.InStr(iSQL, "GradeCount") <> &H0, &H0, " ") Else RecFldValue = vRs(vLv.ColumnHeaders(vCol).Tag)
                
                'If the datatype of the column is boolean then...
                If LvItemDataType(vCol - &H1) = "B" Then
                    
                    Select Case RecFldValue
                        Case "TRUE", "ON": RecFldValue = "Yes"
                        Case "FALSE", "OFF": RecFldValue = "No"
                    End Select 'Close SELECT..CASE block statement
                    
                End If 'Close respective IF..THEN block statement
                
                'Assign default color
                Lst.ForeColor = vLv.ForeColor
                
                'If coloring condition has been specified and the current column is the first in the list item then...
                If VBA.LenB(VBA.Trim$(iFieldCondition)) <> &H0 And vCol = &H1 Then
                    
                    Dim vFldSection(&H1) As String
1:
                    vArrayList = VBA.Split(iFieldCondition, "|")
                    vArrayListTmp1 = VBA.Split(iFieldConditionColor, "|")
                    
                    Lst.ToolTipText = VBA.vbNullString
                    
                    iStatus = "Check Condition"
                    
                    'For each condition specified...
                    For vIndex(&H0) = &H0 To UBound(vArrayList) Step &H1
                        
                        'If condition has not been specified then quit this FOR..LOOP
                        If iFieldCondition = VBA.vbNullString Then Exit For
                        
                        vArrayListTmp = VBA.Split(vArrayList(vIndex(&H0)), ":")
                        
                        vFldSection(&H0) = vArrayListTmp(&H0)
                        vFldSection(&H1) = vArrayListTmp(&H1)
                        
                        'If the fields contains valid data then...
                        If Not VBA.IsNull(nRs(vFldSection(&H0))) Then
                            
                            'Assign the fields value to the variable
                            vBuffer(&H0) = nRs(vArrayListTmp(&H0))
                            
                            'Convert the boolean value to YES for True or NO for False
                            If nRs(vArrayListTmp(&H0)).Type = adBoolean Or VBA.UCase$(nRs(vFldSection(&H0))) = "YES" Or VBA.UCase$(nRs(vFldSection(&H0))) = "NO" Then vFldSection(&H0) = VBA.IIf(VBA.UCase$(nRs(vFldSection(&H0))) = "YES", "True", VBA.IIf(VBA.UCase$(nRs(vFldSection(&H0))) = "NO", "False", nRs(vFldSection(&H0))))
                            
                            'If the datatype of the column is date/time then...
                            If nRs(vArrayListTmp(&H0)).Type = adDate Then
                                
                                dtFldDate.Value = nRs(vArrayListTmp(&H0))
                                If VBA.InStr(VBA.LCase$(vFldSection(&H1)), "date()") <> &H0 Then
                                    dtDefinedDate.Value = VBA.Date
                                Else
                                    vBuffer(&H0) = ReFormattedDate(vFldSection(&H1))
                                End If
                                
                                vFldSection(&H0) = vArrayListTmp(&H0)
                                vFldSection(&H1) = VBA.Replace(VBA.LCase$(vFldSection(&H1)), "date()", VBA.FormatDateTime(VBA.Date, vbShortDate))
                                
                            End If 'Close respective IF..THEN block statement
                            
                            If VBA.InStr(vFldSection(&H1), ">=") <> &H0 Then
                                
                                'If the datatype of the column is date/time then...
                                If nRs(vArrayListTmp(&H0)).Type = adDate Then
                                    
                                    'If the condition is true then change its color to the specified one
                                    If dtFldDate.Value >= dtDefinedDate.Value Then Lst.ForeColor = vArrayListTmp1(vIndex(&H0)): Exit For
                                    
                                ElseIf nRs(vArrayListTmp(&H0)).Type = adInteger Or nRs(vArrayListTmp(&H0)).Type = adNumeric Then
                                                                        
                                    'If the condition is true then change its color to the specified one
                                    If VBA.CLng(nRs(vArrayListTmp(&H0))) >= VBA.CLng(VBA.Replace(vFldSection(&H1), ">=", VBA.vbNullString)) Then Lst.ForeColor = vArrayListTmp1(vIndex(&H0)): Exit For
                                    
                                Else
                                    
                                    'If the condition is true then change its color to the specified one
                                    If nRs(vArrayListTmp(&H0)) >= VBA.Replace(VBA.Replace(vFldSection(&H1), ">=", VBA.vbNullString), "''", VBA.vbNullString) Then Lst.ForeColor = vArrayListTmp1(vIndex(&H0)): Exit For
                                    
                                End If 'Close respective IF..THEN block statement
                                
                            ElseIf VBA.InStr(vFldSection(&H1), "<=") <> &H0 Then
                                
                                'If the datatype of the column is date/time then...
                                If nRs(vArrayListTmp(&H0)).Type = adDate Then
                                    
                                    'If the condition is true then change its color to the specified one
                                    If dtFldDate.Value <= dtDefinedDate.Value Then Lst.ForeColor = vArrayListTmp1(vIndex(&H0)): Exit For
                                    
                                ElseIf nRs(vArrayListTmp(&H0)).Type = adInteger Or nRs(vArrayListTmp(&H0)).Type = adNumeric Then
                                    
                                    'If the condition is true then change its color to the specified one
                                    If VBA.CLng(nRs(vArrayListTmp(&H0))) <= VBA.CLng(VBA.Replace(vFldSection(&H1), "<=", VBA.vbNullString)) Then Lst.ForeColor = vArrayListTmp1(vIndex(&H0)): Exit For
                                    
                                Else
                                    
                                    'If the condition is true then change its color to the specified one
                                    If nRs(vArrayListTmp(&H0)) <= VBA.Replace(VBA.Replace(vFldSection(&H1), "<=", VBA.vbNullString), "''", VBA.vbNullString) Then Lst.ForeColor = vArrayListTmp1(vIndex(&H0)): Exit For
                                    
                                End If 'Close respective IF..THEN block statement
                                
                            ElseIf VBA.InStr(vFldSection(&H1), "<>") <> &H0 Then
                                
                                'If the datatype of the column is date/time then...
                                If nRs(vArrayListTmp(&H0)).Type = adDate Then
                                    
                                    'If the condition is true then change its color to the specified one
                                    If dtFldDate.Value <> dtDefinedDate.Value Then Lst.ForeColor = vArrayListTmp1(vIndex(&H0)): Exit For
                                    
                                ElseIf nRs(vArrayListTmp(&H0)).Type = adInteger Or nRs(vArrayListTmp(&H0)).Type = adNumeric Then
                                    
                                    'If the condition is true then change its color to the specified one
                                    If VBA.CLng(nRs(vArrayListTmp(&H0))) <> VBA.CLng(VBA.Replace(vFldSection(&H1), "<>", VBA.vbNullString)) Then Lst.ForeColor = vArrayListTmp1(vIndex(&H0)): Exit For
                                    
                                ElseIf nRs(vArrayListTmp(&H0)).Type = adBoolean Or VBA.UCase$(nRs(vArrayListTmp(&H0))) = "YES" Or VBA.UCase$(nRs(vArrayListTmp(&H0))) = "NO" Then
                                    
                                    vBuffer(&H0) = nRs(vArrayListTmp(&H0))
                                    If VBA.InStr("falsetrue", VBA.LCase$(vBuffer(&H0))) <> &H0 Then vBuffer(&H0) = VBA.IIf(VBA.LCase$(vBuffer(&H0)) = "false", "No", "Yes")
                                    
                                    'If the condition is true then change its color to the specified one
                                    If VBA.UCase$(vBuffer(&H0)) <> VBA.Replace(VBA.Replace(vFldSection(&H1), "<>", VBA.vbNullString), "''", VBA.vbNullString) Then Lst.ForeColor = vArrayListTmp1(vIndex(&H0)): Exit For
                                    
                                Else
                                    
                                    'If the condition is true then change its color to the specified one
                                    If nRs(vArrayListTmp(&H0)) <> VBA.Replace(VBA.Replace(vFldSection(&H1), "<>", VBA.vbNullString), "''", VBA.vbNullString) Then Lst.ForeColor = vArrayListTmp1(vIndex(&H0)): Exit For
                                    
                                End If 'Close respective IF..THEN block statement
                                
                            ElseIf VBA.InStr(vFldSection(&H1), ">") <> &H0 Then
                                
                                'If the datatype of the column is date/time then...
                                If nRs(vArrayListTmp(&H0)).Type = adDate Then
                                    
                                    'If the condition is true then change its color to the specified one
                                    If dtFldDate.Value > dtDefinedDate.Value Then Lst.ForeColor = vArrayListTmp1(vIndex(&H0)): Exit For
                                                                        
                                ElseIf nRs(vArrayListTmp(&H0)).Type = adInteger Or nRs(vArrayListTmp(&H0)).Type = adNumeric Then
                                    
                                    'If the condition is true then change its color to the specified one
                                    If VBA.CLng(nRs(vArrayListTmp(&H0))) > VBA.CLng(VBA.Replace(vFldSection(&H1), ">", VBA.vbNullString)) Then Lst.ForeColor = vArrayListTmp1(vIndex(&H0)): Exit For
                                    
                                Else
                                    
                                    'If the condition is true then change its color to the specified one
                                    If nRs(vArrayListTmp(&H0)) > VBA.Replace(VBA.Replace(vFldSection(&H1), ">", VBA.vbNullString), "''", VBA.vbNullString) Then Lst.ForeColor = vArrayListTmp1(vIndex(&H0)): Exit For
                                    
                                End If 'Close respective IF..THEN block statement
                                
                            ElseIf VBA.InStr(vFldSection(&H1), "<") <> &H0 Then
                                
                                'If the datatype of the column is date/time then...
                                If nRs(vArrayListTmp(&H0)).Type = adDate Then
                                    
                                    'If the condition is true then change its color to the specified one
                                    If dtFldDate.Value < dtDefinedDate.Value Then Lst.ForeColor = vArrayListTmp1(vIndex(&H0)): Exit For
                                    
                                ElseIf nRs(vArrayListTmp(&H0)).Type = adInteger Or nRs(vArrayListTmp(&H0)).Type = adNumeric Then
                                    
                                    'If the condition is true then change its color to the specified one
                                    If VBA.CLng(nRs(vArrayListTmp(&H0))) < VBA.CLng(VBA.Replace(vFldSection(&H1), "<", VBA.vbNullString)) Then Lst.ForeColor = vArrayListTmp1(vIndex(&H0)): Exit For
                                    
                                Else
                                    
                                    'If the condition is true then change its color to the specified one
                                    If nRs(vArrayListTmp(&H0)) < VBA.Replace(VBA.Replace(vFldSection(&H1), "<", VBA.vbNullString), "''", VBA.vbNullString) Then Lst.ForeColor = vArrayListTmp1(vIndex(&H0)): Exit For
                                    
                                End If 'Close respective IF..THEN block statement
                                
                            Else
                                
                                'If the datatype of the column is date/time then...
                                If nRs(vArrayListTmp(&H0)).Type = adDate Then
                                    
                                    'If the condition is true then change its color to the specified one
                                    If dtFldDate.Value = dtDefinedDate.Value Then Lst.ForeColor = vArrayListTmp1(vIndex(&H0)): Exit For
                                    
                                ElseIf nRs(vArrayListTmp(&H0)).Type = adInteger Or nRs(vArrayListTmp(&H0)).Type = adNumeric Then
                                    
                                    'If the condition is true then change its color to the specified one
                                    If nRs(vFldSection(&H0)) = VBA.Replace(vFldSection(&H1), "=", VBA.vbNullString) Then Lst.ForeColor = vArrayListTmp1(vIndex(&H0)): Exit For
                                    
                                ElseIf nRs(vArrayListTmp(&H0)).Type = adBoolean Or VBA.UCase$(nRs(vArrayListTmp(&H0))) = "YES" Or VBA.UCase$(nRs(vArrayListTmp(&H0))) = "NO" Then
                                    
                                    vBuffer(&H0) = nRs(vArrayListTmp(&H0))
                                    If VBA.InStr("falsetrue", VBA.LCase$(vBuffer(&H0))) <> &H0 Then vBuffer(&H0) = VBA.IIf(VBA.LCase$(vBuffer(&H0)) = "false", "No", "Yes")
                                    
                                    'If the condition is true then change its color to the specified one
                                    If VBA.UCase$(vBuffer(&H0)) = VBA.Replace(VBA.Replace(vFldSection(&H1), "=", VBA.vbNullString), "''", VBA.vbNullString) Then Lst.ForeColor = vArrayListTmp1(vIndex(&H0)): Exit For
                                    
                                Else
                                    
                                    'If the condition is true then change its color to the specified one
                                    If nRs(vArrayListTmp(&H0)) = VBA.Replace(VBA.Replace(vFldSection(&H1), "=", VBA.vbNullString), "''", VBA.vbNullString) Then Lst.ForeColor = vArrayListTmp1(vIndex(&H0)): Exit For
                                    
                                End If 'Close respective IF..THEN block statement
                                
                            End If 'Close respective IF..THEN block statement
                            
                        End If 'Close respective IF..THEN block statement
NextCondition:
                    Next vIndex(&H0) 'Move to the next condition
                    
                    vBuffer(&H0) = Lst.ForeColor
                    Lst.ToolTipText = VBA.IIf(VBA.CLng(vBuffer(&H0)) = VBA.CLng(vArrayListTmp1(vIndex(&H0))), VBA.Replace(vArrayList(vIndex(&H0)), ":", VBA.vbNullString), VBA.vbNullString)
                    
                End If 'Close respective IF..THEN block statement
                
                'If the datatype of the column is..
                Select Case LvItemDataType(vCol - &H1)
                    
                    'If the datatype of the column is boolean then...
                    Case "B":
                        
                        Select Case VBA.UCase$(RecFldValue)
                            Case "TRUE", "ON": RecFldValue = "Yes"
                            Case "FALSE", "OFF": RecFldValue = "No"
                        End Select 'Close SELECT..CASE block statement
                        
                    Case "O": 'OLE-OBJECT then...
                        
                        'Just display 'No Photo' when the Record is blank, or 'Has Photo' when not blank
                        RecFldValue = VBA.IIf(VBA.Trim$(RecFldValue) <> VBA.vbNullString, "[Has Photo]", "[No Photo]")
                        
                    Case "D": 'DATE-TIME then...
                        
                        '..Format it to Mon 11 Oct 2010
                        RecFldValue = VBA.Format$(RecFldValue, "ddd dd MMM yyyy" & VBA.IIf(VBA.Format$(RecFldValue, "HH:nn:ss AMPM") <> "12:00:00 AM", " HH:nn:ss AMPM", VBA.vbNullString))
                    
                End Select 'Close SELECT..CASE block statement
                
                'Format currency data types to 2 decimal places
                If LvItemDataType(vCol - &H1) = "C" Then RecFldValue = VBA.FormatNumber(VBA.CLng(VBA.Replace(RecFldValue, ",", VBA.vbNullString)), &H2)
                If (LvItemDataType(vCol - &H1) = "C" Or LvItemDataType(vCol - &H1) = "N") And VBA.Val(VBA.Replace(RecFldValue, ",", VBA.vbNullString)) < &H0 Then RecFldValue = "(" & VBA.Replace(VBA.Replace(RecFldValue, ",", VBA.vbNullString), "-", "") & ")"
                
                If vLv.ColumnHeaders.Count >= vCol Then If vLv.ColumnHeaders(vCol).Text = "Mean Score" Then RecFldValue = VBA.Format$(VBA.CLng(RecFldValue), "00.0000")
                If vLv.ColumnHeaders.Count >= vCol Then If vLv.ColumnHeaders(vCol).Text = "Avg Score" Or vLv.ColumnHeaders(vCol).Text = "Average Score" Then RecFldValue = VBA.Format$(VBA.CLng(RecFldValue), "00")
                If vLv.ColumnHeaders.Count >= vCol Then If VBA.InStr("|A|A-|B+|B|B-|C+|C|C-|D+|D|D-|E|X|Y|Z|", "|" & vLv.ColumnHeaders(vCol).Text & "|") <> &H0 Then RecFldValue = VBA.Format$(VBA.CLng(RecFldValue), "0")
                
                Lst.Text = VBA.Replace(RecFldValue, VBA.vbCrLf, " ~|~ ")
                If vCol = &H1 And nRs.State = adStateOpen Then If Not VBA.IsNull(nRs(&H0)) Then vLv.Parent.iSearchIDs = vLv.Parent.iSearchIDs & "|" & nRs(&H0).Value Else vLv.Parent.iSearchIDs = vLv.Parent.iSearchIDs & "|"
                
                Dim nNum&
                
                If (LvItemDataType(vCol - &H1) = "C" Or LvItemDataType(vCol - &H1) = "N") And VBA.IsNumeric(VBA.Replace(VBA.Replace(RecFldValue, "(", VBA.vbNullString), ")", VBA.vbNullString)) And VBA.InStr(RecFldValue, "(") <> &H0 And VBA.InStr(RecFldValue, ")") <> &H0 Then RecFldValue = "-" & VBA.Replace(VBA.Replace(RecFldValue, "(", VBA.vbNullString), ")", VBA.vbNullString)
                If vLv.ColumnHeaders.Count >= vCol Then If VBA.InStr("|" & iLastRowSumFields & "|", "|" & vLv.ColumnHeaders(vCol).Text & "|") <> &H0 Then iLastRowSumTotals(vCol - &H1) = iLastRowSumTotals(vCol - &H1) + VBA.IIf(Not IsNumeric(VBA.Val(VBA.Replace(RecFldValue, ",", VBA.vbNullString))), &H0, VBA.Val(VBA.Replace(RecFldValue, ",", VBA.vbNullString)))
                
                If vLv.ColumnHeaders.Count >= vCol Then If vLv.ColumnHeaders(vCol).Text = "Basic Salary" Then nNum = nNum + VBA.CLng(VBA.Replace(RecFldValue, ",", VBA.vbNullString))
                
                'Assign color to subitems
                Lst.ForeColor = vLv.ListItems(vLv.ListItems.Count).ForeColor
                
                iStatus = "Find Tag"
                
                'If the tag field has been specified then attach its value to the Tag property of the List Item
                If VBA.LenB(VBA.Trim$(iTagField)) <> &H0 Then
                    
                    nTagField = VBA.Split(VBA.Replace(iTagField, ";", "|"), "|")
                    
                    If vCol <= UBound(nTagField) + &H1 Then
                        If VBA.LenB(VBA.Trim$(nTagField(vCol - &H1))) <> &H0 Then
                            If Not VBA.IsNull(nRs(nTagField(vCol - &H1))) Then Lst.Tag = nRs(nTagField(vCol - &H1))
                        End If 'Close respective IF..THEN block statement
                    End If 'Close respective IF..THEN block statement
                    
                End If 'Close respective IF..THEN block statement
                
                If VBA.LenB(VBA.Trim$(iExcludeRecord)) <> &H0 Then
                    
                    Dim strChar$
                    Dim strFound As Boolean
                    
                    nArray = VBA.Split(iExcludeRecord, "|")
                    For iVal = &H0 To UBound(nArray) Step &H1
                        
                        strFound = False
                        
                        nArrayTmp = VBA.Split(nArray(iVal), ":")
                        
                        If nArrayTmp(&H0) = vLv.ColumnHeaders(vCol).Text Then
                            
                            If VBA.InStr(nArrayTmp(&H1), "=") <> &H0 Then
                                
                                Select Case VBA.Replace(nArrayTmp(&H1), "=", "")
                                    Case "''": strChar = " "
                                    Case Else: strChar = VBA.Replace(nArrayTmp(&H1), "=", "")
                                End Select
                                
                                If RecFldValue = strChar Then strFound = True: Exit For
                                
                            ElseIf VBA.InStr(nArrayTmp(&H1), "<>") <> &H0 Then
                                
                                Select Case VBA.Replace(nArrayTmp(&H1), "<>", "")
                                    Case "''": strChar = " "
                                    Case Else: strChar = VBA.Replace(nArrayTmp(&H1), "<>", "")
                                End Select
                                
                                If RecFldValue <> strChar Then strFound = True: Exit For
                                
                            End If
                            
                        End If
                        
                    Next iVal
                    
                    If strFound Then vLv.ListItems.Remove vRow: vRow = vRow - &H1: RecFldValue = VBA.vbNullString: GoTo iNextRow
                    
                End If 'Close respective IF..THEN block statement
NextCol:
                RecFldValue = VBA.vbNullString 'Initialize variable
                
            Next vCol 'Increment Column counter by 1
            
NextRow:
            
            Dim iNum&
            
            iStatus = "Find Tag"
            
            If vRow <= vLv.ListItems.Count Then vLv.ListItems(vRow).Tag = VBA.vbNullString
            
            If VBA.LenB(VBA.Trim$(iTagField)) <> &H0 And vRow <= vLv.ListItems.Count Then
                
                nTagField = VBA.Split(VBA.Replace(iTagField, ";", "|"), "|")
                
                For iNum = &H0 To UBound(nTagField) Step &H1
                    
                    If nTagField(iNum) <> VBA.vbNullString Then
                        If Not VBA.IsNull(nRs(nTagField(iNum))) Then vLv.ListItems(vRow).Tag = vLv.ListItems(vRow).Tag & "|" & nRs(nTagField(iNum))
                    Else
                        vLv.ListItems(vRow).Tag = vLv.ListItems(vRow).Tag & "|"
                    End If
                    
                Next iNum
                
                If vRow <= vLv.ListItems.Count Then If VBA.Left$(vLv.ListItems(vRow).Tag, &H1) = "|" Then vLv.ListItems(vRow).Tag = VBA.Right$(vLv.ListItems(vRow).Tag, VBA.Len(vLv.ListItems(vRow).Tag) - &H1)
                
            End If 'Close respective IF..THEN block statement
            
            If iToolTipText <> VBA.vbNullString Then If Not VBA.IsNull(vRs(iToolTipText)) Then vLv.ListItems(vRow).ToolTipText = vRs(iToolTipText)
            
            If vLv.Checkboxes Then vLv.ListItems(vRow).Checked = iCheckAll
            
            iStatus = VBA.vbNullString 'Initialize variable
iNextRow:
            .MoveNext 'Move to the next Record
            
            If nRs.State = adStateOpen Then nRs.MoveNext
            
            If ShowPleaseWait Then
                
                iCnt = iCnt + (Frm_PleaseWait.lblProgressBar.Width / iCounter)
                Frm_PleaseWait.ImgProgressBar.Width = iCnt
                Frm_PleaseWait.lblStatus.Caption = VBA.Int(((VBA.CLng(iCnt) * 100) / Frm_PleaseWait.lblProgressBar.Width)) & "% Complete"
                
            End If 'Close respective IF..THEN block statement
                        
            VBA.DoEvents 'Yield execution so that the operating system can process other events
            
        Next vRow 'Increment Row counter by 1
        
        .Close 'Close the opened object and any dependent objects
        
    End With 'Close the WITH block statements
    
ResizeColumns:
    
    iStatus = VBA.vbNullString  'Initialize variable
    
    If VBA.LenB(VBA.Trim$(iLastRowTotal)) <> &H0 And vLv.ListItems.Count > &H0 Then
        
        vLv.ListItems.Add vLv.ListItems.Count + &H1
        vLv.ListItems(vLv.ListItems.Count).Bold = True
        vLv.ListItems(vLv.ListItems.Count).Ghosted = True
        vLv.ListItems(vLv.ListItems.Count).ForeColor = &H80&
        
    End If 'Close respective IF..THEN block statement
    
    Dim iLst
    
    'For each column in the specified Listview...
    For vCol = &H1 To vLv.ColumnHeaders.Count Step &H1
        
        'Add the summation Row if specified
        If VBA.LenB(VBA.Trim$(iLastRowTotal)) <> &H0 And vLv.ListItems.Count > &H0 Then
            
            If vCol > &H1 Then vLv.ListItems(vLv.ListItems.Count).ListSubItems.Add (vCol - &H1), , " "
            If vCol = &H1 Then Set iLst = vLv.ListItems(vLv.ListItems.Count) Else Set iLst = vLv.ListItems(vLv.ListItems.Count).ListSubItems(vCol - &H1)
            iLst.Bold = True
            iLst.ForeColor = vLv.ListItems(vLv.ListItems.Count).ForeColor
            
        End If 'Close respective IF..THEN block statement
        
        If vLv.ListItems.Count > &H0 Then If vCol = &H1 Then Set iLst = vLv.ListItems(vLv.ListItems.Count) Else Set iLst = vLv.ListItems(vLv.ListItems.Count).ListSubItems(vCol - &H1)
        
        'Fill the summation Row with data
        If VBA.InStr("|" & iLastRowSumFields & "|", "|" & vLv.ColumnHeaders(vCol).Text & "|") <> &H0 And vLv.ListItems.Count > &H0 Then
            
            If vLv.ColumnHeaders(vCol).Text = "Mean Score" Then
                iLst.Text = VBA.Format$(VBA.CLng(VBA.Replace(iLastRowSumTotals(vCol - &H1), ",", VBA.vbNullString)) / (vLv.ListItems.Count - &H1), "00.0000")
            Else
                
                'If it's just a Currency then...
                If LvItemDataType(vCol - &H1) = "C" Then
                    iLst.Text = VBA.FormatNumber(VBA.CLng(VBA.Replace(iLastRowSumTotals(vCol - &H1), ",", VBA.vbNullString)), &H2)
                Else 'If it's just a number then...
                    If VBA.InStr("|A|A-|B+|B|B-|C+|C|C-|D+|D|D-|E|X|Y|Z|", "|" & vLv.ColumnHeaders(vCol).Text & "|") <> &H0 Then
                        iLst.Text = VBA.Format$(VBA.CLng(VBA.Replace(iLastRowSumTotals(vCol - &H1), ",", VBA.vbNullString)), "0")
                    Else
                        iLst.Text = VBA.Format$(VBA.CLng(VBA.Replace(iLastRowSumTotals(vCol - &H1), ",", VBA.vbNullString)), "00")
                    End If 'Close respective IF..THEN block statement
                End If 'Close respective IF..THEN block statement
                
            End If 'Close respective IF..THEN block statement
            
            vLv.ListItems(vLv.ListItems.Count).Tag = "Function" 'Denote it is a Col Function Row
            
        End If 'Close respective IF..THEN block statement
        
        'AutoFit Column to contents
        Call SendMessageB(vLv.hWnd, Lvm_SetColumnWidth, vCol - &H1, Lvm_AutoSize_UseHeader)
        
        If VBA.LenB(vLv.ColumnHeaders(vCol).Text) < 80 Then
            If VBA.CLng(strTxtWidth(vCol - &H1)) < VBA.CLng(vLv.ColumnHeaders(vCol).Width) - 50 And VBA.CLng(strTxtWidth(vCol - &H1)) <> &H0 Then strTxtWidth(vCol - &H1) = vLv.ColumnHeaders(vCol).Width - 50
        Else
            strTxtWidth(vCol - &H1) = 1000
        End If 'Close respective IF..THEN block statement
        
        vLv.ColumnHeaders(vCol).Width = VBA.IIf(strTxtWidth(vCol - &H1) = &H0, &H0, strTxtWidth(vCol - &H1) + 50)
        
        If strTxtWidth(vCol - &H1) = &H0 And vCol = &H1 And (Not Nothing Is vLv.SmallIcons Or vLv.Checkboxes) Then
            
            vLv.ColumnHeaders(vCol).Width = VBA.IIf(Not Nothing Is vLv.SmallIcons, 255, &H0) + VBA.IIf(vLv.Checkboxes, 255, &H0)
            vLv.ColumnHeaders(vCol).Text = "    " & vLv.ColumnHeaders(vCol).Text
            
        Else
            
            If vFirstVisibleCol <> &H0 And vLv.ColumnHeaders.Count = vFirstVisibleCol And vCol = vFirstVisibleCol Then
                
                'If there are Records retrieved from the database then...
                If vLv.ListItems.Count <> &H0 Then
                    
                    'Fit the whole column to the Listview width
                    vLv.ColumnHeaders(vFirstVisibleCol).Width = vLv.Width - VBA.IIf(Not Nothing Is vLv.SmallIcons, 255, &H0) - VBA.IIf(vLv.Checkboxes, 255, &H0) + VBA.IIf(vLv.Checkboxes, 255, &H0) - VBA.IIf((vLv.ListItems.Count * vLv.ListItems(&H1).Height) + vLv.ListItems(&H1).Height > vLv.Height, 255, &H0)
                    
                End If 'Close respective IF..THEN block statement
                
            Else
                
                'Resize the Column Width to fit the longest text in it
                vLv.ColumnHeaders(vCol).Width = VBA.IIf(VBA.CLng(strTxtWidth(vCol - &H1)) = &H0, &H0, strTxtWidth(vCol - &H1) + 50 + VBA.IIf(Not Nothing Is vLv.SmallIcons, 255, &H0) + VBA.IIf(vLv.Checkboxes, 255, &H0))
                
            End If 'Close respective IF..THEN block statement
            
        End If 'Close respective IF..THEN block statement
        
        If iAfterFillLv Then vLv.Parent.iMaxColWidths = vLv.Parent.iMaxColWidths & "|" & vLv.ColumnHeaders(vCol).Width
        
    Next vCol 'Increment the X variable by the value in the Step option
    
    'If the data starts with '|' character then remove it
    If iAfterFillLv Then If VBA.Left$(vLv.Parent.iMaxColWidths, &H1) = "|" Then vLv.Parent.iMaxColWidths = VBA.Right$(vLv.Parent.iMaxColWidths, VBA.Len(vLv.Parent.iMaxColWidths) - &H1)
    
Exit_FillListView:
    
    'If coloring condition has been specified then...
    If VBA.LenB(VBA.Trim$(iFieldCondition)) <> &H0 Then
        
        'Remove the loaded DTPickers
        vLv.Parent.Controls.Remove "DTP1": vLv.Parent.Controls.Remove "DTP2"
        
    End If 'Close respective IF..THEN block statement
    
    'Assign the total no of Records retrieved
    FillListView = VBA.IIf(vFirstVisibleCol = &H0, -&H1, vLv.ListItems.Count)
    
    'Alternate Lv background ground between light grey and white
    Call AltLvBackground(vLv, tTheme.tLvColorOne, tTheme.tLvColorTwo)
    
    'If the data starts with '|' character then remove it
    If VBA.Left$(vLv.Parent.iSearchIDs, &H1) = "|" Then vLv.Parent.iSearchIDs = VBA.Right$(vLv.Parent.iSearchIDs, VBA.Len(vLv.Parent.iSearchIDs) - &H1)
    
    If iAfterFillLv Then Call vLv.Parent.AfterFillLv
    
    vLv.Visible = LvState 'Display Listview
    vLv.Refresh
    
    vCancelOperation = &H0 'Initialize variable
    
    If ShowPleaseWait Then Unload Frm_PleaseWait
    
    'Initialize variable
    iStatus = VBA.vbNullString
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Function 'Quit this Function
    
Handle_FillListView_Error:
    
    'If tag field does not exist then resume execution at the next code
    If Err.Number = 3265 Then If iStatus = "Find Tag" Then Resume Next Else Resume NextCondition
    
    'Ignore the known errors
    If Err.Number = 727 Or Err.Number = 730 Or Err.Number = &H9 Then Resume Next
    
'    'Quit Function in case of an SQL error
'    If Err.Number = -2147217904 Or Err.Number = 424 Then Resume Exit_FillListView
    
    'Hinder it from displaying No Records message
    vFirstVisibleCol = &H0
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'If a Field Name in the specified Query does not exist then...
    If Err.Number = -2147217904 Then
        
        'Warn User, displaying the Full Query
        If vMsgBox(Err.Description & VBA.vbCrLf & VBA.vbCrLf & "{" & iSQL & "}" & VBA.vbCrLf & VBA.vbCrLf & "Do you want to Retry?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title & " : Record Retrieval Error - " & Err.Number, vLv.Parent) = vbYes Then Resume
        
    Else 'If it is a different error then...
        
        'Warn User
        vMsgBox Err.Description, vbExclamation, App.Title & " : Record Retrieval Error - " & Err.Number, vLv.Parent
        
    End If 'Close respective IF..THEN block statement
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Resume execution at the specified Label
    Resume Exit_FillListView
    
End Function

Public Function FitPicTo(SourceImg As Image, DestImg As Image, BoundObject As Object, Optional myPicOffSet& = 100, Optional myTreatAsFrame As Boolean = True) As Boolean
On Local Error GoTo Handle_FitPicTo_Error
    
    Dim MyImgWidth%, MyImgHeight%
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'If the Source Image's Stretch property is TRUE then Quit this procedure
    If SourceImg.Stretch = True Then GoTo Exit_FitPicTo
    
    'If the Destination Image's Stretch property is FALSE then...
    If DestImg.Stretch = False Then
        
        'Warn User
        vMsgBox "Invalid Destination Image. Stretching cancelled", vbExclamation, App.Title & " : Destination Img not Stretched", SourceImg.Parent
        GoTo Exit_FitPicTo 'Quit this procedure
        
    End If 'End respective IF..THEN block statement
    
    'Set Mouse pointer to indicate beginning of process or operation
    Screen.MousePointer = vbHourglass
    
    MyImgWidth = BoundObject.Width - myPicOffSet
    MyImgHeight = BoundObject.Height - myPicOffSet - VBA.IIf(TypeOf BoundObject Is Frame, 200, &H0)
    
    Dim DH&, DW&, OF&
    Dim SH&, SW&, FH&, FW&
    
    OF = 100
    SH = SourceImg.Height: SW = SourceImg.Width
    FH = BoundObject.Height - myPicOffSet - VBA.IIf(TypeOf BoundObject Is Frame And myTreatAsFrame, OF, &H0)
    FW = BoundObject.Width - myPicOffSet
    
    DH = SH: DW = SW
    
    If DH < FH And DW < FW Then DH = FH: DW = (SW * FH) / SH
    If DH > FH Then DH = FH: DW = (SW * FH) / SH
    If DW > FW Then DW = FW: DH = (SH * FW) / SW
    
    DestImg.Height = DH: DestImg.Width = DW
    
    'If the outline object is its container then...
    If BoundObject.Name = DestImg.Container.Name Then
        
        'Center DestImg in the Outline control
        DestImg.Top = (BoundObject.Height / &H2) - (DestImg.Height / &H2) + VBA.IIf(TypeOf BoundObject Is Frame And myTreatAsFrame, OF / &H2, &H0)
        DestImg.Left = (BoundObject.Width / &H2) - (DestImg.Width / &H2)
        
    Else 'If the outline object is not its container then...
        
        DestImg.Top = BoundObject.Top + ((BoundObject.Height / &H2) - (DestImg.Height / &H2))
        DestImg.Left = BoundObject.Left + ((BoundObject.Width / &H2) - (DestImg.Width / &H2))
        
    End If 'End respective IF..THEN block statement
    
    'Copy picture from source image control to destination image control
    DestImg.Picture = SourceImg.Picture: DestImg.ToolTipText = SourceImg.ToolTipText
    
Exit_FitPicTo:
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Function 'Quit this Function
    
Handle_FitPicTo_Error:
    
    'Too small Screen
    If Err.Number = 380 Then Resume Next
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Image Resize Error - " & Err.Number, SourceImg.Parent
    
    'Resume execution at the specified Label
    Resume Exit_FitPicTo
    
End Function

'Procedure to drag any object that has HWnd property
Public Function FormDrag(Obj As Object)
    
    Obj.MousePointer = vbSizeAll
    ReleaseCapture
    Call SendMessage(Obj.hWnd, &HA1, 2, 0&)
    Obj.MousePointer = vbDefault
    
End Function

Public Function GetGrade(ScoreValue&, Optional iSubjectName$, Optional ProfileID& = 1, Optional iComment As Boolean = False, Optional iAbbrev As Boolean = False) As String
    
    Dim GradeRs As New ADODB.Recordset
    
    If ProfileID = &H0 Then ProfileID = &H1
    
    Set GradeRs = New ADODB.Recordset
    GradeRs.Open "SELECT * FROM [Qry_Gradings] WHERE [Grade From] <= " & ScoreValue & " AND [Grade To] >= " & ScoreValue & VBA.IIf(ProfileID = &H1, " AND [Deletable] = False", " AND [Grading System ID] IN (SELECT [Grading System ID] FROM [Qry_Subjects] WHERE [" & VBA.IIf(iAbbrev, "Abbreviation", "Subject Name") & "] = '" & VBA.Replace(iSubjectName, "'", "''") & "')") & " ORDER BY [Grade To] DESC", vAdoCNN, adOpenKeyset, adLockReadOnly
    If GradeRs.RecordCount <> &H0 Then GetGrade = GradeRs![Grade Name] & VBA.IIf(iComment, "|" & GradeRs![Comment], VBA.vbNullString) & VBA.IIf(iComment, "|" & GradeRs![Grade Point], VBA.vbNullString)
    Set GradeRs = Nothing
    
End Function

Public Function GetSoftwareSettings(Optional iConnect As Boolean = True) As String
On Local Error GoTo Handle_GetSoftwareSettings_Error
    
    Dim MousePointerState%
    Dim sRs As New ADODB.Recordset
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    If iConnect Then ConnectDB , False 'Call Function in this Module to connect to Software's Database
     
    Set sRs = New ADODB.Recordset
    With sRs 'Execute a series of statements on sRs recordset
        
        .Open "SELECT * FROM `Tbl_Setup` ORDER BY `Date Created` DESC", vAdoCNN, adOpenKeyset, adLockReadOnly
        
        'If there are records in the table then...
        If Not (.BOF And .EOF) Then
            
            If Not VBA.IsNull(![Max LoginRecords]) Then
                SoftwareSetting.Max_Login_Records = ![Max LoginRecords]
            Else
                SoftwareSetting.Max_Login_Records = 500
            End If
            
            If Not VBA.IsNull(![Max ApplicationLogs]) Then
                SoftwareSetting.Max_App_Logs = ![Max ApplicationLogs]
            Else
                SoftwareSetting.Max_App_Logs = 4000
            End If
            
            If Not VBA.IsNull(![Splash Screen]) Then
                SoftwareSetting.Show_Splash_Screen = ![Splash Screen]
            Else
                SoftwareSetting.Show_Splash_Screen = True
            End If
            
            If Not VBA.IsNull(![Request Alternative Login]) Then
                SoftwareSetting.RequestAlternativeLogin = ![Request Alternative Login]
            Else
                SoftwareSetting.RequestAlternativeLogin = True
            End If
            
            If Not VBA.IsNull(![MinUserPwdCharacters]) Then
                SoftwareSetting.Min_User_Password_Characters = ![MinUserPwdCharacters]
            Else
                SoftwareSetting.Min_User_Password_Characters = &H0
            End If
            
            If Not VBA.IsNull(![MinUsernameCharacters]) Then
                SoftwareSetting.Min_Username_Characters = ![MinUsernameCharacters]
            Else
                SoftwareSetting.Min_Username_Characters = &H0
            End If
            
            If Not VBA.IsNull(![Term]) Then
                SoftwareSetting.Term = ![Term]
            Else
                SoftwareSetting.Term = &H1
            End If
            
            If Not VBA.IsNull(![Academic Year]) Then
                SoftwareSetting.Academic_Year = ![Academic Year]
            Else
                SoftwareSetting.Academic_Year = VBA.Date
            End If
            
            If Not VBA.IsNull(![Min Year]) Then
                SoftwareSetting.Min_Year = ![Min Year]
            Else
                SoftwareSetting.Min_Year = VBA.Year(VBA.Date)
            End If
            
        Else
            
            SoftwareSetting.Max_Login_Records = 500 ' VBA.CLng(Db.TableDefs(.Fields("Max LoginRecords").Properties(1).Value).Fields("Max LoginRecords").DefaultValue)
            SoftwareSetting.Max_App_Logs = 4000 ' VBA.CLng(Db.TableDefs(.Fields("Max ApplicationLogs").Properties(1).Value).Fields("Max ApplicationLogs").DefaultValue)
            SoftwareSetting.Min_User_Password_Characters = &H0 ' VBA.CLng(Db.TableDefs(.Fields("MinUserPwdCharacters").Properties(1).Value).Fields("MinUserPwdCharacters").DefaultValue)
            SoftwareSetting.Min_Username_Characters = &H0 ' VBA.CLng(Db.TableDefs(.Fields("MinUsernameCharacters").Properties(1).Value).Fields("MinUsernameCharacters").DefaultValue)
            SoftwareSetting.Academic_Year = VBA.Date
            SoftwareSetting.Term = &H1
            SoftwareSetting.Min_Year = VBA.Year(VBA.Date)
            SoftwareSetting.Show_Splash_Screen = True ' (VBA.InStr("yestrueon", VBA.LCase$(Db.TableDefs(.Fields("Splash Screen").Properties(1).Value).Fields("Splash Screen").DefaultValue)) <> &H0)
            SoftwareSetting.RequestAlternativeLogin = True ' (VBA.InStr("yestrueon", VBA.LCase$(Db.TableDefs(.Fields("Request Alternative Login").Properties(1).Value).Fields("Request Alternative Login").DefaultValue)) <> &H0)
            
        End If 'Close respective IF..THEN block statement
        
        .Close 'Close the opened object and any dependent objects
        
    End With 'Close the WITH block statements
    
Exit_GetSoftwareSettings:
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Function 'Quit this Function
    
Handle_GetSoftwareSettings_Error:
    
    'Too small Screen
    If Err.Number = 380 Then Resume Next
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error Retrieving Software Settings - " & Err.Number
    
    'Resume execution at the specified Label
    Resume Exit_GetSoftwareSettings
    
End Function

Public Function GetWindowsDir() As String
    
    Dim Temp As String
    Dim ret As Long
    
    Const MAX_LENGTH = 145

    Temp = VBA.String$(MAX_LENGTH, &H0)
    ret = GetWindowsDirectory(Temp, MAX_LENGTH)
    Temp = VBA.Left$(Temp, ret)
    GetWindowsDir = VBA.IIf(Temp <> VBA.vbNullString And VBA.Right$(Temp, &H1) <> "\", Temp & "\", Temp)
    
End Function

Public Function LoadReport(Frm As Form, Rpt As Object, RptRs As ADODB.Recordset, Optional PrintDirect As Boolean = False, Optional RptCondition$ = VBA.vbNullString, Optional RefreshRs As Boolean = True) As Boolean
On Local Error GoTo Handle_LoadReport_Error
    
    'If the target report has not been specified then quit this procedure
    If Nothing Is Rpt Then Exit Function
    
    Dim MousePointerState%
    Dim InitialRptRecordsetQry$
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'If the Data Environment has to be refreshed then...
    If RefreshRs Then
        
        '---These codes enable the Reports and the Software to share the same database
        
        ConnectDB , False 'Call Procedure to update the current database location
'        If DEnv_Xpros.RptCNN.State = adStateClosed Then DEnv_Xpros.RptCNN.Open 'Open Report connection
        
        'Retrieve Recordset's initial value
        InitialRptRecordsetQry = RptRs.Source
        
        'If the Report's Recordset is open then close it
        If RptRs.State = adStateOpen Then RptRs.Close
        
        'Re-Open the Report's Recordset with the new SQL string
        RptRs.Open ' InitialRptRecordsetQry & " " & RptCondition, vAdoCNN, adOpenKeyset, adLockReadOnly
        
    Else 'If the Data Environment should not be refreshed then...
        
        On Error Resume Next
        
        'If a Condition has been specified then the Report's Recordset has to filter
        'records that meet the condition
        RptRs.Filter = VBA.Trim$(VBA.Replace(RptCondition, "WHERE", ""))
        
    End If 'Close respective IF..THEN block statement
    
    Screen.MousePointer = vbDefault 'Change Mouse Pointer to show end of Processing state
    
    'If the report should not be sent directly to the Printer then display it else send it to the Printer
    If Not PrintDirect Then Rpt.Show vbModal Else Rpt.PrintReport True
    
    'If the Report's Recordset is open then close it
    If RptRs.State = adStateOpen Then RptRs.Close
    
    'Reset Recordset to its initial value
    RptRs.Source = InitialRptRecordsetQry
    
    Set Rpt = Nothing 'Free object from Memory
    
Exit_LoadReport:
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Function 'Quit this Procedure
    
Handle_LoadReport_Error:
    
    'If type mismatch error then execute the next line
    If Err.Number = &HD Then Resume Next
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error Loading Report - " & Err.Number, Frm
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Resume execution at the specified Label
    Resume Exit_LoadReport
    
End Function

Public Function LoadUnloadedFormViaStringName(mFrm As Form, mFormName$, Optional mFrmDefinitions$ = VBA.vbNullString, Optional mFrmEntryDefn$ = VBA.vbNullString, Optional mLoadForm As Boolean = True, Optional mLoadFormAsInstance As Boolean = False, Optional iRecordID&) As Form
On Local Error GoTo Handle_LoadUnloadedFormViaStringName_Error
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    Dim iFrm As Form
    
    'Check if the specified Form has already been loaded
    For Each iFrm In VB.Forms
        If iFrm.Name = mFormName Then Exit For
    Next iFrm
    
    'If the specified Form has not been loaded then...
    If Nothing Is iFrm Then
        
        'Add it to the collection of loaded Forms and get it
        Set iFrm = Forms.Add(mFormName)
        
    Else
        
        'Add it to the collection of loaded Forms and get it
        Set iFrm = Forms.Add(mFormName)
        
    End If 'Close respective IF..THEN block statement
    
    Set LoadUnloadedFormViaStringName = iFrm
    
    'If the Form should not be loaded then...
    If Not mLoadForm Then
        
        'Unload it in order to assign values to its public variables
        vSilentClosure = True: Unload iFrm: vSilentClosure = False
        
        GoTo Exit_LoadUnloadedFormViaStringName
        
    End If 'Close respective IF..THEN block statement
    
    'If the Form's public variables have to be assigned values then...
    If VBA.LenB(VBA.Trim$(mFrmDefinitions)) <> &H0 Then
        
        'Unload it in order to assign values to its public variables
        If Not mLoadFormAsInstance Then vSilentClosure = True: Unload iFrm: vSilentClosure = False
        
        'Assign values to its public variables
        iFrm.FrmDefinitions = VBA.Replace(VBA.Replace(VBA.Replace(mFrmDefinitions, "~", ":"), "-", "|"), ",", "~")
        If VBA.Val(mFrmEntryDefn) > &H0 Then iFrm.FrmEntryDefn = VBA.Val(mFrmEntryDefn)
        
    End If 'Close respective IF..THEN block statement
    
    If VBA.CLng(iRecordID) > &H0 Then iFrm.iRecordID = VBA.CLng(iRecordID)
        
    Dim nFrm As Form
    Dim nFrmName As String
    
    nFrmName = iFrm.Name
    iFrm.Show vbModal, mFrm 'Display the Form to the User
    
    'For each loaded Form...
    For Each nFrm In VB.Forms
        
        'If one of them is the closed one then unload it
        If nFrm.Name = nFrmName Then vSilentClosure = True: Unload nFrm
        
    Next nFrm 'Move to the next open Form
    
Exit_LoadUnloadedFormViaStringName:
    
    vSilentClosure = False
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Function 'Quit this Sub-Procedure
    
Handle_LoadUnloadedFormViaStringName_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error Loading Form - " & Err.Number, mFrm
    
    'Resume execution at the specified Label
    Resume Exit_LoadUnloadedFormViaStringName
    
End Function

Public Function NavigateToRec(iFrm As Form, iTableName$, iPryKeyFld$, iDirection%, iRecordNo&, Optional CheckMode As Boolean = True) As Long
On Local Error GoTo Handle_NavigateToRec_Error
    
    'If the User wants to continue Modifying a record then Quit this Procedure
    If CheckMode Then If ProceedEditting(iFrm) = True Then Exit Function
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    If VBA.LenB(VBA.Trim$(vAdoCNN.ConnectionString)) = &H0 Then Call ConnectDB  'Call Function in this Module to connect to Software's Database
    Set vRs = New ADODB.Recordset
    With vRs 'Execute a series of statements on vRs recordset
        
        .Open iTableName, vAdoCNN, adOpenKeyset, adLockReadOnly
        
        If Not (.BOF And .EOF) Then 'If there are records in the table then...
            
            Select Case iDirection 'Check value of CmdMoveRec Index(0)
                
                Case &H0: iRecordNo = &H1 'If CmdMoveRec Index(0) value is 0 then assign 1 to iRecordNo variable
                Case &H1: iRecordNo = iRecordNo - &H1 'If CmdMoveRec Index(0) value is 1 then add 1 to the value in iRecordNo variable
                Case &H2: iRecordNo = iRecordNo + &H1 'If CmdMoveRec Index(0) value is 2 then subtract 1 to the value in iRecordNo variable
                Case &H3: iRecordNo = .RecordCount 'If CmdMoveRec Index(0) value is 3 then
                
            End Select 'End Checking Index(0) value
            
            Select Case iRecordNo 'Check value of iRecordNo Variable
                
                Case Is < &H1: 'If iRecordNo Variable value is less than 1 then inform user that it's the First Record
                    
                    'Indicate that a process or operation is complete.
                    Screen.MousePointer = vbDefault
                    
                    .MoveFirst 'Select Record in position same as that of value in iRecordNo variable subtracted 1
                    
                    If iRecordNo = -&H1 Then
                        iFrm.DisplayRecord VBA.CLng(vRs(iPryKeyFld)) 'Call procedure in the specified Form
                    End If 'Close respective IF..THEN block statement
                    
                    iRecordNo = &H1 'Assign 1 to iRecordNo variable
                    
                    'Inform User
                    vMsgBox "This is the First Record", vbInformation, App.Title & " : Navigation", iFrm
                    
                Case Is > .RecordCount: 'If iRecordNo Variable value is greater than number of records then inform user that it's the Last Record
                    
                    'Indicate that a process or operation is complete.
                    Screen.MousePointer = vbDefault
                    
                    .MoveLast 'Select Record in position same as that of value in iRecordNo variable subtracted 1
                    
                    iRecordNo = .RecordCount 'Assign total number of records to iRecordNo variable
                    
                    'Inform User
                    vMsgBox " This is the Last Record", vbInformation, App.Title & " : Record Navigation", iFrm
                    
                Case Else
                    
                    .Move iRecordNo - &H1 'Select Record in position same as that of value in iRecordNo variable subtracted 1
                    iFrm.DisplayRecord VBA.CLng(vRs(iPryKeyFld)) 'Call procedure in the specified Form
                    
            End Select 'End Checking iRecordNo value
            
        Else: 'If there are on records in the table
            
            'Indicate that a process or operation is complete.
            Screen.MousePointer = vbDefault
            
            iFrm.ClearEntries 'Call procedure in the specified Form
            
            'If there are no records in the table Inform User
            vMsgBox "There are no records to Display", vbInformation, App.Title & " : Record Navigation", iFrm
            
        End If 'Close respective IF..THEN block statement
        
        If .State = adStateOpen Then .Close 'Close the opened object and any dependent objects
        
    End With 'Close the WITH block statements
    
Exit_NavigateToRec:
    
    NavigateToRec = iRecordNo 'Assign the current record position
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Function 'Quit this Function
    
Handle_NavigateToRec_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Navigation Error - " & Err.Number, iFrm
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Resume execution at the specified Label
    Resume Exit_NavigateToRec
    
End Function

Public Function OpenNotes(vShpBttnObj As Object, Optional vLocked As Boolean = False, Optional vTitle$ = VBA.vbNullString, Optional vCustomMsg$ = VBA.vbNullString, Optional vHeading$ = VBA.vbNullString) As String
On Local Error GoTo Handle_OpenNotes_Error
    
    Dim vData$
    
    If TypeOf vShpBttnObj Is ShapeButton Then vData = vShpBttnObj.TagExtra Else vData = vShpBttnObj.Tag
    
    'If no Notes were previously saved then...
    If VBA.LenB(VBA.Trim$(vData)) = &H0 And vLocked Then
        
        'Inform User
        vMsgBox VBA.IIf(VBA.LenB(VBA.Trim$(vCustomMsg)) <> &H0, vCustomMsg, "No Saved Notes for this " & vTitle), vbInformation, App.Title & " : " & vTitle, vShpBttnObj.Parent
        Exit Function 'Quit this Procedure
        
    End If 'Close respective IF..THEN block statement
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    Load Frm_BriefNotes 'Load the Form onto the Memory
    
    With Frm_BriefNotes
        
        'Lock/UnLock the Form's Notes textbox
        .txtBriefNotes.Locked = vLocked
        .txtBriefNotes.Text = vData 'Assign the saved entry to the Form's Notes textbox
        .Caption = App.Title & " : " & vTitle
        .lblHeading.Caption = vHeading
        
        vBuffer(&H0) = .txtBriefNotes.Text 'Assign default
        
        CenterForm Frm_BriefNotes, vShpBttnObj.Parent 'Display the Form to the User
        
        vData = vBuffer(&H0)  'Assign the entry in the Form's Notes textbox
        OpenNotes = vBuffer(&H0)  'Assign the entry in the Form's Notes textbox
        vBuffer(&H0) = VBA.vbNullString  'Initialize variable
        
        If TypeOf vShpBttnObj Is ShapeButton Then vShpBttnObj.TagExtra = vData Else vShpBttnObj.Tag = vData
        
    End With 'End WITH statement
    
Exit_OpenNotes:
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Function 'Quit this Function
    
Handle_OpenNotes_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error Opening Note - " & Err.Number, vShpBttnObj.Parent
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Resume execution at the specified Label
    Resume Exit_OpenNotes
    
End Function

Public Function OpenFile(vFrm As Form, vFilePath$, Optional vIsURL As Boolean = False) As String
On Local Error GoTo Handle_OpenFile_Error
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Open the specified file
    ShellExecute vFrm.hWnd, "open", vFilePath, VBA.vbNullString, VBA.CStr(vFso.GetDrive(vFso.GetDriveName(GetWindowsDir))), ByVal 1&
    
Exit_OpenFile:
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Function 'Quit this Function
    
Handle_OpenFile_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error Opening URL - " & Err.Number, vFrm
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Resume execution at the specified Label
    Resume Exit_OpenFile
    
End Function

Public Function PerformMemoryCleanup()
On Local Error GoTo Handle_PerformMemoryCleanup_Error
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'If vRs has been opened then close it
    If vRs.State = adStateOpen Then vRs.Close
    
    'If vRsTmp has been opened then close it
    If vRsTmp.State = adStateOpen Then vRsTmp.Close
    
    'If vAdoCNN has been opened then close it
    If vAdoCNN.State = adStateOpen Then vAdoCNN.Close
    
    'Disassociate ADODB object variables from their actual objects
    Set vRs = Nothing:  Set vRsTmp = Nothing: Set vAdoCNN = Nothing
    
Exit_PerformMemoryCleanup:
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Function 'Quit this Function
    
Handle_PerformMemoryCleanup_Error:
    
    'If known error then execute the next line of code
    If Err.Number = 3219 Then Resume Next
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Navigation Error - " & Err.Number
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Resume execution at the specified Label
    Resume Exit_PerformMemoryCleanup
    
End Function

Public Function PhotoClicked(Img As Image, VirtualImg As Image, ObjOutline As Object, Optional Locked As Boolean = False, Optional DirectlyExecMenu$ = VBA.vbNullString) As Boolean
On Local Error GoTo Handle_PhotoClicked_Error
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Executes a series of statements for Frm_PopupMenus Form
    With Frm_PhotoZoom
        
        If VBA.LenB(VBA.Trim$(DirectlyExecMenu)) <> &H0 And Img.Picture <> &H0 Then GoTo DefineObjects
        
        .MnuPhoto(&H0).Visible = True: .MnuPhoto(&H1).Visible = False: .MnuPhoto(&H2).Visible = False
        
        'If there is an image displayed then...
        If Img.Picture <> &H0 Then
            
            'Display appropriate Menus
            .MnuPhoto(&H0).Caption = "&Replace Photo"
            .MnuPhoto(&H1).Visible = True: .MnuPhoto(&H2).Visible = True
            
        Else 'If there is no image displayed then...
            
            'If the Record is not in Edit mode then quit this procedure, else allow User to attach a photo
            If Locked Then GoTo Exit_PhotoClicked Else .MnuPhoto(&H0).Caption = "&Attach Photo"
            
            'Display appropriate Menus
            .MnuPhoto(&H3).Visible = False: .MnuPhoto(&H4).Visible = False
            .MnuPhoto(&H5).Visible = False: .MnuPhoto(&H6).Visible = False
            
        End If 'Close respective IF..THEN block statement
        
        'If the Record is not in Edit mode then on show the Maximize Menu
        If Locked Then .MnuPhoto(&H2).Visible = True: .MnuPhoto(&H0).Visible = False: .MnuPhoto(&H1).Visible = False
        
DefineObjects:
        
        Set .DestImg = Img 'Assign current image control
        Set .SourceImg = VirtualImg 'Assign virtual image control
        Set .DestImgOutline = ObjOutline  'Assign image control outline
        
        'If execute the click event of the Menu rather than wait for User to respond then...
        If VBA.LenB(VBA.Trim$(DirectlyExecMenu)) <> &H0 Then
            
            'Execute the click event
            Call .MnuPhoto_Click(VBA.IIf(Locked, &H2, VBA.CLng(DirectlyExecMenu)))
            
        Else 'If more than one Menu should be displayed for selection then...
            
            'Indicate that a process or operation is complete.
            Screen.MousePointer = vbDefault
            
            'Display specified Form's MnuPhotos Menu as this Form's Popup Menu
            Img.Parent.PopupMenu .MnuPhotos:
            
        End If 'Close respective IF..THEN block statement
        
        'Indicate that a process or operation is in progress.
        Screen.MousePointer = vbHourglass
        
    End With 'End WITH statement
    
    PhotoClicked = True
    
Exit_PhotoClicked:
    
    Unload Frm_PhotoZoom
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Function 'Quit this Function
    
Handle_PhotoClicked_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Image Click Error - " & Err.Number, Frm_PhotoZoom.SourceImg.Parent
    
    'Resume execution at the specified Label
    Resume Exit_PhotoClicked
    
End Function

'This Procedure enables a User to extract a text FROM a table to a destination control
Public Function PickDetails(iFrm As Form, iTableName$, Optional InitEntry$ = VBA.vbNullString, Optional iTargetFldNo$ = "1", Optional iHideFlds$ = VBA.vbNullString, Optional iExcludeFlds$ = VBA.vbNullString, Optional iMultiselect As Boolean = False, Optional iTitle$, Optional iSelColPos& = &H0, Optional iPhotoSpecifications$, Optional NoSelection As Boolean = False, Optional iFrmWidth& = &H0, Optional iNoRecordCustomMsg$ = VBA.vbNullString) As String
On Local Error GoTo Handle_PickDetails_Error
    
'    VB.Clipboard.Clear: VB.Clipboard.SetText iTableName
    
    Dim iTotalRecords&
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    vFrmWidth = iFrmWidth
    
    Load Frm_PickDetails 'Load the specified Form into the memory
    
    'Execute a series of statements for the specified Form
    With Frm_PickDetails
        
        'Assign title if specified
        .lblSelectedCategory.Caption = iTitle
        
        .Lv.Checkboxes = iMultiselect
        .ShpBttnCheck(&H0).Visible = .Lv.Checkboxes
        .ShpBttnCheck(&H1).Visible = .Lv.Checkboxes
        
        .iTargetFields = iTargetFldNo
        .iPhotoSpecifications = iPhotoSpecifications
        .iSelColPos = iSelColPos
        .ShpBttnOK.Visible = Not NoSelection
        .ShpBttnCancel.Visible = Not NoSelection
        
        'Get the total no of Records retrieved from the database
        iTotalRecords = FillListView(.Lv, iTableName, iHideFlds, iExcludeFlds)
        
        'If there are no Records then...
        If iTotalRecords <= &H0 Then
            
            'If there are no Records then...
            If iTotalRecords = &H0 Then
                
                'Indicate that a process or operation is complete.
                Screen.MousePointer = vbDefault
                
                'Inform User
                vMsgBox VBA.IIf(iNoRecordCustomMsg <> VBA.vbNullString, iNoRecordCustomMsg, "There are no Records to be displayed."), vbInformation, App.Title & " : No Records", iFrm
                
            End If 'Close respective IF..THEN block statement
            
            'Unload the Form from the memory
            Unload Frm_PickDetails
            
            GoTo Exit_PickDetails 'Quit displaying the Form
            
        End If 'Close respective IF..THEN block statement
        
        'Display the Total no of Records retrieved
        .lblTotalRecords.Caption = "Total Records: " & iTotalRecords
        
        'If an initial selection had been made then...
        If VBA.LenB(VBA.Trim$(InitEntry)) <> &H0 Then
            
            Dim nCnt&
            Dim nArray() As String
            Dim Lst As ListItem
            
            nArray = VBA.Split(InitEntry, ";")
            
            Set .Lv.SelectedItem = Nothing
                
            For nCnt = &H0 To UBound(nArray) Step &H1
                
                'Search for the value in List Items
                Set Lst = .Lv.FindItem(nArray(nCnt), &H0)
                
                'If the item is not found then search in Sub Items
                If Nothing Is Lst Then Set Lst = .Lv.FindItem(nArray(nCnt), &H1)
                
                'If the item is found then
                If Not Nothing Is Lst Then
                    
                    If .Lv.Checkboxes Then .Lv.ListItems(Lst.Index).Checked = True 'Check it,
                    .Lv.ListItems(Lst.Index).Selected = True 'Highlight it and...
                    If nCnt = UBound(nArray) Then Lst.EnsureVisible 'Make it visible to the User (Scrolls the Listview till it's visible)
                    
                End If 'Close respective IF..THEN block statement
                
            Next nCnt
            
        End If 'Close respective IF..THEN block statement
        
        'Call Function to load the specified Form
        CenterForm Frm_PickDetails, iFrm
        
        'Get the details of the first Record selected
        PickDetails = vMultiSelectedData
        
    End With 'End WITH block statements
    
Exit_PickDetails:
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Function 'Quit this Function
    
Handle_PickDetails_Error:
    
    'If known error, assume it
    If Err.Number = 91 Or Err.Number = &H9 Then Resume Next
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error Selecting Details - " & Err.Number, iFrm
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Resume execution at the specified Label
    Resume Exit_PickDetails
    
End Function

Public Function PlaySound(iFrm As Form, iSoundFilePath$) As Boolean
On Local Error GoTo Handle_PlaySound_Error
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'If the identified sound does not exist then Quit this Function
    If Not vFso.FileExists(iSoundFilePath) Then GoTo Exit_PlaySound
    
    Dim wFlags%, mVar%
    Dim mHasSoundCard As Boolean
    
    'Check if Sound Card is installed and enabled
    mHasSoundCard = (waveOutGetNumDevs() > &H0)
    
    'If it has no Sound Card installed then Quit this Function
    If Not mHasSoundCard Then GoTo Exit_PlaySound
    
    wFlags = SND_ASYNC Or SND_NODEFAULT
    
    'Play the sound
    mVar = sndPlaySound(iSoundFilePath, wFlags)
    
    PlaySound = True
    
Exit_PlaySound:
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Function 'Quit this Function
    
Handle_PlaySound_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error Playing Sound - " & Err.Number, iFrm
    
    'Resume execution at the specified Label
    Resume Exit_PlaySound
    
End Function

'Check if the User is Modifying a record and returns User Choice
Public Function ProceedEditting(iFrm As Form, Optional InputForm As Boolean = True) As Boolean
On Local Error GoTo Handle_ProceedEditting_Error
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    ProceedEditting = True 'Denote that the User wants to proceed with editting the Record
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'If the User was Modifying a record then...
    If Not iFrm.IsNewRecord Then
        
        'Confirm if User wants to proceed with Editting the Record, if Yes then quit this procedure
        If vMsgBox("Do you want to cancel the on-going Edit operation?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title & " : Confirming Mode") = vbNo Then GoTo Exit_ProceedEditting Else iFrm.IsNewRecord = True
        
        'Indicate that a process or operation is in progress.
        Screen.MousePointer = vbHourglass
        
    End If 'Close respective IF..THEN block statement
    
    ProceedEditting = False 'Denote that the User doesn't want to proceed with editting the Record
    
Exit_ProceedEditting:
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Function 'Quit this Function
    
Handle_ProceedEditting_Error:
    
    'If the Form is not an entry Form then resume execution at the specified Label
    If Err.Number = 438 Then ProceedEditting = False: Resume Exit_ProceedEditting
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Edit Mode Confirmation Error - " & Err.Number, iFrm
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Resume execution at the specified Label
    Resume Exit_ProceedEditting
    
End Function

Public Function ReFormattedDate(iDateValue$, Optional iDateFormat$ = "dd MMM yyyy", Optional iAttachDateFormat As Boolean = True) As String
On Local Error GoTo Handle_ReFormattedDate_Error
    
    If iDateValue = VBA.vbNullString Then Exit Function
    
    'A date of this format "Tue 11-Aug 2009 04:35:38 AM" with the weekday as string is not recognizable by the
    'Date Functions. Therefore we need to remove the "Tue" part
    '=================================================================================================
    
    Dim nCnt&
    Dim nDateArray() As String
    Dim nSeparatorArray() As String
    Dim nAsc$, nDate$, nMonths$, nWeekdays$, nNewTime$, nFormat$, iFormat$, nSeparator$
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    nDate = iDateValue ' "Tue 11-Aug 2009 04:35:38 AM"
    
    nWeekdays = "Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday"
    nMonths = "January|February|March|April|May|June|July|August|September|October|November|December"
    
    Dim idate$
    
    idate = VBA.vbNullString
    
    For nCnt = &H1 To VBA.LenB(nDate)
        
        nAsc = VBA.Asc(VBA.UCase(VBA.Mid$(nDate, nCnt, &H1)))
        
        If VBA.IsDate(VBA.Mid$(nDate, nCnt, VBA.LenB(nDate) - nCnt)) Then
            
            If nNewTime = VBA.vbNullString Then nNewTime = VBA.Format$(VBA.Mid$(nDate, nCnt, VBA.LenB(nDate) - nCnt + &H1), "hh:nn:ss AMPM")
            If VBA.CDate(VBA.Mid$(nDate, nCnt, VBA.LenB(nDate) - nCnt + &H1)) = nNewTime Then nSeparator = nSeparator & " ": Exit For
            
        End If 'Close respective IF..THEN block statement
        
        If Not (nAsc >= vbKeyA And nAsc <= vbKeyZ) And Not (nAsc >= vbKey0 And nAsc <= vbKey9) Then
            nSeparator = nSeparator & VBA.Chr$(nAsc) & "|": idate = idate & VBA.IIf(VBA.Chr$(nAsc) = ":", ":", " ")
        Else
            idate = idate & VBA.Mid$(nDate, nCnt, &H1)
        End If 'Close respective IF..THEN block statement
        
    Next nCnt
    
    nDate = idate & VBA.IIf(nNewTime <> VBA.vbNullString, " " & nNewTime, VBA.vbNullString)
    
    If VBA.Right$(nSeparator, &H1) = "|" Then nSeparator = VBA.Left$(nSeparator, VBA.Len(nSeparator) - &H1)
    
    nSeparatorArray = VBA.Split(nSeparator, "|")
    
    nDateArray = VBA.Split(iDateValue, " ")
    
    nFormat = VBA.vbNullString
    
    For nCnt = &H0 To UBound(nDateArray)
        
        If Not VBA.IsNumeric(nDateArray(nCnt)) Then
            
            If VBA.InStr(nWeekdays, nDateArray(nCnt)) <> &H0 Then
                iFormat = VBA.String$(VBA.LenB(nDateArray(nCnt)), "d"): nDateArray(nCnt) = VBA.vbNullString
            Else
                If VBA.InStr(nMonths, nDateArray(nCnt)) <> &H0 Then
                    iFormat = VBA.String$(VBA.LenB(nDateArray(nCnt)), "M")
                Else
                    If VBA.InStr(nDateArray(nCnt), ":") <> &H0 Then iFormat = "hh:nn:ss" Else iFormat = VBA.IIf(VBA.InStr(nDateArray(nCnt), "AM") <> &H0 Or VBA.InStr(nDateArray(nCnt), "PM") <> &H0, " AMPM", VBA.vbNullString)
                End If 'Close respective IF..THEN block statement
            End If 'Close respective IF..THEN block statement
            
        Else
            
            If VBA.IsNumeric(nDateArray(nCnt)) And VBA.CLng(nDateArray(nCnt)) > &HC And VBA.CLng(nDateArray(nCnt)) <= 31 And VBA.InStr(VBA.UCase$(nFormat), "M") <> &H0 Then
                iFormat = VBA.String$(VBA.LenB(nDateArray(nCnt)), "M")
            Else
                If VBA.IsNumeric(nDateArray(nCnt)) And VBA.CLng(nDateArray(nCnt)) > 1000 Then
                    iFormat = VBA.String$(VBA.LenB(nDateArray(nCnt)), "y")
                Else
                    iFormat = VBA.String$(VBA.LenB(nDateArray(nCnt)), "d")
                End If 'Close respective IF..THEN block statement
            End If 'Close respective IF..THEN block statement
            
        End If 'Close respective IF..THEN block statement
        
        nFormat = nFormat & iFormat
        
        If VBA.InStr(iFormat, ":") = &H0 And nCnt <= UBound(nSeparatorArray) Then nFormat = nFormat & nSeparatorArray(nCnt)
        
    Next nCnt
    
    ReFormattedDate = VBA.FormatDateTime(VBA.IIf(VBA.IsDate(VBA.Trim$(VBA.Join(nDateArray, " "))), VBA.Trim$(VBA.Join(nDateArray, " ")), VBA.vbNullString), vbShortDate)
    
    ReFormattedDate = VBA.IIf(ReFormattedDate <> VBA.vbNullString, ReFormattedDate & VBA.IIf(iAttachDateFormat, "|" & VBA.Trim$(nFormat), VBA.vbNullString), VBA.vbNullString)
    
    '=================================================================================================
    
Exit_ReFormattedDate:
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Function 'Quit this Function
    
Handle_ReFormattedDate_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Date Formatting Error - " & Err.Number
    
    'Resume execution at the specified Label
    Resume Exit_ReFormattedDate
    
End Function

Public Function RefreshWindows() As Boolean
    
    Const SHCNE_ASSOCCHANGED = &H8000000
    Const SHCNF_FLUSH = &H1000
    
    SHChangeNotify SHCNE_ASSOCCHANGED, SHCNF_FLUSH, &H0, &H0
    
End Function

Public Function SearchCboLst(sObj As Object, sSearchstr$) As Long
On Local Error GoTo Exit_SearchCboLst
    
    SearchCboLst = -&H1 'Assign default value
    SearchCboLst = SendMessage(sObj.hWnd, VBA.IIf(TypeOf sObj Is ComboBox, &H14C, &H18F), -&H1, ByVal CStr(sSearchstr))
    
Exit_SearchCboLst:
    
End Function

Public Function SearchLv(vLv As ListView, Optional vEnableFilterOption As Boolean = True, Optional vSelCol&) As Boolean
On Local Error GoTo Handle_SearchLv_Error
    
    'If the specified Listview does not have items the
    If vLv.ListItems.Count = &H0 Then
        
        'Inform User
        vMsgBox "There are no items to search from.", vbInformation, App.Title & " : No Items", vLv.Parent
        Exit Function 'Quit this Function
        
    End If 'Close respective IF..THEN block statement
    
    'If the specified Listview does not have items the
    If vLv.ListItems(&H1).Tag = "Function" Then
        
        'Inform User
        vMsgBox "There are no items to search from.", vbInformation, App.Title & " : No Items", vLv.Parent
        Exit Function 'Quit this Function
        
    End If 'Close respective IF..THEN block statement
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    Load Frm_Search 'Load this form into memory.
    
    'Execute a series of statements on Frm_Search Form
    With Frm_Search
        
        Set .TargetLv = vLv 'Specify Listview control
        
        'For each column header in the Listview...
        For vIndex(&H0) = &H1 To vLv.ColumnHeaders.Count Step &H1
            
            'Add it to the 'Look In' ComboBox in the specified Form
            .CboSearch.AddItem VBA.Trim$(vLv.ColumnHeaders(vIndex(&H0)).Text)
            
        Next vIndex(&H0) 'Move to the next column header
        
        If vSelCol > &H0 Then .CboSearch.ListIndex = vSelCol: .ChkOption(&H0).Value = vbChecked
        
        If Not Nothing Is .TargetLv.SelectedItem And vSelCol > &H0 Then
            If vSelCol = &H1 Then .TxtSearch.Text = .TargetLv.SelectedItem.Text Else .TxtSearch.Text = .TargetLv.SelectedItem.ListSubItems(vSelCol - &H1).Text
        Else
            .CboSearch.ListIndex = &H0 'Select the first item in the List
        End If
        
        If Not vEnableFilterOption Then .ChkOption(&H2).Value = vbUnchecked
        .ChkOption(&H2).Visible = vEnableFilterOption
        
        'Indicate that a process or operation is complete.
        Screen.MousePointer = vbDefault
        
        'Place the Form in the center of the Listview
        CenterForm Frm_Search, vLv
        
        'Indicate that a process or operation is in progress.
        Screen.MousePointer = vbHourglass
        
    End With 'Close the WITH block statements
    
    'Denote that the search is successfully complete
    SearchLv = True
    
Exit_SearchLv:
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Function 'Quit this Function
    
Handle_SearchLv_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Search Error - " & Err.Number, vLv.Parent
    
    'Resume execution at the specified Label
    Resume Exit_SearchLv
    
End Function

'Receives the listview, the column on which the click is made and type of data contained in the column
Public Function SortListview(vLv As ListView, LvCol As Integer, Optional LvSortOrder As LvwSortOrder = &H0)
On Local Error GoTo Handle_SortListview_Error
    
    'If there are no items in the Listview then Quit this Procedure
    If vLv.ListItems.Count = &H0 Then Exit Function
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    Dim LvLst As ListItem
    Dim LvColType$, iDateFormat$
    
    vLv.Visible = False
    
    For vIndex(&H0) = &H1 To vLv.ColumnHeaders.Count
        If vLv.ColumnHeaders(vIndex(&H0)).Position = LvCol Then LvCol = vLv.ColumnHeaders(vIndex(&H0)).Index: Exit For
    Next vIndex(&H0)
    
    iDateFormat = "dd/mm/yyyy hh:mm:ss"
    
    'Get the DataType of the specified Column
    
    Dim sArrayDataTypes() As String
    
    sArrayDataTypes = VBA.Split(vLv.Parent.iLvwItemDataType, "|")
    
    Select Case sArrayDataTypes(LvCol - &H1)
        
        Case "D":  LvColType = LvDateTime
        Case "N":  LvColType = LvNumber
        Case "C":  LvColType = LvCurrency
        Case Else: LvColType = LvText
        
    End Select 'End SELECT..CASE statement
    
    Dim Lst
    Dim ItemCnt&, LstCnt&
    Dim iItemTag() As String
    
    'For each listitem in the specified Listview
    For Each LvLst In vLv.ListItems
        
        If LvCol > &H1 Then Set Lst = LvLst.ListSubItems(LvCol - &H1) Else Set Lst = LvLst
        
        Lst.Tag = Lst.Text & "|" & Lst.Tag 'Preserve the original Text and Tag values
        
        'If the column datatype is..
        Select Case LvColType
            
            Case LvNumber: 'NUMBER then...
                
                If VBA.Replace(LvLst.Tag, "|", "") = "Function" Then
                    'Hinder the last column from being sorted to another position
                    Lst.Text = VBA.IIf(LvSortOrder = &H0, "9999999999.00000000", "0000000000.00000000")
                Else
                    Lst.Text = VBA.Format(VBA.Val(VBA.Replace(Lst.Text, ",", VBA.vbNullString)), "0000000000.00000000")
                End If
                
            Case LvCurrency:  'CURRENCY then...
                
                'Preserve the original Text and Tag values
                Lst.Tag = VBA.Format(VBA.Val(VBA.Replace(Lst.Text, ",", VBA.vbNullString)), "#,##0.00") & "|" & Lst.Tag 'Preserve the original Text and Tag values
                
                If VBA.Replace(LvLst.Tag, "|", "") = "Function" Then
                    'Hinder the last column from being sorted to another position
                    Lst.Text = VBA.IIf(LvSortOrder = &H0, "9999999999.00000000", "0000000000.00000000")
                Else
                    Lst.Text = VBA.Format(VBA.Val(VBA.Replace(Lst.Text, ",", VBA.vbNullString)), "0000000000.00000000")
                End If
                
            Case LvDateTime: 'DATE/TIME then...
                
                If VBA.InStr(VBA.LCase$(Lst.Text), "monday") <> &H0 Or VBA.InStr(VBA.LCase$(Lst.Text), "tuesday") <> &H0 Or VBA.InStr(VBA.LCase$(Lst.Text), "wednesday") <> &H0 Or VBA.InStr(VBA.LCase$(Lst.Text), "thursday") <> &H0 Or VBA.InStr(VBA.LCase$(Lst.Text), "friday") <> &H0 Or VBA.InStr(VBA.LCase$(Lst.Text), "saturday") <> &H0 Or VBA.InStr(VBA.LCase$(Lst.Text), "sunday") <> &H0 Then
                    
                    Lst.Text = VBA.Replace(VBA.Trim$(VBA.Replace(VBA.LCase$(Lst.Text), "monday", "")), "  ", " ")
                    Lst.Text = VBA.Replace(VBA.Trim$(VBA.Replace(VBA.LCase$(Lst.Text), "tuesday", "")), "  ", " ")
                    Lst.Text = VBA.Replace(VBA.Trim$(VBA.Replace(VBA.LCase$(Lst.Text), "wednesday", "")), "  ", " ")
                    Lst.Text = VBA.Replace(VBA.Trim$(VBA.Replace(VBA.LCase$(Lst.Text), "thursday", "")), "  ", " ")
                    Lst.Text = VBA.Replace(VBA.Trim$(VBA.Replace(VBA.LCase$(Lst.Text), "friday", "")), "  ", " ")
                    Lst.Text = VBA.Replace(VBA.Trim$(VBA.Replace(VBA.LCase$(Lst.Text), "saturday", "")), "  ", " ")
                    Lst.Text = VBA.Replace(VBA.Trim$(VBA.Replace(VBA.LCase$(Lst.Text), "sunday", "")), "  ", " ")
                    
                End If
                
                If VBA.InStr(VBA.LCase$(Lst.Text), "mon") <> &H0 Or VBA.InStr(VBA.LCase$(Lst.Text), "tue") <> &H0 Or VBA.InStr(VBA.LCase$(Lst.Text), "wed") <> &H0 Or VBA.InStr(VBA.LCase$(Lst.Text), "thu") <> &H0 Or VBA.InStr(VBA.LCase$(Lst.Text), "fri") <> &H0 Or VBA.InStr(VBA.LCase$(Lst.Text), "sat") <> &H0 Or VBA.InStr(VBA.LCase$(Lst.Text), "sun") <> &H0 Then
                    
                    Lst.Text = VBA.Replace(VBA.Trim$(VBA.Replace(VBA.LCase$(Lst.Text), "mon", "")), "  ", " ")
                    Lst.Text = VBA.Replace(VBA.Trim$(VBA.Replace(VBA.LCase$(Lst.Text), "tue", "")), "  ", " ")
                    Lst.Text = VBA.Replace(VBA.Trim$(VBA.Replace(VBA.LCase$(Lst.Text), "wed", "")), "  ", " ")
                    Lst.Text = VBA.Replace(VBA.Trim$(VBA.Replace(VBA.LCase$(Lst.Text), "thu", "")), "  ", " ")
                    Lst.Text = VBA.Replace(VBA.Trim$(VBA.Replace(VBA.LCase$(Lst.Text), "fri", "")), "  ", " ")
                    Lst.Text = VBA.Replace(VBA.Trim$(VBA.Replace(VBA.LCase$(Lst.Text), "sat", "")), "  ", " ")
                    Lst.Text = VBA.Replace(VBA.Trim$(VBA.Replace(VBA.LCase$(Lst.Text), "sun", "")), "  ", " ")
                    
                End If
                
                Lst.Text = VBA.FormatDateTime(Lst.Text, vbShortDate)
                
                If VBA.Replace(LvLst.Tag, "|", "") Then
                    'Hinder the last column from being sorted to another position
                    Lst.Text = VBA.IIf(LvSortOrder = &H0, VBA.DateSerial(9999, 12, 31), VBA.DateSerial(1800, &H1, &H1))
                End If
                
                Lst.Text = VBA.Format(VBA.CDate(Lst.Text), "yyyymmddhhnnss")
                
            Case LvText: 'TEXT/STRING then...
                
                If VBA.Replace(LvLst.Tag, "|", "") = "Function" Then
                    'Hinder the last column from being sorted to another position
                    Lst.Text = VBA.IIf(LvSortOrder = &H0, VBA.String$(50, "Z"), VBA.String$(50, " "))
                End If
                
        End Select 'End SELECT..CASE statement
        
    Next LvLst 'Move to the next listitem
    
    vLv.SortOrder = LvSortOrder
    vLv.SortKey = LvCol - &H1
    vLv.Sorted = True 'Sort data in the Listview
    vLv.Sorted = False 'Sort data in the Listview
    
    'Assign new values
    '-----------------
    
    'For each list item in the specified Listview
    For Each LvLst In vLv.ListItems
        
        If LvCol > &H1 Then Set Lst = LvLst.ListSubItems(LvCol - &H1) Else Set Lst = LvLst
        
        iItemTag = VBA.Split(Lst.Tag, "|")
        
        'If the Column's Data Type is a number then....
        If LvColType = LvNumber Or LvColType = LvCurrency Then
            
            Lst.Text = VBA.IIf(LvColType <> LvCurrency, iItemTag(&H0), VBA.FormatNumber(VBA.Val(VBA.Format(iItemTag(&H0), "0000.0000")), &H2))
            
        Else 'If the Column's Data Type is a DATE/TIME then....
            
            Lst.Text = iItemTag(&H0) 'Assign its original value
            
        End If 'End IF..THEN block statement
        
        Lst.Tag = VBA.vbNullString 'Initialize property
        
        'If the ListItem had a value at its tag property then reassign its initial data
        If UBound(iItemTag) >= &H1 Then Lst.Tag = iItemTag(&H1)
        
    Next LvLst 'Move to the next listitem
    
Exit_SortListview:
    
    vLv.Visible = True
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Function 'Quit this Function
    
Handle_SortListview_Error:
    
    'If failed to recognize a date then...
    If Err.Number = 13 And VBA.Val(LvColType) = LvDateTime Then
        
        Dim NewVal$
        Dim iArray() As String
        
        'Call Procedure in this Module to reformat it correctly
        NewVal = ReFormattedDate(Lst.Text)
        iArray = VBA.Split(NewVal, "|")
        If NewVal = VBA.vbNullString Then Resume Next Else iDateFormat = iArray(&H1): Lst.Text = iArray(&H0): Resume
        
    End If 'End IF..THEN block statement
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Listview Sorting Error " & Err.Number, vLv.Parent
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Resume execution at the specified Label
    Resume Exit_SortListview
    
End Function

'This Procedure sets the specified Form's transparency level to the specified level
Public Function Transparency(Obj As Object, Level%) As Boolean
    
    Dim Msg As Long
    
    Msg = GetWindowLong(Obj.hWnd, (-20))
    Msg = Msg Or &H80000
    SetWindowLong Obj.hWnd, (-20), Msg
    SetLayeredWindowAttributes Obj.hWnd, &H0, Level, &H2
    
End Function

Public Function ValidUserAccess(iFrm As Form, mySetNo&, mySetIndex&, Optional myWholeSet As Boolean = False, Optional myCustomMessage$ = "You have insufficient privileges to perform the Operation. Allow another User with sufficient privileges to grant you access?", Optional ClearVirtualUserEntryImmediately As Boolean = False, Optional iMsgTitle$ = VBA.vbNullString) As Boolean
On Local Error GoTo Handle_ValidUserAccess_Error
    
    'Give exclusive rights to the top-most User even when limited
    If User.Hierarchy = &H0 Then ValidUserAccess = True: GoTo Exit_ValidUserAccess
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    Dim bFrm As Form
    Dim FrmLoaded As Boolean
    Dim myAccessRightsArray() As String
        
    myAccessRightsArray = VBA.Split(VBA.IIf(VirtualUser.User_Name <> VBA.vbNullString, VirtualUser.Privileges, User.Privileges), "|")
    
    If mySetNo - &H1 > UBound(myAccessRightsArray) Then GoTo AccessDenied
    
    If myWholeSet Then
        ValidUserAccess = (VBA.Replace(myAccessRightsArray(mySetNo - &H1), "1", VBA.vbNullString) = VBA.vbNullString)
    Else
        If mySetIndex <= VBA.Len(myAccessRightsArray(mySetNo - &H1)) Then ValidUserAccess = (VBA.Mid$(myAccessRightsArray(mySetNo - &H1), mySetIndex, &H1) = &H1)
    End If 'End IF..THEN block statement
    
    If Not ValidUserAccess Then GoTo AccessDenied
    
Exit_ValidUserAccess:
    
    'Display the PleaseWait Form if it was loaded
    If FrmLoaded And ValidUserAccess Then Frm_PleaseWait.Show
    
    If ClearVirtualUserEntryImmediately Then VirtualUser.User_ID = &H0: VirtualUser.User_Name = VBA.vbNullString
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    Exit Function 'Quit this Function
    
AccessDenied:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'If the custom message has not been set to blank then...
    If myCustomMessage <> VBA.vbNullString And SoftwareSetting.RequestAlternativeLogin Then
        
        For Each bFrm In VB.Forms
            FrmLoaded = (bFrm.Name = "Frm_PleaseWait")
            If FrmLoaded Then Exit For
        Next bFrm
        
        'If the PleaseWait Form has been loaded then hide it in order to display the Message
        If FrmLoaded And myCustomMessage <> VBA.vbNullString Then Frm_PleaseWait.Hide
        
        'Indicate that a process or operation is complete.
        Screen.MousePointer = vbDefault
        
        'Warn User
        If vMsgBox(myCustomMessage, vbQuestion + vbYesNo + vbDefaultButton2, App.Title & " : " & VBA.IIf(iMsgTitle <> VBA.vbNullString, iMsgTitle, "Access Denied"), iFrm) = vbYes Then
            
            Frm_VirtualLogin.mySetNo = mySetNo
            Frm_VirtualLogin.mySetIndex = mySetIndex
            Frm_VirtualLogin.myWholeSet = myWholeSet
            
            CenterForm Frm_VirtualLogin, iFrm
            
            If VirtualUser.User_Name <> VBA.vbNullString Then ValidUserAccess = True
            
        End If 'End IF..THEN block statement
        
        'Indicate that a process or operation is in progress.
        Screen.MousePointer = vbHourglass
        
    Else
        
        If Not SoftwareSetting.RequestAlternativeLogin Then
            
            'Indicate that a process or operation is complete.
            Screen.MousePointer = vbDefault
            
            'Warn User
            vMsgBox "You have insufficient Privileges to perform the operation.", vbExclamation, App.Title & " : " & VBA.IIf(iMsgTitle <> VBA.vbNullString, iMsgTitle, "Access Denied"), iFrm
            
            'Indicate that a process or operation is in progress.
            Screen.MousePointer = vbHourglass
            
        End If 'End IF..THEN block statement
        
    End If 'End IF..THEN block statement
    
    'Resume execution at the specified Label
    GoTo Exit_ValidUserAccess
    
Handle_ValidUserAccess_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Edit Mode Confirmation Error - " & Err.Number, iFrm
    
    'Resume execution at the specified Label
    Resume Exit_ValidUserAccess
    
End Function

'
''This Procedure sets the specified Form's transparency level to the specified level
'Public Function ValidUserAccess(vFrm As Form, vSetNo&, Optional vPrivilegePos&, Optional vCustomMsg$ = "", Optional vShowMsg As Boolean = True) As Boolean
'
'    If User.Hierarchy = &H0 Then ValidUserAccess = True: Exit Function
'
'    Dim vArray() As String
'
'    vArray = VBA.Split(User.Privileges, "|")
'
'    'If the specified set no exists then...
'    If UBound(vArray) >= vSetNo - &H1 And vSetNo > &H0 Then
'
'        If VBA.Len(vArray(vSetNo - &H1)) >= vPrivilegePos Then
'
'            'If checking the whole set then...
'            If vPrivilegePos = &H0 Then
'                ValidUserAccess = (VBA.Replace(vArray(vSetNo - &H1), "0", "") <> "")
'            Else 'If checking a privilege in the set then...
'                ValidUserAccess = VBA.Mid$(vArray(vSetNo - &H1), vPrivilegePos, &H1)
'            End If 'Close respective IF..THEN block statement
'
'        Else
'            ValidUserAccess = True
'        End If 'Close respective IF..THEN block statement
'
'    End If 'Close respective IF..THEN block statement
'
'    'If the privilege has not been granted then...
'    If Not ValidUserAccess Then
'
'        'If message should be shown then...
'        If vShowMsg Then
'
'            'Indicate that a process or operation is complete.
'            Screen.MousePointer = vbDefault
'
'            'Warn User
'            vMsgBox VBA.IIf(vCustomMsg <> "", vCustomMsg, "You have insufficient privileges to execute this operation. Please contact Software Administrator."), vbExclamation, App.Title & " : Operation Denied", vFrm
'
'        Else 'If message should not be shown then...
'            VBA.Beep 'Sound a beep tone through the computer's speaker.
'        End If 'Close respective IF..THEN block statement
'
'    End If 'Close respective IF..THEN block statement
'
'End Function

Public Function Wait() As Boolean
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    Do While vCancelOperation = &H1
        VBA.DoEvents: VBA.DoEvents
        Sleep 100 'Wait for quarter a second
    Loop
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
End Function

'Minimize all Open windows 'STATE=TRUE {Minimize All}, STATE=FALSE {Restore All}
Public Sub WindowsMinimizeAll(Optional State As Boolean = True)
On Error Resume Next
    
    Dim lngHwnd&
    
    lngHwnd = FindWindow("Shell_TrayWnd", vbNullString)
    Call PostMessage(lngHwnd, WM_COMMAND, VBA.IIf(State, MIN_ALL, MIN_ALL_UNDO), 0&)
  
End Sub

' This function creates an integer value based on the key
' provided.  The principle is simple.  The result is the
' absolute value of the difference between the averages of
' the odd and even characters.

Public Function CreateEncryptCode(Key As String) As Integer
    
    Dim Total(&H1 To &H2) As Integer
    Dim NbChars(&H1 To &H2) As Integer
    Dim vIndex, Index As Integer
    
    Total(&H1) = &H0: Total(&H2) = &H0
    NbChars(&H1) = &H0: NbChars(&H2) = &H0
    
    For vIndex = &H1 To VBA.LenB(Key) Step &H1
        
        Index = VBA.IIf(vIndex Mod &H2 = &H0, &H1, &H2) ' Characters in an even/odd position
        Total(Index) = Total(Index) + VBA.Asc(VBA.Mid(Key, vIndex, &H1))
        NbChars(Index) = NbChars(Index) + &H1
        
    Next vIndex
    
    ' A division by zero must be avoided.
    ' This will be the new value used for encryption
    ' Else If the key is less than 2 characters long, the code becomes &H1
    CreateEncryptCode = VBA.IIf(NbChars(&H1) > &H0 And NbChars(&H2) > &H0, VBA.Abs((Total(&H1) / NbChars(&H1)) - (Total(&H2) / NbChars(&H2))), &H1)
    
End Function

' I prefer alternating between an addition and a subtraction
' to provide a more complex encryption method.  It is more
' difficult to crack due to the alternations and the
' encryption key.


' OrigStr : The original string value before encryption.
' Key     : The key used for encrypting/decrypting the string

Public Function EncryptStr(ByVal OrigStr As String, Optional Key As String = "Ketheline", Optional Decrypt As Boolean = False) As String
    
    'If no value has been supplied then quit this Function
    If VBA.LenB(VBA.Trim$(OrigStr)) = &H0 Then Exit Function
    
    Dim vIndex, EncCode As Integer
    
    ' First thing done is a calculation upon the encryption key
    ' to determine how the original string will be encrypted.
    EncCode = CreateEncryptCode(Key)
    EncryptStr = VBA.vbNullString
    
    ' Now the string will be changed according to the new encryption values
    For vIndex = &H1 To VBA.LenB(OrigStr) Step &H1
        EncryptStr = EncryptStr + VBA.IIf(Decrypt, VBA.IIf(vIndex Mod &H2 = &H0, VBA.Chr(VBA.Asc(VBA.Mid(OrigStr, vIndex, &H1)) - EncCode), VBA.Chr(VBA.Asc(VBA.Mid(OrigStr, vIndex, &H1)) + EncCode)), VBA.IIf(vIndex Mod &H2 = &H0, VBA.Chr(VBA.Asc(VBA.Mid(OrigStr, vIndex, &H1)) + EncCode), VBA.Chr(VBA.Asc(VBA.Mid(OrigStr, vIndex, &H1)) - EncCode)))
    Next vIndex
    
End Function

'------------------------------------------------------------------------------------------
Public Function SmartEncrypt(StringToEncrypt As String, Optional AlphaEncoding As Boolean = True) As String
On Local Error GoTo ErrorHandler
    
    Dim i&
    Dim Char$, strEncrypt$
    
    If VBA.Len(VBA.Trim$(StringToEncrypt)) = &H0 Then Exit Function
    
    For i = &H1 To VBA.Len(StringToEncrypt) Step &H1
        Char = VBA.Asc(VBA.Mid(StringToEncrypt, i, &H1))
        SmartEncrypt = SmartEncrypt & VBA.Len(Char) & Char
    Next i
    
    If AlphaEncoding Then
    
        strEncrypt = SmartEncrypt
        SmartEncrypt = VBA.vbNullString
        
        For i = &H1 To VBA.Len(strEncrypt) Step &H1
            SmartEncrypt = SmartEncrypt & VBA.Chr(VBA.Mid(strEncrypt, i, &H1) + &H93)
        Next i
        
    End If
    
    Exit Function
    
ErrorHandler:
    
    SmartEncrypt = "Error encrypting string"
    
End Function

Public Function SmartDecrypt(StringToDecrypt As String, Optional AlphaDecoding As Boolean = True) As String
On Local Error GoTo ErrorHandler
    
    Dim i&
    Dim CharPos%
    Dim Char$, CharCode$, strEncrypt$
    
    If VBA.LenB(VBA.Trim$(StringToDecrypt)) = &H0 Then Exit Function
        
    If AlphaDecoding Then
    
        SmartDecrypt = StringToDecrypt
        
        For i = &H1 To VBA.Len(SmartDecrypt) Step &H1
            strEncrypt = strEncrypt & (VBA.Asc(VBA.Mid(SmartDecrypt, i, &H1)) - &H93)
        Next i
        
    End If
    
    SmartDecrypt = VBA.vbNullString
    
    If VBA.LenB(VBA.Trim$(strEncrypt)) = &H0 Then strEncrypt = StringToDecrypt
    
    Do While VBA.LenB(VBA.Trim$(strEncrypt)) <> &H0
        
        CharPos = VBA.Left(strEncrypt, &H1)
        strEncrypt = VBA.Mid(strEncrypt, &H2)
        CharCode = VBA.Left(strEncrypt, CharPos)
        strEncrypt = VBA.Mid(strEncrypt, VBA.Len(CharCode) + &H1)
        SmartDecrypt = SmartDecrypt & VBA.Chr(CharCode)
                
    Loop
    
    Exit Function
    
ErrorHandler:
    
End Function

'------------------------------------------------------------------------------------------
'The Following Codes convert currency (in Kshs) numbers to Words.
'It is essential for Receipt Generation

'Define values between 0 and 9
Private Function Ones(mValue%) As String
    
    Select Case mValue
        
        Case &H0:
        Case &H1: Ones = "One"
        Case &H2: Ones = "Two"
        Case &H3: Ones = "Three"
        Case &H4: Ones = "Four"
        Case &H5: Ones = "Five"
        Case &H6: Ones = "Six"
        Case &H7: Ones = "Seven"
        Case &H8: Ones = "Eight"
        Case &H9: Ones = "Nine"
        
    End Select
    
End Function

'Define values between 10 and 20
Private Function Teens(mValue%) As String
    
    Select Case mValue
        
        Case &HB: Teens = "Eleven"
        Case &HC: Teens = "Twelve"
        Case &HD: Teens = "Thirteen"
        Case &HE: Teens = "Fourteen"
        Case &HF: Teens = "Fifteen"
        Case &H10: Teens = "Sixteen"
        Case &H11: Teens = "Seventeen"
        Case &H12: Teens = "Eighteen"
        Case &H13: Teens = "Nineteen"
        
    End Select
    
End Function

'Define Tens values
Private Function Tens(mValue%) As String
    
    Dim mChar%
    
    mChar = VBA.Mid$(mValue, &H1, &H1)
    
    Select Case mChar
        
        Case &H1: Tens = "Ten"
        Case &H2: Tens = "Twenty"
        Case &H3: Tens = "Thirty"
        Case &H4: Tens = "Fourty"
        Case &H5: Tens = "Fifty"
        Case &H6: Tens = "Sixty"
        Case &H7: Tens = "Seventy"
        Case &H8: Tens = "Eighty"
        Case &H9: Tens = "Ninety"
        
    End Select
    
    mChar = VBA.Mid$(mValue, &H2, &H1)
    
    If mChar <> &H0 Then Tens = VBA.Trim$(Tens & " " & Ones(mChar))
    
End Function

Public Function ConvertNoToText(mValue$, Optional IsMantissa As Boolean = True, Optional IsCurrency As Boolean = True) As String
    
    Dim iValue$
    Dim iNum As Long
    Dim iDecimal() As String
    Dim iTriNo() As String
    Dim iTriWord() As String
    
    mValue = VBA.Replace(VBA.FormatNumber(mValue, &H2), ",", VBA.vbNullString)
    iDecimal = VBA.Split(mValue, ".")
    
    If UBound(iDecimal) < &H0 Then Exit Function
    
    mValue = VBA.Val(iDecimal(&H0))
    If VBA.Val(mValue) = &H0 Then GoTo ConvertDecimals
    iValue = VBA.IIf(mValue = "0", mValue, VBA.Fix((VBA.Len(mValue) - &H1) / &H3) + &H1)
    iValue = VBA.IIf(iValue = "0", iValue, VBA.Format$(VBA.Format$(mValue, "#,###"), "#," & VBA.String$(VBA.Val(iValue) * 3, "0")))
    
    iTriNo = VBA.Split(iValue, ",")
    iTriWord = VBA.Split(iValue, ",")
    
    For iNum = LBound(iTriNo) To UBound(iTriNo) Step &H1
        
        Dim iCurrNo%
        
        iTriWord(iNum) = VBA.vbNullString
        iCurrNo = VBA.Val(iTriNo(iNum))
        
CheckTriValue:
        
        Select Case iCurrNo
            
            Case Is < &HA:
                
                iTriWord(iNum) = VBA.Trim$(iTriWord(iNum) & " " & Ones(VBA.Val(iCurrNo)))
                
            Case Is > 99:
                
                iTriWord(iNum) = Ones(VBA.Mid$(iCurrNo, &H1, &H1)) & " Hundred"
                
                If VBA.Val(VBA.Mid$(iCurrNo, &H2, &H3)) <> &H0 Then
                    iTriWord(iNum) = iTriWord(iNum) & " and"
                    iCurrNo = VBA.Val(VBA.Mid$(iCurrNo, &H2, &H3))
                    GoTo CheckTriValue
                End If
                
            Case Is >= &HA:
                
                If iCurrNo >= &HB And iCurrNo <= 19 Then
                    iTriWord(iNum) = VBA.Trim$(iTriWord(iNum) & " " & Teens(iCurrNo))
                Else
                    iTriWord(iNum) = VBA.Trim$(iTriWord(iNum) & " " & Tens(iCurrNo))
                End If
                
        End Select
        
        Select Case UBound(iTriNo) - iNum
            
            Case &H0:
            Case &H1: iTriWord(iNum) = iTriWord(iNum) & " Thousand"
            Case &H2: iTriWord(iNum) = iTriWord(iNum) & " Million"
            Case &H3: iTriWord(iNum) = iTriWord(iNum) & " Billion"
            Case &H4: iTriWord(iNum) = iTriWord(iNum) & " Trillion"
            Case Is >= &H6: iTriWord(iNum) = iTriWord(iNum) & " ?"
            
        End Select
        
        If UBound(iTriNo) > &H0 And iNum <> UBound(iTriNo) Then iTriWord(iNum) = iTriWord(iNum) & VBA.IIf(VBA.Val(iTriNo(iNum + 1)) > 0, ",", VBA.vbNullString)
        
    Next iNum
    
ConvertDecimals:
    
    ConvertNoToText = VBA.Join(iTriWord, " ")
    ConvertNoToText = VBA.Replace(ConvertNoToText & VBA.IIf(IsCurrency, VBA.IIf(IsMantissa, " Shilling", " Cent") & VBA.IIf(ConvertNoToText = "One", VBA.vbNullString, "s"), VBA.vbNullString), "  ", " ")
    
    If UBound(iDecimal) = &H1 Then
        If VBA.Val(iDecimal(&H1)) <> &H0 Then ConvertNoToText = ConvertNoToText & VBA.IIf(ConvertNoToText = VBA.vbNullString, VBA.vbNullString, " and ") & ConvertNoToText(iDecimal(1), False)
    End If
    
    ConvertNoToText = VBA.Trim$(VBA.IIf(VBA.Left$(ConvertNoToText, &H5) = " Shil", VBA.Replace(ConvertNoToText, " Shillings and", VBA.vbNullString), ConvertNoToText))
    ConvertNoToText = VBA.IIf((ConvertNoToText = VBA.vbNullString Or ConvertNoToText = "Shillings") And mValue = 0, "Zero Shillings", ConvertNoToText)
    
End Function

Public Function ConvertNo(No$) As String
    'If the entered number is zero then return blank else convert to word
    If No = VBA.vbNullString Then ConvertNo = VBA.vbNullString Else ConvertNo = ConvertNoToText(No) & " Only"
End Function
