Option Explicit
Dim CustID As String, SelItem As Long
#If VBA7 Then
    Private Declare PtrSafe Function GetDC& Lib "user32.dll" (ByVal hwnd&)
    Private Declare PtrSafe Function GetDeviceCaps& Lib "gdi32" (ByVal hDC&, ByVal nIndex&)
#Else
    Private Declare Function GetDC& Lib "user32.dll" (ByVal hwnd&)
    Private Declare Function GetDeviceCaps& Lib "gdi32" (ByVal hDC&, ByVal nIndex&)
#End If

Sub ShowCustMgmtForm()
Dim XPos As Double, YPos As Double, LeftPos As Double, TopPos As Double
Dim LastRow As Long, LastResultRow As Long, ContRow As Long, ContCol As Long
Dim EmailCol As Long, EmailRow As Long, ApptCol As Long, ApptRow As Long, DocRow As Long, DocCol As Long
Dim PicFolder As String, PicName As String, PicPath As String, DocIcon As String
PicFolder = [ContactPicFolder] 'Set Contact Picture Folder
CustMgmtFrm.CustNmLbl.Caption = Customers.Range("B" & ActiveCell.Row).Value 'Customer Name

'Initialize Customer Cont. List View
With CustMgmtFrm.LvContacts
        .FullRowSelect = True
        .View = lvwReport
End With

'Load Customer Contact Details
With CustContDB
    For ContCol = 19 To 22
        With CustMgmtFrm.LvContacts.ColumnHeaders.Add(, , .Cells(2, ContCol).Value, .Cells(2, ContCol).Width)
            .Alignment = lvwColumnLeft
        End With
    Next ContCol
'Get Contact Data
LastRow = .Range("A99999").End(xlUp).Row 'Last Contact Row
If LastRow < 3 Then GoTo NoContacts
.Range("O3").Value = Customers.Range("A" & ActiveCell.Row).Value 'Set Customer ID as criteria
.Range("A3:J" & LastRow).AdvancedFilter xlFilterCopy, CriteriaRange:=.Range("O2:O3"), CopyToRange:=.Range("S2:X2"), Unique:=True
LastResultRow = .Range("S99999").End(xlUp).Row
If LastResultRow < 3 Then GoTo NoContacts
If LastResultRow < 4 Then GoTo SkipContSort
    With .Sort
    .SortFields.Clear
    .SortFields.Add Key:=CustContDB.Range("S3"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal  'Sort
    .SetRange CustContDB.Range("S3:X" & LastResultRow) 'Set Range
    .Apply 'Apply Sort
    End With
    
    For ContRow = 3 To LastResultRow
            If .Range("W" & ContRow).Value <> Empty Then  'Contact Picture exists
                PicName = .Range("W" & ContRow).Value
                PicPath = PicFolder & "\" & PicName 'Picture Path
                If Dir(PicPath, vbDirectory) <> "" Then 'Accurate Picture Path
                    On Error Resume Next
                    CustMgmtFrm.ImageList.ListImages.Add , PicName, LoadPicture(PicPath) 'Load Pictures Into Image List Object
                    On Error GoTo 0
                End If
            End If
    Next ContRow
    On Error Resume Next
    CustMgmtFrm.LvContacts.SmallIcons = CustMgmtFrm.ImageList 'Associate Images In List with List View
    
For ContRow = 3 To LastResultRow
    PicName = .Range("W" & ContRow).Value
    With CustMgmtFrm.LvContacts.ListItems.Add(, , "     " & CustContDB.Cells(ContRow, 19).Value) 'Add Contact Name
        If CustContDB.Range("W" & ContRow).Value <> Empty Then .SmallIcon = PicName
        For ContCol = 20 To 22
            .ListSubItems.Add , , CustContDB.Cells(ContRow, ContCol).Value
        Next ContCol
    End With
    
Next ContRow

SkipContSort:
NoContacts:
End With

'Add Customer Emails
     With CustMgmtFrm.LvEmails
            .FullRowSelect = True
            .View = lvwReport
     End With

With EmailLogDB
        'Add Headers In
        For EmailCol = 19 To 22
            CustMgmtFrm.LvEmails.ColumnHeaders.Add , , .Cells(2, EmailCol).Value, .Cells(2, EmailCol).Width
        Next EmailCol
        LastRow = .Range("A99999").End(xlUp).Row 'Last Contact Row
        If LastRow < 3 Then GoTo NoEmails
        .Range("O3").Value = Customers.Range("A" & ActiveCell.Row).Value 'Set Customer ID as criteria
        .Range("A3:H" & LastRow).AdvancedFilter xlFilterCopy, CriteriaRange:=.Range("O2:O3"), CopyToRange:=.Range("S2:W2"), Unique:=True
        LastResultRow = .Range("S99999").End(xlUp).Row
        If LastResultRow < 3 Then GoTo NoEmails
        If LastResultRow < 4 Then GoTo SkipEmailSort
            With .Sort
            .SortFields.Clear
            .SortFields.Add Key:=EmailLogDB.Range("S3"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal  'Sort
            .SetRange EmailLogDB.Range("S3:W" & LastResultRow) 'Set Range
            .Apply 'Apply Sort
            End With
SkipEmailSort:
        For EmailRow = 3 To LastResultRow
            With CustMgmtFrm.LvEmails.ListItems.Add(, , EmailLogDB.Range("S" & EmailRow).Value)
                For EmailCol = 20 To 22
                    If EmailCol = 21 Then 'Sent At time Column
                       .ListSubItems.Add , , Format(EmailLogDB.Cells(EmailRow, EmailCol).Value, "[$-en-US]h:mm AM/PM;@")
                    Else
                        .ListSubItems.Add , , EmailLogDB.Cells(EmailRow, EmailCol).Value
                    End If
                Next EmailCol
            End With
        Next EmailRow
End With
NoEmails:

'Add Customer Appointments
     With CustMgmtFrm.LvAppts
            .FullRowSelect = True
            .View = lvwReport
     End With

With Appts
        'Add Headers In
        For ApptCol = 36 To 41
            CustMgmtFrm.LvAppts.ColumnHeaders.Add , , .Cells(2, ApptCol).Value, .Cells(2, ApptCol).Width
        Next ApptCol
        LastRow = .Range("A99999").End(xlUp).Row 'Last Appt Row
        If LastRow < 3 Then GoTo NoAppts
        .Range("AF3").Value = Customers.Range("A" & ActiveCell.Row).Value 'Set Customer ID as criteria
        .Range("A4:K" & LastRow).AdvancedFilter xlFilterCopy, CriteriaRange:=.Range("AF2:AF3"), CopyToRange:=.Range("AI2:AP2"), Unique:=True
        LastResultRow = .Range("AI99999").End(xlUp).Row
        If LastResultRow < 3 Then GoTo NoAppts
        If LastResultRow < 4 Then GoTo SkipApptSort
            With .Sort
            .SortFields.Clear
            .SortFields.Add Key:=Appts.Range("AL3"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal  'Sort
            .SetRange Appts.Range("AI3:AP" & LastResultRow) 'Set Range
            .Apply 'Apply Sort
            End With
SkipApptSort:
        For ApptRow = 3 To LastResultRow
            With CustMgmtFrm.LvAppts.ListItems.Add(, , Appts.Range("AJ" & ApptRow).Value) 'Interaction Type (First Col)
                For ApptCol = 37 To 41
                    If ApptCol = 39 Then 'Sent At time Column
                          .ListSubItems.Add , , Format(Appts.Cells(ApptRow, ApptCol).Value, "[$-en-US]h:mm AM/PM;@")
                    ElseIf ApptCol = 40 Then 'Duration
                        .ListSubItems.Add , , Format(Appts.Cells(ApptRow, ApptCol).Value, "h:mm;@")
                    Else
                        .ListSubItems.Add , , Appts.Cells(ApptRow, ApptCol).Value
                    End If
                Next ApptCol
            End With
        Next ApptRow
End With
NoAppts:

'Add Customer Attachments & Documents
'Add Icons Picture to Image View
With CustMgmtFrm.ImageList
On Error Resume Next
    .ListImages.Add , "Excel_Icon.jpg", LoadPicture([DocumentsFolder] & "\" & "Excel_Icon.jpg")
    .ListImages.Add , "PDF_Icon.jpg", LoadPicture([DocumentsFolder] & "\" & "PDF_Icon.jpg")
    .ListImages.Add , "Picture_Icon.jpg", LoadPicture([DocumentsFolder] & "\" & "Picture_Icon.jpg")
    .ListImages.Add , "Word_Icon.jpg", LoadPicture([DocumentsFolder] & "\" & "Word_Icon.jpg")
    .ListImages.Add , "Other_Icon.jpg", LoadPicture([DocumentsFolder] & "\" & "Other_Icon.jpg")
On Error Resume Next
End With
With CustMgmtFrm.LvDocs
        .View = lvwIcon 'Set To Icon View
        .HideColumnHeaders = True
        .Icons = CustMgmtFrm.ImageList 'Assign Image List To LV Docs
End With
With DocDB
        LastRow = .Range("A99999").End(xlUp).Row 'Last Appt Row
        If LastRow < 3 Then GoTo NoDocs
        .Range("N3").Value = Customers.Range("A" & ActiveCell.Row).Value 'Set Customer ID as criteria
        .Range("A3:G" & LastRow).AdvancedFilter xlFilterCopy, CriteriaRange:=.Range("N2:N3"), CopyToRange:=.Range("R2:V2"), Unique:=True
        LastResultRow = .Range("R99999").End(xlUp).Row
        If LastResultRow < 3 Then GoTo NoDocs
        For DocRow = 3 To LastResultRow
            Select Case .Range("T" & DocRow).Value 'Set Type Of Icon
                Case "jpg", "jpeg", "bmp", "png", "giff"
                    DocIcon = "Picture_Icon.jpg"
                Case "xls", "xlsx", "xlsm", "xlsb"
                    DocIcon = "Excel_Icon.jpg"
                Case "pdf"
                    DocIcon = "PDF_Icon.jpg"
                Case "docx", "doc"
                    DocIcon = "Word_Icon.jpg"
                Case Else
                    DocIcon = "Other_Icon.jpg"
            End Select
            CustMgmtFrm.LvDocs.ListItems.Add , , .Range("S" & DocRow).Value, DocIcon
        Next DocRow

End With
NoDocs:

With CustMgmtFrm
        XPos = GetDeviceCaps(GetDC(0), 88) / 72 'Set Leftmost Horizontal position
        YPos = GetDeviceCaps(GetDC(0), 90) / 72 'Set Upper Vertical Screen Position
        LeftPos = ActiveCell.Left
        TopPos = ActiveCell.Top
        .Left = (ActiveWindow.PointsToScreenPixelsX(LeftPos * XPos) * 1 / XPos)
        .Top = (ActiveWindow.PointsToScreenPixelsY(TopPos * YPos) * 1 / YPos) + ActiveCell.Height
        .LvContacts.SetFocus
        .Show
End With
End Sub

Sub GoToSelContact()
Dim ContactDBRow As Long
CustID = Customers.Range("A" & ActiveCell.Row).Value 'Customer ID
SelItem = CustMgmtFrm.LvContacts.SelectedItem.Index 'Select List View Item #
CustMgmt.Range("B2").Value = CustID 'Place Customer ID in cell
Customer_Load 'Run Macro To Load Cust. Info
Customer_ContactTab 'Display Customer Contact Tab
ContactDBRow = CustContDB.Range("X" & SelItem + 2).Value 'Contact DB Row
CustMgmt.Range("B16").Value = ContactDBRow 'Set Contact DB Row
Contact_Load 'Load Contact
Unload CustMgmtFrm
CustMgmt.Activate
End Sub

Sub GoToSelDocument()
Dim DocDBRow As Long
CustID = Customers.Range("A" & ActiveCell.Row).Value 'Customer ID
SelItem = CustMgmtFrm.LvDocs.SelectedItem.Index 'Select List View Item #
DocDBRow = DocDB.Range("V" & SelItem + 2).Value 'Set Doc. DB Row
CustMgmt.Range("B2").Value = CustID 'Place Customer ID in cell
Customer_Load 'Run Macro To Load Cust. Info
Customer_DocumentTab 'Run macro to load documents tab
CustMgmt.Range("B8").Value = DocDBRow 'Set Doc. DB Row
Customer_LoadDocument 'Load Document
Unload CustMgmtFrm 'Clears Userform and hides
CustMgmt.Activate 'Activeate Cust. Mgr Sheet
End Sub

Sub GoToSelAppt()
Dim ApptID As String
SelItem = CustMgmtFrm.LvAppts.SelectedItem.Index 'Select List View Item #
ApptID = Appts.Range("AI" & SelItem + 2).Value 'Appt ID
Calendar.Range("B10").Value = ApptID 'Set Appt. ID
Calendar.Activate
Unload CustMgmtFrm 'Clears Userform and hides
Appt_Load 'Run Macro to Load Appt.
End Sub
