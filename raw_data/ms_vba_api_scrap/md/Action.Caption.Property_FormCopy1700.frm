VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormCopy1700 
   Caption         =   "17.00 uur Afspraken overnemen naar actuele afspraken"
   ClientHeight    =   11130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16590
   OleObjectBlob   =   "FormCopy1700.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormCopy1700"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const constVoeding As String = "Txt_Neo_InfB_Voeding_"
Private Const constVoedingCount As Integer = 10
Private Const constIVCont As String = "Txt_Neo_InfB_ContIV_"
Private Const constContIVCount As Integer = 16
Private Const constTPN As String = "Txt_Neo_InfB_TPN_"
Private Const constTPNCount As Integer = 13

Private Sub chkContinueMedicatie_Change()

    ContMedSelected chkContinueMedicatie.Value

End Sub

Private Sub chkTPN_Change()

    TPNSelected chkTPN.Value

End Sub

Private Sub chkVoeding_Change()

    VoedingSelected chkVoeding.Value

End Sub

Private Sub cmdCancel_Click()
    
    Me.Hide

End Sub

Private Sub cmdOK_Click()
    
    lblAction.Caption = "OK"
    Me.Hide

End Sub

Private Sub optAlles_Click()
    
    chkContinueMedicatie.Value = True
    chkTPN.Value = True
    chkVoeding.Value = True

End Sub

Private Sub optPerBlok_Click()
    
    chkContinueMedicatie.Value = False
    chkTPN.Value = False
    chkVoeding.Value = False

End Sub

Private Sub RemoveDoubles(ByVal strList As String)

    Dim intActN As Integer
    Dim intActC As Integer
    Dim int1700N As Integer
    Dim int1700C As Integer
    Dim strAct As String
    Dim str1700 As String
    Dim objListAct As MSForms.ListBox
    Dim objList1700 As MSForms.ListBox
    
    Set objListAct = Me.Controls(strList)
    strList = Replace(strList, "Act", "1700")
    Set objList1700 = Me.Controls(strList)
    
    intActC = objListAct.ListCount - 1
    int1700C = objList1700.ListCount - 1
    
    For intActN = 0 To intActC
        strAct = objListAct.List(intActN)
        
        If Not strAct = vbNullString Then
            For int1700N = 0 To int1700C
                str1700 = objList1700.List(int1700N)
                If strAct = str1700 Then
                    objListAct.List(intActN) = vbNullString
                    objList1700.List(int1700N) = vbNullString
                    
                    Exit For
                End If
            Next
        End If
    Next

End Sub

Private Sub AddItemToList(ByVal strList As String, ByVal strItem As String, ByVal intN As Integer, ByVal bln1700 As Boolean)

    ' Do not add vocht intake as item
    If strItem = constTPN And intN = 2 Then Exit Sub

    strList = IIf(bln1700, Replace(strList, "Act", "1700"), strList)
    strItem = IIf(intN < 10, strItem & "0" & intN, strItem & intN)
    
    Me.Controls(strList).AddItem ModRange.GetRangeValue(strItem, vbNullString)
    
End Sub

Private Sub AddItems(ByVal bln1700 As Boolean)

    Dim intN As Integer
    Dim strList As String
    Dim strItem As String
        
    strList = "lstActVoed"
    strItem = constVoeding
    For intN = 1 To constVoedingCount
        AddItemToList strList, strItem, intN, bln1700
    Next intN

    strList = "lstActMed"
    strItem = constIVCont
    For intN = 1 To constContIVCount
        AddItemToList strList, strItem, intN, bln1700
    Next intN

    strList = "lstActTPN"
    strItem = constTPN
    For intN = 1 To constTPNCount
        AddItemToList strList, strItem, intN, bln1700
    Next intN
    
End Sub

Private Sub CenterForm()

    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)

End Sub

Private Sub UserForm_Activate()

    CenterForm

End Sub

Private Sub VoedingSelected(ByVal blnSelect As Boolean)

    If blnSelect Then
        lstActVoed.BackColor = &H80000005
        lst1700Voed.BackColor = &H80000005
    Else
        lstActVoed.BackColor = &H8000000B
        lst1700Voed.BackColor = &H8000000B
    End If

End Sub

Private Sub TPNSelected(ByVal blnSelect As Boolean)

    If blnSelect Then
        lstActTPN.BackColor = &H80000005
        lst1700TPN.BackColor = &H80000005
    Else
        lstActTPN.BackColor = &H8000000B
        lst1700TPN.BackColor = &H8000000B
    End If

End Sub

Private Sub ContMedSelected(ByVal blnSelect As Boolean)

    If blnSelect Then
        lstActMed.BackColor = &H80000005
        lst1700Med.BackColor = &H80000005
    Else
        lstActMed.BackColor = &H8000000B
        lst1700Med.BackColor = &H8000000B
    End If

End Sub

Private Sub UserForm_Initialize()

    ModProgress.StartProgress "Afspraken laden"
    
    ' First get the actual items
    ModNeoInfB.NeoInfB_SelectInfB False, False
    AddItems False
    
    ' Then get the 1700 items
    ModNeoInfB.NeoInfB_SelectInfB True, False
    AddItems True
    
    RemoveDoubles "lstActVoed"
    RemoveDoubles "lstActMed"
    RemoveDoubles "lstActTPN"
    
    ' Set defaults
    optPerBlok.Value = True
    chkTPN.Value = True
    VoedingSelected False
    ContMedSelected False
    
    lblBijschrift.Caption = "Per item groep (voeding, TPN, continue medicatie) worden alleen de afspraken getoond die verschillen. Zijn er geen verschillen dan is de item groep leeg."
    lblBijschrift.Caption = lblBijschrift.Caption & vbNewLine & "N.B. Actuele afspraken worden overschreven voor de geselecteerde item groep(-en)."
    
    cmdOK.SetFocus
    
    ModProgress.FinishProgress

End Sub
