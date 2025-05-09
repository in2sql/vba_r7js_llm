VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "xxProcessItemLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Enum ControlTypeEnum
    ControlType_BackgroundLabel = 1
    ControlType_StatusIcon = 2
    ControlType_NameLabel = 3
    ControlType_RecordsCounterLabel = 4
    ControlType_ExpanderButton = 5
    ControlType_DetailsBackground = 6
    ControlType_DescriptionLabel = 7
    ControlType_PreviewButton = 8
    ControlType_ReprocessItemButton = 9
    ControlType_RemoveItemButton = 10
End Enum

Private Const CLASS_NAME As String = "xxProcessItemLine"
'[Caption tags] -------------------------------------------------------------------------------------------
Private Const OPEN_ITEM_IN_SEPARATE_WINDOW As String = "ProcessingItems.Actions.ItemPreview"
Private Const REPROCESS_ITEM_TOOLTIP As String = "ProcessingItems.Actions.ReprocessItem"
Private Const REMOVE_ITEM_TOOLTIP As String = "ProcessingItems.Actions.RemoveItem"
'[UI constants] -------------------------------------------------------------------------------------------
Private Const STATUS_ICON_LEFT_OFFSET As Single = 3
Private Const STATUS_ICON_TOP_OFFSET As Single = 3
Private Const STATUS_ICON_WIDTH As Single = 15
Private Const STATUS_ICON_HEIGHT As Single = 12
Private Const NAME_LABEL_TO_ICON_OFFSET As Single = 3
Private Const NAME_LABEL_HEIGHT As Single = 15
Private Const DEFAULT_FONT_FAMILY As String = "Segoe UI" '"Century Gothic"
Private Const DEFAULT_FONT_SIZE As Long = 9
Private Const DEFAULT_HORIZONTAL_SPACE As Single = 3
Private Const DEFAULT_LEFT_INDENT As Single = 0
Private Const DEFAULT_LINE_HEIGHT As Single = 18
Private Const EXPAND_ICON_TOP As Single = 3
Private Const EXPAND_ICON_WIDTH As Single = 15
Private Const EXPAND_ICON_HEIGHT As Single = 15
Private Const EXPAND_ICON_RIGHT As Single = 16
Private Const ICONS_SPACE As Single = 3
Private Const DESCRIPTION_BOTTOM_PADDING As Single = 1.5
'[UI components] ------------------------------------------------------------------------------------------
Private pParent As XProcessItemsList
Private pFrame As MSForms.frame
Private pBackgroundLabel As MSForms.label
Private pStatusIcon As MSForms.label
Private pNameLabel As MSForms.label
Private pRecordsCounterLabel As MSForms.label
Private WithEvents pReprocessButton As MSForms.label
Attribute pReprocessButton.VB_VarHelpID = -1
Private WithEvents pPreviewItemButton As MSForms.label
Attribute pPreviewItemButton.VB_VarHelpID = -1
Private WithEvents pExpandCollapseButton As MSForms.label
Attribute pExpandCollapseButton.VB_VarHelpID = -1
Private WithEvents pRemoveButton As MSForms.label
Attribute pRemoveButton.VB_VarHelpID = -1
Private pDetailsBackground As MSForms.label
Private pDescriptionLabel As MSForms.label
'[UI variables] -------------------------------------------------------------------------------------------
Private pTop As Single
Private pLeft As Single
Private pLineHeight As Single
Private pHorizontalSpace As Single
Private pFontFamily As String
Private pFontSize As Single
Private pStatusesDictionary As Scripting.Dictionary
'[State] --------------------------------------------------------------------------------------------------
Private pItem As IProcessItem
Private pIsRendered As Boolean
Private pIsDetailsPanelRendered As Boolean
Private pName As String
Private pIndex As Long
Private pStatus As ItemProcessStatusEnum
Private pIsExpanded As Boolean
Private pDescription As String
Private pIsHidden As Boolean
'[Events] -------------------------------------------------------------------------------------------------
Event AfterExpandCollapse()
Event AfterPreviewItemClicked()
Event AfterReprocessClicked()
Event AfterRemoveClicked()
'----------------------------------------------------------------------------------------------------------




'[INITIALIZE]
Private Sub Class_Initialize()
    Call initializeDefaultValues
End Sub

Private Sub initializeDefaultValues()
    pLeft = DEFAULT_LEFT_INDENT
    pHorizontalSpace = DEFAULT_HORIZONTAL_SPACE
    pLineHeight = DEFAULT_LINE_HEIGHT
    pFontFamily = DEFAULT_FONT_FAMILY
    pFontSize = DEFAULT_FONT_SIZE
End Sub



'[SETTERS]
Public Function setParent(value As XProcessItemsList) As xxProcessItemLine
    Set setParent = Me
    Set pParent = value
    Set pFrame = pParent.getItemLinesContainerFrame
End Function

Public Function setName(value As String) As xxProcessItemLine
    Set setName = Me
    pName = value
End Function

Public Function setIndex(value As Long) As xxProcessItemLine
    Set setIndex = Me
    pIndex = value
End Function

Public Function setItem(value As IProcessItem) As xxProcessItemLine
    Set setItem = Me
    Set pItem = value
End Function

Public Function setIsHidden(value As Boolean) As xxProcessItemLine
    Set setIsHidden = Me
    pIsHidden = value
    If pIsRendered Then Call updateVisibility
End Function


'[Font]
Public Function setFontSize(value As Single) As xxProcessItemLine
    Set setFontSize = Me
    pFontSize = value
    If pIsRendered Then
        Call refreshNameLabel
        Call refreshDescriptionLabel
    End If
End Function

Public Function setFontFamily(value As String) As xxProcessItemLine
    Set setFontFamily = Me
    pFontFamily = value
    If pIsRendered Then
        Call refreshNameLabel
        Call refreshDescriptionLabel
    End If
End Function


'[Position]
Public Function setTop(value As Single) As xxProcessItemLine
    Set setTop = Me
    pTop = value
    If pIsRendered Then Call arrangeControls
End Function

Public Function setBottom(value As Single) As xxProcessItemLine
    Set setBottom = Me
    pTop = value - pLineHeight
    If pIsRendered Then Call arrangeControls
End Function

Public Function setLeft(value As Single) As xxProcessItemLine
    Set setLeft = Me
    pLeft = value
    If pIsRendered Then Call arrangeControls
End Function

Public Function setRight(value As Single) As xxProcessItemLine
    Set setRight = Me
    pLeft = value - getWidth
    If pIsRendered Then Call arrangeControls
End Function

Public Function setHorizontalSpace(value As Single) As xxProcessItemLine
    Set setHorizontalSpace = Me
    pHorizontalSpace = value
    If pIsRendered Then Call arrangeControls
End Function




'[Size]
Public Function setWidth(value As Single) As xxProcessItemLine
    Set setWidth = Me
End Function

Public Function setHeight(value As Single) As xxProcessItemLine
    Set setHeight = Me
    pLineHeight = value
    If pIsRendered Then Call arrangeControls
End Function

Public Function setLineHeight(value As Single) As xxProcessItemLine
    Set setLineHeight = setHeight(value)
End Function



'[Statuses view]
Public Function setStatusesDictionary(value As Scripting.Dictionary) As xxProcessItemLine
    Set setStatusesDictionary = Me
    Set pStatusesDictionary = value
End Function




'[GETTERS]
Public Function getName() As String
    getName = pName
End Function

Public Function getIndex() As Integer
    getIndex = pIndex
End Function

Public Function getItem() As IProcessItem
    Set getItem = pItem
End Function


'[State]
Public Function isValid() As Boolean
    isValid = UI.Controls.isValid(pBackgroundLabel)
End Function

Public Function isHidden() As Boolean
    isHidden = pIsHidden
End Function


'[Position]
Public Function getLeft() As Single
    getLeft = pBackgroundLabel.left
End Function

Public Function getRight() As Single
    getRight = UI.Controls.getRight(pBackgroundLabel)
End Function

Public Function getTop() As Single
    getTop = pBackgroundLabel.top
End Function

Public Function getBottom() As Single
    If Not pDetailsBackground Is Nothing Then
        If pDetailsBackground.Visible Then
            getBottom = UI.Controls.getBottom(pDetailsBackground)
        End If
    End If
    If getBottom = 0 Then
        getBottom = UI.Controls.getBottom(pBackgroundLabel)
    End If
End Function


'[Size]
Public Function getWidth() As Single
    getWidth = pParent.getWidth
End Function

Public Function getHeight() As Single
    getHeight = pBackgroundLabel.height
    If pIsExpanded Then
        getHeight = getHeight + pDetailsBackground.height
    End If
End Function




'[API]
Public Sub render()
    If Not pIsRendered Then
        Call insertBackgroundLabel
        Call insertStatusIcon
        Call insertNameLabel
        Call insertReprocessItemButton
        Call insertPreviewItemButton
        Call insertExpandCollapseButton
        Call insertRemoveButton
        Call arrangeControls
        pIsRendered = True
    End If
End Sub

Public Sub move(Optional left As Variant, Optional top As Variant)
    If F.Variables.isOfNumericType(left) Then pLeft = left
    If F.Variables.isOfNumericType(top) Then pTop = top
    Call arrangeControls
End Sub

Public Sub resize(Optional width As Variant, Optional height As Variant)
    If height >= 0 Then pLineHeight = height
    Call arrangeControls
End Sub

Public Sub updateInfo(status As ItemProcessStatusEnum, message As String)
    Call updateStatus(status)
    Call updateDetailsDescription(message)
    Call arrangeControls
End Sub

Public Sub updateCaption()
    Call refreshNameLabel
End Sub



'[RENDERING]
Private Function getLabelName(controlType As ControlTypeEnum) As String
    Dim tag As String
    '----------------------------------------------------------------------------------------------------------
    tag = CLASS_NAME & "|" & pName & "|" & controlType
    getLabelName = "ctrl_" & VBA.left$(F.Crypto.SHA1HASH(tag), 25)
End Function


'[Inserting controls]
Private Sub insertBackgroundLabel()
    Set pBackgroundLabel = pFrame.Controls.Add(CREATE_LABEL_ID, getLabelName(ControlType_BackgroundLabel), True)
    Call refreshBackgroundLabel
End Sub

Private Sub insertStatusIcon()
    Set pStatusIcon = pFrame.Controls.Add(CREATE_LABEL_ID, getLabelName(ControlType_StatusIcon), True)
    Call refreshStatusIcon
End Sub

Private Sub insertNameLabel()
    Set pNameLabel = pFrame.Controls.Add(CREATE_LABEL_ID, getLabelName(ControlType_NameLabel), True)
    Call refreshNameLabel
End Sub

Private Sub insertRemoveButton()
    Set pRemoveButton = pFrame.Controls.Add(CREATE_LABEL_ID, getLabelName(ControlType_RemoveItemButton), True)
    Call refreshRemoveButton
End Sub

Private Sub insertPreviewItemButton()
    Set pPreviewItemButton = pFrame.Controls.Add(CREATE_LABEL_ID, getLabelName(ControlType_PreviewButton), True)
    Call refreshPreviewItemButton
End Sub

Private Sub insertReprocessItemButton()
    Set pReprocessButton = pFrame.Controls.Add(CREATE_LABEL_ID, getLabelName(ControlType_ReprocessItemButton), True)
    Call refreshReprocessItemButton
End Sub

Private Sub insertExpandCollapseButton()
    Set pExpandCollapseButton = pFrame.Controls.Add(CREATE_LABEL_ID, getLabelName(ControlType_ExpanderButton), True)
    Call refreshExpandCollapseButton
End Sub

Private Sub insertDetailsBackground()
    Set pDetailsBackground = pFrame.Controls.Add(CREATE_LABEL_ID, getLabelName(ControlType_DetailsBackground), False)
    Call refreshDetailsBackground
End Sub

Private Sub insertDescriptionLabel()
    Set pDescriptionLabel = pFrame.Controls.Add(CREATE_LABEL_ID, getLabelName(ControlType_DescriptionLabel), False)
    Call refreshDescriptionLabel
End Sub



'[Refreshing controls]
Private Sub refreshBackgroundLabel()
    With pBackgroundLabel
        .BackStyle = VBA.IIf(pStatus = ItemProcessStatus_Unknown, fmBackStyleTransparent, fmBackStyleOpaque)
        .BackColor = getBackColorByStatus(pStatus)
        .BorderStyle = fmBorderStyleNone
    End With
End Sub

Private Sub refreshStatusIcon()
    With pStatusIcon
        .BackStyle = fmBackStyleTransparent
        .BorderStyle = fmBorderStyleNone
        .Picture = getPictureByStatus(pStatus)
    End With
End Sub

Private Sub refreshNameLabel()
    With pNameLabel
        .BackStyle = fmBackStyleTransparent
        .Font.name = pFontFamily
        .Font.size = UI.Normalize.FontSize(pFontSize)
        .TextAlign = fmTextAlignLeft
        .caption = getCaptionForDisplay(pItem.getCaption)
    End With
End Sub

Private Function getCaptionForDisplay(fullName As String) As String
    Dim maxLength As Long:              maxLength = pParent.getWidth * 0.15
    '------------------------------------------------------------------------------------------------------
    If VBA.Len(fullName) > maxLength Then
        If F.files.isValidPath(fullName) Then
            getCaptionForDisplay = getCaptionForDisplay("...\" & F.files.getNameFromUncPath(fullName))
        Else
            getCaptionForDisplay = VBA.left$(fullName, maxLength - 3) & "..."
        End If
    Else
        getCaptionForDisplay = fullName
    End If
End Function

Private Sub refreshExpandCollapseButton()
    With pExpandCollapseButton
        .BackStyle = fmBackStyleTransparent
        '.BorderStyle = fmBorderStyleSingle
        .MousePointer = fmMousePointerCustom
        .MouseIcon = UI.Icons.getMousePointer
        .Visible = False
    End With
End Sub

Private Sub refreshDetailsBackground()
    With pDetailsBackground
        .BackStyle = fmBackStyleOpaque
        .BackColor = getDetailsBackColorByStatus(pStatus)
    End With
End Sub

Private Sub refreshDescriptionLabel()
    With pDescriptionLabel
        .BackStyle = fmBackStyleTransparent
        .Font.name = pFontFamily
        .Font.size = pFontSize
        .TextAlign = fmTextAlignLeft
        .caption = pDescription
    End With
End Sub

Private Sub refreshPreviewItemButton()
    With pPreviewItemButton
        .BackStyle = fmBackStyleTransparent
        .MousePointer = fmMousePointerCustom
        .MouseIcon = UI.Icons.getMousePointer
        .ControlTipText = MsgService.getText(OPEN_ITEM_IN_SEPARATE_WINDOW)
        .Visible = False
    End With
End Sub

Private Sub refreshReprocessItemButton()
    With pReprocessButton
        .BackStyle = fmBackStyleTransparent
        .MousePointer = fmMousePointerCustom
        .MouseIcon = UI.Icons.getMousePointer
        .ControlTipText = MsgService.getText(REPROCESS_ITEM_TOOLTIP)
        .Visible = False
    End With
End Sub

Private Sub refreshRemoveButton()
    With pRemoveButton
        .BackStyle = fmBackStyleTransparent
        .MousePointer = fmMousePointerCustom
        .MouseIcon = UI.Icons.getMousePointer
        .ControlTipText = MsgService.getText(REMOVE_ITEM_TOOLTIP)
        .Visible = False
    End With
End Sub


'[Conditional UI properties]
Private Function getPropertyValueByStatus(propertyTag As String, status As ItemProcessStatusEnum) As Variant
    Dim dict As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    Set dict = F.dictionaries.GetObject(pStatusesDictionary, status)
    If Not dict Is Nothing Then
        getPropertyValueByStatus = F.dictionaries.getItem(dict, propertyTag)
    End If
End Function

Private Function getPropertyObjectByStatus(propertyTag As String, status As ItemProcessStatusEnum) As Object
    Dim dict As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    Set dict = F.dictionaries.GetObject(pStatusesDictionary, status)
    If Not dict Is Nothing Then
        Set getPropertyObjectByStatus = F.dictionaries.GetObject(dict, propertyTag)
    End If
End Function

Private Function getBackColorByStatus(status As ItemProcessStatusEnum) As Long
    getBackColorByStatus = getPropertyValueByStatus(TAG_MAIN_BACKGROUND, status)
End Function

Private Function getPictureByStatus(status As ItemProcessStatusEnum) As Variant
    Set getPictureByStatus = getPropertyObjectByStatus(TAG_ICON_STATUS, status)
End Function

Private Function getExpandPicture(status As ItemProcessStatusEnum) As Variant
    If pIsExpanded Then
        Set getExpandPicture = getPropertyObjectByStatus(TAG_ICON_COLLAPSE, status)
    Else
        Set getExpandPicture = getPropertyObjectByStatus(TAG_ICON_EXPAND, status)
    End If
End Function

Private Function getPreviewPicture(status As ItemProcessStatusEnum) As Variant
    Set getPreviewPicture = getPropertyObjectByStatus(TAG_ICON_PREVIEW, status)
End Function

Private Function getReprocessItemPicture(status As ItemProcessStatusEnum) As Variant
    Set getReprocessItemPicture = getPropertyObjectByStatus(TAG_ICON_REPROCESS, status)
End Function

Private Function getRemoveItemPicture(status As ItemProcessStatusEnum) As Variant
    Set getRemoveItemPicture = getPropertyObjectByStatus(TAG_ICON_REMOVE, status)
End Function

Private Function getDetailsBackColorByStatus(status As ItemProcessStatusEnum) As Long
    getDetailsBackColorByStatus = getPropertyValueByStatus(TAG_DETAILS_BACKGROUND, status)
End Function




'[Size & location]
Private Sub arrangeControls()
    Dim left As Single
    Dim top As Single
    Dim buttons As VBA.Collection
    Dim button As MSForms.label
    '------------------------------------------------------------------------------------------------------
    
    On Error Resume Next
    Call resizeControls
    Call UI.Controls.locate(pBackgroundLabel, pLeft, pTop)
    Call UI.Controls.locate(pStatusIcon, pLeft + STATUS_ICON_LEFT_OFFSET, pTop + (pLineHeight - pStatusIcon.height) / 2)
    Call UI.Controls.locate(pNameLabel, pLeft + NAME_LABEL_TO_ICON_OFFSET + STATUS_ICON_WIDTH + STATUS_ICON_LEFT_OFFSET, _
                                        pTop + (pLineHeight - pNameLabel.height) / 2)
    
    Set buttons = F.Collections.Create(pExpandCollapseButton, pReprocessButton, pPreviewItemButton, pRemoveButton)
    left = pFrame.width - EXPAND_ICON_WIDTH - EXPAND_ICON_RIGHT
    top = EXPAND_ICON_TOP + pTop
    For Each button In buttons
        Call locateFunctionButton(button, left, top)
    Next button
    
    If pIsExpanded Then
        Call UI.Controls.locate(pDetailsBackground, pLeft, pTop + pLineHeight)
        Call UI.Controls.locate(pDescriptionLabel, pLeft + STATUS_ICON_LEFT_OFFSET, pTop + pLineHeight)
    End If

    '[To avoid bug with smaller font-size]
    Call adjustControlsTop

End Sub

Private Sub locateFunctionButton(button As MSForms.label, ByRef left As Single, top As Single)
    If UI.Controls.isValid(button) Then
        If button.Picture Is Nothing Then
            button.Visible = False
        ElseIf button.Picture.width = 0 Then
            button.Visible = False
        Else
            button.Visible = True
            Call UI.Controls.locate(button, left, top)
            left = left - button.width - ICONS_SPACE
        End If
    End If
End Sub

Private Sub adjustControlsTop()
    Const ICON_SIZE As Single = 10.5
    '------------------------------------------------------------------------------------------------------
    Dim top As Single
    '------------------------------------------------------------------------------------------------------
    
    On Error Resume Next
    
    top = pBackgroundLabel.top + (pLineHeight - ICON_SIZE) / 2
    
    pStatusIcon.top = top
    pPreviewItemButton.top = top
    pExpandCollapseButton.top = top
    pRemoveButton.top = top
    pReprocessButton.top = top
    
'    If (pBackgroundLabel.top + pLineHeight - pPreviewItemButton.top - pPreviewItemButton.height) < 2 * PIXEL_SIZE Then
'        pStatusIcon.top = pStatusIcon.top - Y_OFFSET
'        pPreviewItemButton.top = pPreviewItemButton.top - Y_OFFSET
'        pExpandCollapseButton.top = pExpandCollapseButton.top - Y_OFFSET
'        pRemoveButton.top = pRemoveButton.top - Y_OFFSET
'        pReprocessButton.top = pReprocessButton.top - Y_OFFSET
'    End If
    
End Sub


Private Sub resizeControls()
    Call UI.Controls.resize(pBackgroundLabel, getWidth, pLineHeight)
    Call UI.Controls.resize(pStatusIcon, STATUS_ICON_WIDTH, STATUS_ICON_HEIGHT)
    Call resizeNameLabel
    Call UI.Controls.resize(pReprocessButton, EXPAND_ICON_WIDTH, EXPAND_ICON_HEIGHT)
    Call UI.Controls.resize(pPreviewItemButton, EXPAND_ICON_WIDTH, EXPAND_ICON_HEIGHT)
    Call UI.Controls.resize(pExpandCollapseButton, EXPAND_ICON_WIDTH, EXPAND_ICON_HEIGHT)
    Call UI.Controls.resize(pRemoveButton, EXPAND_ICON_WIDTH, EXPAND_ICON_HEIGHT)
    If pIsExpanded Then Call resizeDetailsBackgroundSize
End Sub

Private Sub resizeNameLabel()
    Dim caption As String
    '----------------------------------------------------------------------------------------------------------
    
    With pNameLabel
        caption = .caption
        .caption = "tj"
        .AutoSize = True
        .AutoSize = False
        .width = getWidth - STATUS_ICON_WIDTH - NAME_LABEL_TO_ICON_OFFSET
        .caption = caption
        .height = .height + 0.05
    End With
    
End Sub

Private Sub resizeDetailsBackgroundSize()
    Dim height As Single
    '----------------------------------------------------------------------------------------------------------
    On Error Resume Next
    With pDescriptionLabel
        .width = getWidth - 2 * STATUS_ICON_LEFT_OFFSET
        .AutoSize = True
        .AutoSize = False
        .width = getWidth - 2 * STATUS_ICON_LEFT_OFFSET
    End With
    height = UI.Controls.getBottom(pDescriptionLabel) - pDetailsBackground.top + DESCRIPTION_BOTTOM_PADDING
    With pDetailsBackground
        .height = height
        .width = getWidth
    End With
End Sub


'[Clearing]
Public Sub clear()
    On Error Resume Next
    With pFrame.Controls
        Call .Remove(getLabelName(ControlType_BackgroundLabel))
        Call .Remove(getLabelName(ControlType_StatusIcon))
        Call .Remove(getLabelName(ControlType_NameLabel))
        Call .Remove(getLabelName(ControlType_RecordsCounterLabel))
        Call .Remove(getLabelName(ControlType_ExpanderButton))
        Call .Remove(getLabelName(ControlType_RemoveItemButton))
        Call .Remove(getLabelName(ControlType_ReprocessItemButton))
        Call .Remove(getLabelName(ControlType_PreviewButton))
        Call .Remove(getLabelName(ControlType_DescriptionLabel))
        Call .Remove(getLabelName(ControlType_DetailsBackground))
    End With
End Sub





'[UPDATING INFO]
Private Sub updateStatus(status As ItemProcessStatusEnum)
    pStatus = status
    pStatusIcon.Picture = getPictureByStatus(pStatus)
    pBackgroundLabel.BackStyle = VBA.IIf(pStatus = ItemProcessStatus_Unknown, fmBackStyleTransparent, fmBackStyleOpaque)
    pBackgroundLabel.BackColor = getBackColorByStatus(pStatus)
    With pExpandCollapseButton
        .Visible = True
        .Picture = getExpandPicture(pStatus)
    End With
    With pPreviewItemButton
        .Visible = True
        .Picture = getPreviewPicture(pStatus)
    End With
    With pReprocessButton
        .Visible = True
        .Picture = getReprocessItemPicture(pStatus)
    End With
    With pRemoveButton
        .Visible = True
        .Picture = getRemoveItemPicture(pStatus)
    End With
End Sub

Private Sub updateDetailsDescription(message As String)
    Dim heightBefore As Single
    '----------------------------------------------------------------------------------------------------------
    
    pDescription = message
    If UI.Controls.isValid(pDescriptionLabel) Then
        pDescriptionLabel.caption = pDescription
        pDetailsBackground.BackColor = getDetailsBackColorByStatus(pStatus)
    End If
    
    'If item is not expanded, even if height of its description change, it doesn't affect controls layout.
    If pIsExpanded Then
        heightBefore = getHeight
        Call arrangeControls
        'There is no point to rearrange lines if height of this item didn't change.
        If getHeight <> heightBefore Then
            Call pParent.arrangeItemLines(Me)
        End If
    End If

End Sub





'[UI DETAILS FRAME]
Private Sub pExpandCollapseButton_Click()
    Call ErrorManager.clear
    If pIsExpanded Then
        Call collapseDetailsPanel
    Else
        Call expandDetailsPanel
    End If
    
    Call pParent.arrangeItemLines(Me)
    Call showHideDetailsPanel(pIsExpanded)
    pExpandCollapseButton.Picture = getExpandPicture(pStatus)
    Call pParent.triggerAfterExpandCollapseEvent(Me)
End Sub

Private Sub collapseDetailsPanel()
    pIsExpanded = False
    Call arrangeControls
End Sub

Private Sub expandDetailsPanel()
    If Not pIsDetailsPanelRendered Then
        Call insertDetailsPanel
    End If
    pIsExpanded = True
    Call arrangeControls
End Sub

Private Sub insertDetailsPanel()
    Call insertDetailsBackground
    Call insertDescriptionLabel
    pIsDetailsPanelRendered = True
End Sub

Private Sub showHideDetailsPanel(value As Boolean)
    pDetailsBackground.Visible = value
    pDescriptionLabel.Visible = value
End Sub

Private Sub updateVisibility()
    On Error Resume Next
    pBackgroundLabel.Visible = Not pIsHidden
    pStatusIcon.Visible = Not pIsHidden
    pNameLabel.Visible = Not pIsHidden
    pRecordsCounterLabel.Visible = Not pIsHidden
    pReprocessButton.Visible = Not pIsHidden
    pPreviewItemButton.Visible = Not pIsHidden
    pExpandCollapseButton.Visible = Not pIsHidden
    pRemoveButton.Visible = Not pIsHidden
    pDetailsBackground.Visible = pIsExpanded And Not pIsHidden
    pDescriptionLabel.Visible = pIsExpanded And Not pIsHidden
End Sub



'[ACTIONS]
Private Sub pPreviewItemButton_Click()
    Call ErrorManager.clear
    Call pItem.preview
    'Call pParent.triggerAfterPreviewItemEvent
End Sub

Private Sub pReprocessButton_Click()
    Call ErrorManager.clear
    Call pParent.triggerAfterReprocessItemEvent(Me)
End Sub

Private Sub pRemoveButton_Click()
    Call ErrorManager.clear
    Call pParent.getProcessor.removeItem(pName)
End Sub
