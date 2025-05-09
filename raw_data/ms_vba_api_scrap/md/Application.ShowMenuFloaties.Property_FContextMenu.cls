VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FContextMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "FContextMenu"
'----------------------------------------------------------------------------------------------------------
Private Const MENU_ITEM_PREFIX As String = "custom_"
Private Const CONTEXT_MENU_COMMAND_BAR_NAME As String = "Cell"
'----------------------------------------------------------------------------------------------------------


Public Sub addContextMenuItemFromDto(ByRef dto As Scripting.Dictionary)
    Dim methodBook As Excel.Workbook
    Dim methodName As String
    Dim param As String
    Dim buttonCaption As String
    Dim buttonTag As String
    Dim faceId As Long
    Dim separateFromNext As Boolean
    '------------------------------------------------------------------------------------------------------
    
    Set methodBook = f.dictionaries.getObject(dto, CONTEXT_MENU_TAG_METHOD_BOOK)
    methodName = f.Reflection.getFunctionFullName(f.dictionaries.getItem(dto, CONTEXT_MENU_TAG_METHOD_NAME), methodBook)
    param = f.dictionaries.getItem(dto, CONTEXT_MENU_TAG_PARAM)
    If VBA.Len(param) Then methodName = methodName & "(""" & param & """)"
    
    buttonCaption = MsgService.getText(f.dictionaries.getItem(dto, CONTEXT_MENU_TAG_CAPTION))
    buttonTag = f.dictionaries.getItem(dto, CONTEXT_MENU_TAG_TAG)
    faceId = f.dictionaries.getItem(dto, CONTEXT_MENU_TAG_FACE_ID)
    separateFromNext = f.dictionaries.getItem(dto, CONTEXT_MENU_TAG_SEPARATOR)
    
    Call addContextMenuItem(methodName, buttonCaption, buttonTag, faceId, 1, separateFromNext)

End Sub


Public Sub addContextMenuItem(methodName As String, buttonCaption As String, buttonTag As String, _
                            Optional faceId As Long, Optional before As Long = 1, _
                            Optional separateFromNext As Boolean = False)
    Dim ContextMenu As Object:                      Set ContextMenu = getContextMenu
    Dim contextMenuItem As Object
    '------------------------------------------------------------------------------------------------------
    
    Set contextMenuItem = getContextMenuItem(buttonTag)
    If Not contextMenuItem Is Nothing Then
        contextMenuItem.OnAction = methodName
    Else
        With ContextMenu.Controls.Add(Type:=msoControlButton, before:=before)
            .OnAction = methodName
            .faceId = faceId
            .caption = buttonCaption
            .tag = buttonTag
        End With
        
        If separateFromNext Then
            On Error Resume Next
            ContextMenu.Controls(before + 1).BeginGroup = True
        End If
    End If
    
End Sub

Private Function getContextMenu() As Object
    Set getContextMenu = Excel.Application.CommandBars(CONTEXT_MENU_COMMAND_BAR_NAME)
End Function

Private Function getContextMenuItem(tag As String) As Object
    Dim ContextMenu As Object:                      Set ContextMenu = getContextMenu
    Dim ctrl As Object
    '------------------------------------------------------------------------------------------------------
    For Each ctrl In ContextMenu.Controls
        If ctrl.tag = tag Then
            Set getContextMenuItem = ctrl
            Exit For
        End If
    Next ctrl
End Function

Public Sub clearContextMenu()
    Dim ContextMenu As Object:                      Set ContextMenu = getContextMenu
    Dim item As Object
    '------------------------------------------------------------------------------------------------------

    If Not ContextMenu Is Nothing Then
        For Each item In ContextMenu.Controls
            Call item.delete
        Next item
    End If
    
    Application.ShowMenuFloaties = True

End Sub

Public Sub resetContextMenu()
    Dim ContextMenu As Object:                      Set ContextMenu = getContextMenu
    '------------------------------------------------------------------------------------------------------
    
    If Not ContextMenu Is Nothing Then
        Call ContextMenu.reset
    End If
    
    Application.ShowMenuFloaties = False
    
End Sub

Public Sub clearContextMenuFromCustomItems(Optional projectPrefix As String)
    Call deleteContextMenuItemsByRegex("^" & MENU_ITEM_PREFIX, "^" & projectPrefix)
End Sub

Private Sub deleteContextMenuItemsByRegex(ParamArray items() As Variant)
    Dim ContextMenu As Object:                      Set ContextMenu = getContextMenu
    Dim objRegex As Object
    Dim ctrl As Object
    Dim varItem As Variant
    '------------------------------------------------------------------------------------------------------

    Set objRegex = f.regex.Create
    
    For Each varItem In items
        objRegex.pattern = VBA.CStr(varItem)
        For Each ctrl In ContextMenu.Controls
            If objRegex.test(ctrl.tag) Then
                Call ctrl.delete
            End If
        Next ctrl
    Next varItem

    If ContextMenu.Controls.count Then
        ContextMenu.Controls(1).BeginGroup = False
    End If

End Sub
