Attribute VB_Name = "modStyles"
Option Explicit

Private Const CLASS_NAME As String = "modStyles"
'----------------------------------------------------------------------------------------------------------
Public Const DEFAULT_OUTSIDE_BORDER_COLOR As Long = 8421504
Public Const DEFAULT_INSIDE_BORDER_COLOR As Long = 14277081
Public Const DEFAULT_BORDER_WEIGHT As Long = xlMedium
Public Const DEFAULT_BORDER_STYLE As Long = xlContinuous
Public Const BORDER_WEIGHT_TAG As String = "weight"
Public Const BORDER_COLOR_TAG As String = "color"
Public Const BORDER_STYLE_TAG As String = "style"
'----------------------------------------------------------------------------------------------------------



'[ALIGNMENTS]
Public Function convertAlignTextToEnum(value As String, Optional form As Boolean = False) As Variant
    Select Case VBA.LCase$(value)
        Case "center":  convertAlignTextToEnum = VBA.IIf(form, fmTextAlignCenter, xlCenter)
        Case "left":    convertAlignTextToEnum = VBA.IIf(form, fmTextAlignLeft, xlLeft)
        Case "right":   convertAlignTextToEnum = VBA.IIf(form, fmTextAlignRight, xlRight)
        Case "top":     convertAlignTextToEnum = xlTop
        Case "bottom":  convertAlignTextToEnum = xlBottom
        Case Else:      convertAlignTextToEnum = xlCenter
    End Select
End Function


'[RANGE BORDERS]
Public Function convertBorderIndexNameToEnum(value As String) As XlBordersIndex
    Select Case VBA.LCase$(value)
        Case "left":                convertBorderIndexNameToEnum = xlEdgeLeft
        Case "right":               convertBorderIndexNameToEnum = xlEdgeRight
        Case "top":                 convertBorderIndexNameToEnum = xlEdgeTop
        Case "bottom":              convertBorderIndexNameToEnum = xlEdgeBottom
        Case "inside-vertical":     convertBorderIndexNameToEnum = xlInsideVertical
        Case "inside-horizontal":   convertBorderIndexNameToEnum = xlInsideHorizontal
    End Select
End Function

Public Function convertBorderIndexEnumToName(value As XlBordersIndex) As String
    Select Case VBA.LCase$(value)
        Case xlEdgeLeft:            convertBorderIndexEnumToName = "left"
        Case xlEdgeRight:           convertBorderIndexEnumToName = "right"
        Case xlEdgeTop:             convertBorderIndexEnumToName = "top"
        Case xlEdgeBottom:          convertBorderIndexEnumToName = "bottom"
        Case xlInsideVertical:      convertBorderIndexEnumToName = "inside-vertical"
        Case xlInsideHorizontal:    convertBorderIndexEnumToName = "inside-horizontal"
    End Select
End Function

Public Function convertBorderStyleNameToEnum(value As String) As XlLineStyle
    Select Case VBA.LCase$(value)
        Case "continuous":          convertBorderStyleNameToEnum = xlContinuous
        Case "none":                convertBorderStyleNameToEnum = xlLineStyleNone
        Case Else:                  convertBorderStyleNameToEnum = xlContinuous
    End Select
End Function

Public Function convertBorderWeightNameToEnum(value As String) As XlBorderWeight
    Select Case VBA.LCase$(value)
        Case "thin":                convertBorderWeightNameToEnum = xlThin
        Case "medium":              convertBorderWeightNameToEnum = xlMedium
        Case "thick":               convertBorderWeightNameToEnum = xlThick
        Case Else:                  convertBorderWeightNameToEnum = xlThin
    End Select
End Function

Public Function isInsideBorder(borderIndex As XlBordersIndex)
    If borderIndex = xlInsideHorizontal Then
        isInsideBorder = True
    ElseIf borderIndex = xlInsideVertical Then
        isInsideBorder = True
    End If
End Function

Public Function isOutsideBorder(borderIndex As XlBordersIndex)
    isOutsideBorder = Not isInsideBorder(borderIndex)
End Function

Public Function convertRgbToLong(ByVal text As String) As Long
    convertRgbToLong = f.colors.convertCssRgbToLong(text)
End Function

Public Function getDefaultBordersStylesDictionary() As Scripting.Dictionary
    Static dict As Scripting.Dictionary
    '----------------------------------------------------------------------------------------------------------
    Dim dictInside As Scripting.Dictionary
    Dim dictOutside As Scripting.Dictionary
    '----------------------------------------------------------------------------------------------------------
    
    If dict Is Nothing Then
        Set dictInside = f.dictionaries.createWithItems(False, _
                            KeyValue(BORDER_WEIGHT_TAG, "thin"), _
                            KeyValue(BORDER_STYLE_TAG, "continuous"), _
                            KeyValue(BORDER_COLOR_TAG, DEFAULT_INSIDE_BORDER_COLOR))
        Set dictOutside = f.dictionaries.createWithItems(False, _
                            KeyValue(BORDER_WEIGHT_TAG, "medium"), _
                            KeyValue(BORDER_STYLE_TAG, "continuous"), _
                            KeyValue(BORDER_COLOR_TAG, DEFAULT_OUTSIDE_BORDER_COLOR))
        Set dict = f.dictionaries.Create(False)
        With dict
            Call .Add("left", dictOutside)
            Call .Add("right", dictOutside)
            Call .Add("top", dictOutside)
            Call .Add("bottom", dictOutside)
            Call .Add("inside-vertical", dictInside)
            Call .Add("inside-horizontal", dictInside)
        End With
    End If
    
    Set getDefaultBordersStylesDictionary = dict
    
End Function




'[FORMAT CONDITIONS]
Public Function convertFormatConditionTypeToEnum(value As String) As XlFormatConditionType
    Select Case VBA.LCase$(value)
        Case "expression", "formula":
                                    convertFormatConditionTypeToEnum = xlExpression
        Case "cellvalue":           convertFormatConditionTypeToEnum = xlCellValue
        Case "iconsets":            convertFormatConditionTypeToEnum = xlIconSets
        Case Else:                  convertFormatConditionTypeToEnum = xlExpression
    End Select
End Function

Public Function convertFormatConditionOperatorToEnum(value As String) As XlFormatConditionOperator
    Select Case VBA.LCase$(value)
        Case "greater":             convertFormatConditionOperatorToEnum = xlGreater
        Case "equal":               convertFormatConditionOperatorToEnum = xlEqual
        Case "less":                convertFormatConditionOperatorToEnum = xlLess
        Case Else:                  convertFormatConditionOperatorToEnum = xlEqual
    End Select
End Function
