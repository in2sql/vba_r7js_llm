# Worksheet BeforeRightClick Event

## Business Description
Occurs when a worksheet is right-clicked, before the default right-click action.

## Behavior
Occurs when a worksheet is right-clicked, before the default right-click action.

## Example Usage
```vba
Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, _ 
 Cancel As Boolean) 
 Dim icbc As Object 
 For Each icbc In Application.CommandBars("cell").Controls 
 If icbc.Tag = "brccm" Then icbc.Delete 
 Next icbc 
 If Not Application.Intersect(Target, Range("b1:b10")) _ 
 Is Nothing Then 
 With Application.CommandBars("cell").Controls _ 
 .Add(Type:=msoControlButton, before:=6, _ 
 temporary:=True) 
 .Caption = "New Context Menu Item" 
 .OnAction = "MyMacro" 
 .Tag = "brccm" 
 End With 
 End If 
End Sub
```