# Workbook SheetFollowHyperlink Event

## Business Description
Occurs when you click any hyperlink in Microsoft Excel. For worksheet-level events, see the Help topic for the FollowHyperlink event.

## Behavior
Occurs when you click any hyperlink in Microsoft Excel. For worksheet-level events, see the Help topic for theFollowHyperlinkevent.

## Example Usage
```vba
Private Sub Workbook_SheetFollowHyperlink(ByVal Sh as Object, _ 
 ByVal Target As Hyperlink) 
 UserForm1.ListBox1.AddItem Sh.Name & ":" & Target.Address 
 UserForm1.Show 
End Sub
```