# Worksheet FollowHyperlink Event

## Business Description
Occurs when you click any hyperlink on a worksheet. For application- and workbook-level events, see the SheetFollowHyperlink event and SheetFollowHyperlink event.

## Behavior
Occurs when you click any hyperlink on a worksheet. For application- and workbook-level events, see theSheetFollowHyperlinkevent andSheetFollowHyperlinkevent.

## Example Usage
```vba
Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink) 
    With UserForm1 
        .ListBox1.AddItem Target.Address 
        .Show 
    End With 
End Sub
```