# Programming for the Single Document Interface in Excel 2013

## Business Description
Learn about programming considerations for the Single Document Interface in Microsoft Excel 2013.

## Behavior
Applies to:Excel 2013

## Example Usage
```vba
Private Sub UserForm_Layout()
    Static fSetModal As Boolean
    If fSetModal = False Then
        fSetModal = True
        Me.Hide
        Me.Show 1
    End If
End Sub
```