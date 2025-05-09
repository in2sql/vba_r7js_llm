# Initializing Control Properties

## Business Description
You can initialize controls at run time by using Visual Basic code in a macro. For example, you could fill a list box, set text values, or set option buttons.

## Behavior
You can initializecontrolsat run time by using Visual Basic code in a macro. For example, you could fill a list box, set text values, or set option buttons.

## Example Usage
```vba
Private Sub GetUserName() 
 With UserForm1 
 .lstRegions.AddItem "North" 
 .lstRegions.AddItem "South" 
 .lstRegions.AddItem "East" 
 .lstRegions.AddItem "West" 
 .txtSalesPersonID.Text = "00000" 
 .Show 
 ' ... 
 End With 
End Sub
```