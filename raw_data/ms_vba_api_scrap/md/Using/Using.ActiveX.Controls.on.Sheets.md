# Using ActiveX Controls on Sheets

## Business Description
This topic covers specific information about using ActiveX controls on worksheets and chart sheets. For general information on adding and working with controls, see Using ActiveX Controls on a Document and Creating a Custom Dialog Box.

## Behavior
This topic covers specific information about using ActiveX controls on worksheets and chart sheets. For general information on adding and working with controls, seeUsing ActiveX Controls on a DocumentandCreating a Custom Dialog Box.

## Example Usage
```vba
Set t = Sheet1.CommandButton1.TopLeftCell
With ActiveWindow
    .ScrollRow = t.Row
    .ScrollColumn = t.Column
End With
```