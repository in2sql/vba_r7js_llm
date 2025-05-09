# Page Object

## Business Description
Represents a page in a workbook. Use the PageSetup object and the related methods and properties for programmatically defining page layout in a workbook.

## Behavior
Represents a page in a workbook. Use thePageSetupobject and the related methods and properties for programmatically defining page layout in a workbook.

## Example Usage
```vba
Dim objPage As Page 
 
Set objPage = ActiveWorkbook.ActiveWindow _ 
 .Panes(1).Pages.Item(1)
```