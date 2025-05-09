# Workbook AutoUpdateSaveChanges Property

## Business Description
True if current changes to the shared workbook are posted to other users whenever the workbook is automatically updated. False if changes aren't posted (this workbook is still synchronized with changes made by other users). The default value is True.

## Behavior
Trueif current changes to the shared workbook are posted to other users whenever the workbook is automatically updated.Falseif changes aren't posted (this workbook is still synchronized with changes made by other users). The default value isTrue. Read/writeBoolean.

## Example Usage
```vba
ActiveWorkbook.AutoUpdateSaveChanges= True
```