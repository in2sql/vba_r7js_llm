# Workbook UpdateFromFile Method

## Business Description
Updates a read-only workbook from the saved disk version of the workbook if the disk version is more recent than the copy of the workbook that is loaded in memory.

## Behavior
Updates a read-only workbook from the saved disk version of the workbook if the disk version is more recent than the copy of the workbook that is loaded in memory. If the disk copy hasn't changed since the workbook was loaded, the in-memory copy of the workbook isn't reloaded.

## Example Usage
```vba
ActiveWorkbook.UpdateFromFile
```