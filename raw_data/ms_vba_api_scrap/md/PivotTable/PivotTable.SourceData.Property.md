# PivotTable SourceData Property

## Business Description
Returns the data source for the PivotTable report, as shown in the following table. Read-write Variant.

## Behavior
Returns the data source for the PivotTable report, as shown in the following table. Read-writeVariant.

## Example Usage
```vba
Set newSheet = ActiveWorkbook.Worksheets.Add 
sdArray = Worksheets("Sheet1").UsedRange.PivotTable.SourceDataFor i = LBound(sdArray) To UBound(sdArray) 
 newSheet.Cells(i, 1) = sdArray(i) 
Next i
```