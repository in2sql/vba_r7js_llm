# Application.SheetPivotTableBeforeDiscardChanges event (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Parameters
- **Sh**: Required
- **TargetPivotTable**: Required
- **ValueChangeStart**: Required
- **ValueChangeEnd**: Required

## Return Value
Nothing

## Remarks
Occurs immediately before Excel executes a ROLLBACK TRANSACTION statement against the OLAP data source, if a transaction is still active, and then discards all edited values in the PivotTable after the user has chosen to discard changes.

## Example
No VBA example available.
