# Application.SheetPivotTableAfterValueChange event (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Parameters
- **Sh**: Required
- **TargetPivotTable**: Required
- **TargetRange**: Required

## Return Value
Nothing

## Remarks
The PivotTableAfterValueChange event does not occur under any conditions other than editing or recalculating cells. For example, it will not occur when the PivotTable is refreshed, sorted, filtered, or drilled down on, even though those operations move cells and potentially retrieve new values from the OLAP data source.

## Example
No VBA example available.
