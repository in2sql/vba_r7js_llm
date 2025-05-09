# Worksheet PivotTableWizard Method

## Business Description
Creates a new PivotTable report. This method doesn't display the PivotTable Wizard. This method isn't available for OLE DB data sources. Use the Add method to add a PivotTable cache, and then create a PivotTable report based on the cache.

## Behavior
Creates a new PivotTable report. This method doesn't display the PivotTable Wizard. This method isn't available for OLE DB data sources. Use theAddmethod to add a PivotTable cache, and then create a PivotTable report based on the cache.

## Example Usage
```vba
ActiveSheet.PivotTableWizard xlDatabase, Range("A1:C100")
```