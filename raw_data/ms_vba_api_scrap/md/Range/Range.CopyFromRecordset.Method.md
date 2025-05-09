# Range CopyFromRecordset Method

## Business Description
Copies the contents of an ADO or DAO Recordset object onto a worksheet, beginning at the upper-left corner of the specified range. If the Recordset object contains fields with OLE objects in them, this method fails.

## Behavior
Copies the contents of an ADO or DAORecordsetobject onto a worksheet, beginning at the upper-left corner of the specified range. If theRecordsetobject contains fields with OLE objects in them, this method fails.

## Example Usage
```vba
For iCols = 0 to rs.Fields.Count - 1 
 ws.Cells(1, iCols + 1).Value = rs.Fields(iCols).Name 
Next 
ws.Range(ws.Cells(1, 1), _ 
 ws.Cells(1, rs.Fields.Count)).Font.Bold = True 
ws.Range("A2").CopyFromRecordsetrs
```