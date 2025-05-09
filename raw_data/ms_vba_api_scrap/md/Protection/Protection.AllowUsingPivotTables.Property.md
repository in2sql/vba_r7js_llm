# Protection AllowUsingPivotTables Property

## Business Description
Returns True if the user is allowed to manipulate pivot tables on a protected worksheet. Read-only Boolean.

## Behavior
ReturnsTrueif the user is allowed to manipulate pivot tables on a protected worksheet. Read-onlyBoolean.

## Example Usage
```vba
Sub ProtectionOptions() 
 
 ActiveSheet.Unprotect 
 
 ' Allow pivot tables to be manipulated on a protected worksheet. 
 If ActiveSheet.Protection.Allow UsingPivotTables= False Then 
 ActiveSheet.Protect AllowUsingPivotTables:=True 
 End If 
 
 MsgBox "Pivot tables can be manipulated on the protected worksheet." 
 
End Sub
```