# QueryTable WebDisableRedirections Property

## Business Description
True if Web query redirections are disabled for a QueryTable object. The default value is False. Read/write Boolean.

## Behavior
Trueif Web query redirections are disabled for aQueryTableobject. The default value isFalse. Read/writeBoolean.

## Example Usage
```vba
Sub CheckWebQuerySetting() 
 Dim wksSheet As Worksheet 
 Set wksSheet = Application.ActiveSheet 
 MsgBox wksSheet.QueryTables(1).WebDisableRedirectionsEnd Sub
```