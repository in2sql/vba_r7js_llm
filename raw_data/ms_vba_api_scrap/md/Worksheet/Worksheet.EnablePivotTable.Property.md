# Worksheet EnablePivotTable Property

## Business Description
True if PivotTable controls and actions are enabled when user-interface-only protection is turned on. Read/write Boolean.

## Behavior
Trueif PivotTable controls and actions are enabled when user-interface-only protection is turned on. Read/writeBoolean.

## Example Usage
```vba
ActiveSheet.EnablePivotTable= True 
ActiveSheet.Protect contents:=True, userInterfaceOnly:=True
```