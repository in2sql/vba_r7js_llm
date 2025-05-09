# Workbook PasswordEncryptionFileProperties Property

## Business Description
True if Microsoft Excel encrypts file properties for the specified password-protected workbook. Read-only Boolean.

## Behavior
Trueif Microsoft Excel encrypts file properties for the specified password-protected workbook. Read-onlyBoolean.

## Example Usage
```vba
Sub SetPasswordOptions() 
 
 With ActiveWorkbook 
 If .PasswordEncryptionFileProperties= False Then 
 .SetPasswordEncryptionOptions _ 
 PasswordEncryptionProvider:="Microsoft RSA SChannel Cryptographic Provider", _ 
 PasswordEncryptionAlgorithm:="RC4", _ 
 PasswordEncryptionKeyLength:=56, _ 
 PasswordEncryptionFileProperties:=True 
 End If 
 End With 
 
End Sub
```