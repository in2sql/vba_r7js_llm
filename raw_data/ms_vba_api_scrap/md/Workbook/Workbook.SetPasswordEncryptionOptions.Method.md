# Workbook SetPasswordEncryptionOptions Method

## Business Description
Sets the options for encrypting workbooks using passwords.

## Behavior
Sets the options for encrypting workbooks using passwords.

## Example Usage
```vba
Sub SetPasswordOptions() 
 
 ActiveWorkbook.SetPasswordEncryptionOptions_ 
 PasswordEncryptionProvider:="Microsoft RSA SChannel Cryptographic Provider", _ 
 PasswordEncryptionAlgorithm:="RC4", _ 
 PasswordEncryptionKeyLength:=56, _ 
 PasswordEncryptionFileProperties:=True 
 
End Sub
```