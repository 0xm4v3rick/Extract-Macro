## Extract-Macro

This PS script will extract macro from Excel and Word files.Also checks the macro for suspecious code patterns  

**Version**  
0.3

**Dependencies**  
MS Office 2013 or greater  
Administrator privileges  

**Tested on**  
MS Office 2013  
MS Office 2016  

**Supported file types**  
xls,xlsm,doc,docm (haven't checked for others, may work)  

**Usage**  
PS C:\> ./Extract-macro.ps1 C:\Sheet1.xls  

**TODO**  
- [x] Add support for doc files  
- [ ] Adding more malicious/suspecious macro checks  
- [x] Improve Error Handling  
- [ ] Decoding and checking base64 encoded strings for patterns  

**Sample Run**   

	PS C:\> ./Extract-macro.ps1 C:\Sheet1.xls
	======== Macro Code Start ============
	Sub Auto_open()
	intUserOption = MsgBox("Press Yes or No Button", vbYesNo)
	If vbOption = 6 Then
	MsgBox "You Pressed YES Option"
	ElseIf vbOption = 7 Then
	MsgBox "You Pressed NO Option"
	Else
	MsgBox "Nothing!"
	End If
	End Sub
	======== Macro Code End ============

**References**  
https://github.com/enigma0x3/Generate-Macro/blob/master/Generate-Macro.ps1

