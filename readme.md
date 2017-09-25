## Extract-macro

This PS script will extract the macro code from the xls file.  

**Dependencies**  
MS Office 2013 or greater  
Administrator privileges  

**Tested on**  
MS Office 2013  
MS Office 2016  

**Usage**  
PS C:\> ./Extract-macro.ps1 C:\Sheet1.xls

**TODO**
- Add support for doc files

Sample Run 

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

