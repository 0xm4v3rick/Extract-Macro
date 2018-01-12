## Extract-Macro

This PS script will extract macro from Excel and Word files. Also checks the macro for suspecious code patterns  
Includes temporary DDE check for word documents.  

**Version**  
0.3

**Dependencies**  
MS Office 2013 or greater  
Administrator privileges  

**Tested on**  
MS Office 2013  
MS Office 2016  

**Supported file types**  
xls,xlsm,doc,docm,docx (haven't checked for others, may work)  

**Usage**  
PS C:\> ./Extract-macro.ps1 C:\Sheet1.xls  

**TODO**  
- [x] Add support for doc files  
- [ ] Adding more malicious/suspecious macro checks  
- [x] Improve Error Handling  
- [ ] Decoding and checking base64 encoded strings for patterns  
- [ ] Improving DDE check feature for word  

**Sample Run**   

	PS C:\> ./Extract-macro.ps1 C:\Sheet1.xls
	======== Macro Code Start ============
        Sub Auto_open()
            Dim encode As String
            Dim pathName As String
            Dim o As Document
            Set o = ActiveDocument
            
            Dim strResult As String
            Dim objHTTP As Object
            Dim URL As String
            Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
            URL = "http://127.0.0.1:8000/test.txt"
            objHTTP.Open "GET", URL, False
            objHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
            objHTTP.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
            objHTTP.send ("keyword=php")
            strResult = objHTTP.responseText
            MsgBox (strResult)
            
        End Sub





        ======== Macro Code End ============
        ======== Suspecious Macro Code Patterns ============

        Checks_for                                    Count
        ----------                                    -----
        Use of Char encoding                          0    
        Use of shell function                         0    
        base64 encoded strings [Confirmed]            0    
        scheduled tasks invocation. Possible backdoor 0    
        Auto run macro Auto_Open                      1    
        base64 encoded strings [false positive prone] 46   
        Auto run macro Document_Open                  0    



        PS C:\> ./Extract-macro.ps1 C:\dde.docx        
        ======== DDE Code Start ============
        DDEAUTO c:\\windows\\system32\\cmd.exe "/k calc.exe" !Unexpected End of Formula
        ======== DDE Code End ============   

**References**  
https://github.com/enigma0x3/Generate-Macro/blob/master/Generate-Macro.ps1

