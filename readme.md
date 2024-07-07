## Extract-Macro

This PS script will extract macro from Excel and Word files. Also checks the macro for suspecious code patterns  
Includes temporary DDE check for word documents.  

> [!CAUTION]
> I have not tested this but use this script at your own risk.    
> Refer issue https://github.com/0xm4v3rick/Extract-Macro/issues/1 for more details 

**Version**  
0.4

**Dependencies**  
MS Office 2013 or greater  


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
- [x] Decoding and checking base64 encoded
- [ ] Checking base64 encoded strings for patterns  
- [ ] Improving DDE check feature for word  

**Sample Run 1**   

	PS C:\> ./Extract-macro.ps1 C:\Sheet1.xls -fp 0
        ======== Macro Code Start ============
        Sub Auto_open()
            Dim encode As String
            Dim pathName As String
            Dim o As Document
            Set o = ActiveDocument

            Dim strResult As String
            Dim test As String
            Dim objHTTP As Object
            Dim URL As String
            Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
            test = "aHR0cDovLzEyNy4wLjAuMS90ZXN0LnR4dA=="
            URL = "http://127.0.0.1:8000/test.txt"
            objHTTP.Open "GET", URL, False
            objHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
            objHTTP.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
            objHTTP.send ("keyword=php")
            strResult = objHTTP.responseText
            MsgBox (strResult)
        End Sub





        ======== Macro Code End ============
        ========  base64 data found ============

        EncodedText                          DecodedText              
        -----------                          -----------              
        aHR0cDovLzEyNy4wLjAuMS90ZXN0LnR4dA== http://127.0.0.1/test.txt


        ======== Suspecious Macro Code Patterns ============

        Checks_for                                    Count
        ----------                                    -----
        Base64 encoded strings [Confirmed]            1    
        Use of Char encoding                          0    
        string concatination for AV evasion           0    
        Auto run macro Auto_Open                      1    
        IP Address - Possible Data transfer           1    
        HTTP Request modules used                     2    
        base64 encoded strings [false positive prone] 50   
        scheduled tasks invocation. Possible backdoor 0    
        URL detected - Probable data transfer         0    
        Use of shell function                         0    
        Auto run macro Document_Open                  0    
        HTTP Request modules used                     2    



**Sample Run 2**   

	PS C:\> ./Extract-macro.ps1 C:\dde.docx        
	======== DDE Code Start ============
	DDEAUTO c:\\windows\\system32\\cmd.exe "/k calc.exe" !Unexpected End of Formula
	======== DDE Code End ============   

**References**  
https://github.com/enigma0x3/Generate-Macro/blob/master/Generate-Macro.ps1

