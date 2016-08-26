Public gBasePath           ' Stores the value of the current QTP test base folder.
gBasePath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(Environment.Value("TestDir"))
 
 
'Load the properties as Environment Variables of QTP/UFT
LoadProperties(gBasePath & "\properties\" & TestArgs("env") & ".properties")
 
'Below functions find all the properties files and create Environment variables
Sub LoadProperties(ByVal FilePath)
 
    Set ADODB = CreateObject("ADODB.Stream")
    On Error Resume Next
     
    ADODB.CharSet = "utf-8"
    ADODB.Open
    ADODB.LoadFromFile(FilePath)       
     
    arrData = Split(ADODB.ReadText(), vbNewLine)
    For iLoop = 0 To UBound(arrData) Step 1
        txt = arrData(iLoop)
        'Condition to read the property is
        'It should not start with #
        'Min length should be 2
        'Position of = should not be 1
        If Left(txt, 1) <> "#" AND Len(txt) > 2 AND Instr(1, txt, "=") > 1 Then
            intPos = Instr(1, txt, "=")
            strProp = Left(txt, intPos - 1)
            If Len(txt) > intPos Then
                strValue = Mid(txt, intPos + 1, Len(txt))
            Else
                strValue = ""
            End If
            Environment.Value(Trim(strProp)) = Trim(strValue)
        End If
    Next
    ADODB.Close
     
    On Error GoTo 0
    Set ADODB = Nothing
 
End Sub