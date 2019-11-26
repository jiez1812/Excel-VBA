Function num2letter(num As Integer) As String
    Dim cmod As Integer
    Dim cnum As Integer
    Dim strBuffer As String: strBuffer = ""
    
    cnum = num
    
    Do While cnum > 0
        cmod = cnum Mod 26
        strBuffer = Chr(cmod + 64) & strBuffer
        cnum = (cnum - cmod) / 26
    Loop
    num2letter = strBuffer
End Function