Attribute VB_Name = "modHexCorrupt"
'Please refer to HELP (modHelp.bas) for documentation.
'-----------------------------------------------------

Public Function DoHexCorrupt(mFile As String, mKey As String)

    Dim mEncrypt As String * 75
    Dim a, b, c, d As Double
    
    a = Rnd(133 * 121) / Rnd(19 * 157)
    b = Rnd(19 * 13) - Rnd(11 * 7)
    
    c = Rnd(147 * 16) / Rnd(5 * 79)
    d = Rnd(13 * 19) + Rnd(a * b * c)
  
    Dim mTemp As Double
    Open mFile For Random As #1 Len = 75

    For i = 1 To Len(mKey)
    
    mTemp = Asc(Mid$(mKey, i, 1))
    mEncrypt = Str$(((((mTemp * a) / c) * d * b) * i / 3.141592654 * Sqr(a)))
    mEncrypt = StrToHex(mEncrypt)
    
    Put #1, i, mEncrypt
    Put #1, i, mEncrypt
        
    Next i
    Close #1
    
End Function
