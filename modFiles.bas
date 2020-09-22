Attribute VB_Name = "modFiles"
'Please refer to HELP (modHelp.bas) for documentation.
'-----------------------------------------------------

Public Function FileExists(FilePath As String) As Boolean

    'find out if a file exists
    FileExists = Dir(FilePath) <> ""
    
End Function

Public Function GetFileName(FilePath As String) As String
    'return file name from a path
    
    Dim i As Integer
    On Error Resume Next

    For i = Len(FilePath) To 1 Step -1 'i to length of file going back
    If Mid(FilePath, i, 1) = "\" Then 'when it finds the \
    
    Exit For 'stop trying
    End If
    
    Next
     
    GetFileName = Mid(FilePath, i + 1)

End Function
