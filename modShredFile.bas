Attribute VB_Name = "modShredFile"
'Please refer to HELP (modHelp.bas) for documentation.
'-----------------------------------------------------

Global NumberOfTimes As Long 'number of times we should overwrite, configurable by the user
Global HexCorrupt As Boolean 'should we use hexcorrupt?

Global FileTemp As String
Declare Function FlushFileBuffers Lib "kernel32" (ByVal hfile As Long) As Long


Public Sub ShredFile(sFileName As String)

    On Error GoTo ErrSub
    
    Randomize
    Dim OverData As Integer
    Dim OverChar As String
    Dim i As Long
    
    NumberOfTimes = Val(frmOptions.txtNumberOfTimes.Text)
    'number of times to overwrite = the value of that text box on the options screen
    
    OverData = Rnd * 255
    
    If OverData >= 255 Then
    OverData = OverData - 255
    End If
    
    OverChar = Chr(OverData) 'overwrite character is set to a character with a value of rnd (10*10)
    
    Open sFileName For Binary As #1
    
    For i = 1 To NumberOfTimes
    Put #1, i, OverChar
    
    FlushFileBuffers (1)
    'we must flush the file buffers. if windoze sees that
    'we are going to delete the file anyway it won't
    'bother to overwrite it etc, so we use this API call
    'in order to clear its "memory"
    
    Next i
    
    Close #1
    
    If HexCorrupt = True Then 'if user wants to hex corrupt
    DoHexCorrupt sFileName, frmMain.txtKey 'then do so
    
    ElseIf HexCorrupt = False Then 'if they don't
    End If 'then don't hex corrupt

    Open sFileName For Output As #1 'open the file
                                    'to write to
    Print #1, "" 'replace everything with...er, nothing
    Close #1

    Kill sFileName 'delete it

ErrSub: 'should an error occur

    If Err.Number <> 0 Then
    
    MsgBox (Err.Number & vbCrLf & Err.Description), vbCritical + vbOKOnly, "Error"
    frmMain.SB1.Panels(4).Text = GetFileName(FileTemp) & " - Error: " & Err.Number
    
    Close #1 'if file is already open, close it
    End If

End Sub
