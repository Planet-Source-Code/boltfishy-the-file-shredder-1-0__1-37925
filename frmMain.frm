VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "The File Shredder v1.0 - 0 files"
   ClientHeight    =   4455
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   6570
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   6570
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtKey 
      Height          =   375
      Left            =   120
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   4560
      Width           =   6375
   End
   Begin MSComctlLib.StatusBar SB1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   4080
      Width           =   6570
      _ExtentX        =   11589
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1059
            MinWidth        =   1059
            TextSave        =   "21:03"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2648
            MinWidth        =   2648
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3529
            MinWidth        =   3529
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame3 
      Caption         =   "Quick Links"
      Height          =   855
      Left            =   120
      TabIndex        =   8
      Top             =   3120
      Width           =   6330
      Begin VB.CommandButton cmdOptions 
         Caption         =   "&Options"
         Height          =   375
         Left            =   4080
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdAbout 
         Caption         =   "&About"
         Height          =   375
         Left            =   5280
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdClearItem 
         Caption         =   "Clear &Item"
         Height          =   375
         Left            =   720
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "&Clear List"
         Height          =   375
         Left            =   1800
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdDeleteAll 
         Caption         =   "&Delete All"
         Height          =   375
         Left            =   2880
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "..."
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "File Path - Drag && Drop is enabled"
      Height          =   2055
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   6330
      Begin VB.ListBox lstFiles 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1470
         Left            =   120
         OLEDropMode     =   1  'Manual
         Style           =   1  'Checkbox
         TabIndex        =   0
         Top             =   360
         Width           =   5970
      End
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   0
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image imgLogo 
      Height          =   720
      Left            =   120
      Picture         =   "frmMain.frx":0000
      Top             =   120
      Width           =   6300
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuSelectFile 
         Caption         =   "&Select File"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuS1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClearItem 
         Caption         =   "Clear &Item"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuClearList 
         Caption         =   "&Clear List"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuS2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDeleteAll 
         Caption         =   "&Delete All"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuS3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^{F4}
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuOverWrite 
         Caption         =   "Overwriting..."
      End
      Begin VB.Menu mnuEncryption 
         Caption         =   "Corruption"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Please refer to HELP (modHelp.bas) for documentation.
'-----------------------------------------------------


Private Sub cmdBrowse_Click()

    Dim file1 As String

    CD1.ShowOpen
    file1 = FreeFile

    If CD1.FileName <> "" Then 'if file name is true
    file1 = CD1.FileName 'return file path
    
    ElseIf CD1.FileName = "" Then 'if file name is false
    file1 = ""
    
    Exit Sub 'then quit
    End If
          
    If FileExists(file1) = True Then 'if file exists
    lstFiles.AddItem (file1) 'add it to list of files
    
    Me.Caption = "The File Shredder v2.0 - " & lstFiles.ListCount & " files"
    'change caption to include new files number

    ElseIf FileExists(file1) = False Then
    'if file doesn't exist
    
    SB1.Panels(4).Text = Time & ": Error - file does not exist!"
    
    file1 = ""
    CD1.FileName = ""
    
    Exit Sub
    End If

    file1 = ""
    CD1.FileName = ""
    'replace this so that next time we get a clean space
    
End Sub

Private Sub cmdClear_Click() 'clear list

    With lstFiles
    .Clear
    .Refresh
    End With
    
    Me.Caption = "The File Shredder v2.0 - " & lstFiles.ListCount & " files"
    'reset caption to include new files number

End Sub

Private Sub cmdClearItem_Click()

    lstFiles.RemoveItem lstFiles.ListIndex
    'remove the select item from list
    
    Me.Caption = "The File Shredder v2.0 - " & lstFiles.ListCount & " files"
    'reset caption to include new files number
    
End Sub

Private Sub cmdDeleteAll_Click()

    On Error GoTo ErrSub
    
    Dim i As Integer 'counter to go from 1 to no of files
    Dim b As Integer 'no of files
    
    Dim File2Del As String 'file to delete
    Dim msg As String 'message box

    msg = "WARNING: Files cannot be recovered once deleted!"
    msg = msg & vbCrLf & "Are you sure?" 'check if sure

    If MsgBox(msg, vbExclamation + vbYesNo, "Sure?") = vbNo Then
    Exit Sub 'if answer is no then exit
    
    Else 'if answer is yes then go ahead
    b = lstFiles.ListCount '= number of files

    For i = 0 To b - 1 Step 1 'i = 1 to number of files

    frmMain.Enabled = False 'don't let use do anything
    SB1.Panels(2).Text = "Deleting... " & i & " of " & b
    SB1.Panels(3).Text = GetFileName(lstFiles.List(i))

    FileTemp = lstFiles.List(i)
    'set global filetemp to the file to be deleted

    ShredFile (lstFiles.List(i))
    'kill item - file - on list, get the file from i

    Next i

    If i = b Then SB1.Panels(2) = "Deleted " & b & " files!"
    'when finished i.e. when i has reached b (the no. of files)
    
    frmMain.Enabled = True 're-enable the main form
    lstFiles.Clear 'and clear the list

ErrSub:

    Exit Sub
    End If
    
    Me.Caption = "The File Shredder v2.0 - " & lstFiles.ListCount & " files"
    'reset caption to include new files number

End Sub

Private Sub Form_Load()
    
    NumberOfTimes = 1000 'default value of overwriting
    EncMethod = "BlowFish" 'default encryption method
    HexCorrupt = True 'do a hex corruption
    
    Dim i As Integer, newd As Variant
    For i = 1 To 96
    Randomize
    
    newd = newd & Chr(Rnd * 255)
    'make it a random character
    
    txtKey.Text = newd
    txtKey.Text = StrToHex(txtKey.Text) 'convert to hexadecimal
    
    Next i
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload frmAbout: Unload frmOptions: Unload frmOptions2: Unload Me: End
End Sub


Private Sub lstFiles_click() 'file selected

    Dim i As Integer
    For i = 0 To lstFiles.ListCount - 1
    
    If lstFiles.Selected(i) = True Then
    SB1.Panels(3).Text = GetFileName(lstFiles.List(i))
    'set panel 3 to file name of selected file
    
    End If
    Next i
    
End Sub

Private Sub lstFiles_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
    'files are dragged and dropped onto the list box

    On Error Resume Next

    Dim intFiles As Integer
    Dim intLenFile As Integer
    
    Dim intX As Integer
    Dim strFilePath As String

    DoEvents
    intFiles = Data.Files.Count
    
    For intX = 1 To intFiles
    '1 to no of files dropped,
    'i.e. keep adding until all the files have been added
                             
    If (GetAttr(Data.Files(intX)) And vbDirectory) = vbDirectory Then
    'check for directory
    Exit Sub 'if a dir was dropped then stop
    
    Else 'but if a file(s) was dropped then
    intLenFile = Len(Data.Files(intX))
    
    strFilePath = Left(Data.Files(intX), intLenFile)
    'return file path
    
    lstFiles.AddItem "" & strFilePath
    'add the file path to the listbox
    
    End If
    Next intX

    Me.Caption = "The File Shredder v2.0 - " & lstFiles.ListCount & " files"
    'reset the caption to include new files number

End Sub

'menu and simple links section
'menu bit basically just calls button click events

    Private Sub cmdAbout_Click()
    frmAbout.Show
    End Sub

    Private Sub cmdOptions_Click()
    PopupMenu mnuOptions 'popup the menu command
    End Sub

    Private Sub mnuAbout_Click()
    frmAbout.Show
    End Sub

    Private Sub mnuClearItem_Click()
    Call cmdClearItem_Click
    End Sub

    Private Sub mnuClearList_Click()
    Call cmdClear_Click
    End Sub

    Private Sub mnuDeleteAll_Click()
    Call cmdDeleteAll_Click
    End Sub

    Private Sub mnuEncryption_Click()
    frmOptions2.Show
    End Sub

    Private Sub mnuExit_Click()
    Unload frmAbout: Unload frmOptions: Unload frmOptions2: Unload Me: End
    End Sub

    Private Sub mnuOverWrite_Click()
    frmOptions.Show
    End Sub

    Private Sub mnuSelectFile_Click()
    Call cmdBrowse_Click
    End Sub
