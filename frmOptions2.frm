VERSION 5.00
Begin VB.Form frmOptions2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options :: Corruption"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
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
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      Caption         =   "Generate New Key"
      Height          =   315
      Left            =   1200
      TabIndex        =   6
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   315
      Left            =   3840
      TabIndex        =   4
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   3000
      TabIndex        =   5
      Top             =   1800
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Corruption Options"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.CheckBox chkMischa 
         Caption         =   "Protect Files using TFS Hex corruption?"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Value           =   1  'Checked
         Width           =   4095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Use a random key"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   1200
         Value           =   2  'Grayed
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "what is my random key?"
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
         Height          =   255
         Left            =   1920
         TabIndex        =   3
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Before deleting files, it is a good idea to corrupt them - this makes the recovery job a lot harder."
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   4095
      End
   End
End
Attribute VB_Name = "frmOptions2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Please refer to HELP (modHelp.bas) for documentation.
'-----------------------------------------------------

Private Sub Command1_Click() 'OK
    
    If chkMischa.Value = 1 Then 'if TFS Hex is checked
    HexCorrupt = True
    
    ElseIf chkMischa.Value = 0 Then 'if TSF Hex is unchecked
    HexCorrupt = False
    
    End If
    
    Unload Me 'close form
    
End Sub

    Private Sub Command2_Click() 'cancel
    Unload Me
    End Sub

Private Sub Command3_Click() 'generate new key
    
    Dim i As Integer, newd As Variant
    For i = 1 To 96
    Randomize
    
    newd = newd & Chr(Rnd * 255)
    'make it a random character
    
    frmMain.txtKey.Text = newd
    frmMain.txtKey.Text = StrToHex(frmMain.txtKey.Text) 'convert to hexadecimal
    
    Next i
    
End Sub

Private Sub Form_Load()

    'Takes the value of a global value in order to give the
    'option checkbox the correct value when the form
    'is displayed

    If HexCorrupt = True Then
    chkMischa.Value = 1
    
    ElseIf HexCorrupt = False Then
    chkMischa.Value = 0
    
    End If

End Sub

    Private Sub Label2_Click() 'what is my key?
    MsgBox (frmMain.txtKey.Text & vbCrLf & vbCrLf & "A fresh key is generated each session and is used to encrypt your file(s) before" & vbCrLf & "they are deleted. Because they are random, it adds to the encryption security."), vbInformation + vbOKOnly, "Encryption Key"
    End Sub
