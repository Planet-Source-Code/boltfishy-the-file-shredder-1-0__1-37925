VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options :: OverWriting"
   ClientHeight    =   1050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3105
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
   ScaleHeight     =   1050
   ScaleWidth      =   3105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   1440
      TabIndex        =   3
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   315
      Left            =   2280
      TabIndex        =   2
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox txtNumberOfTimes 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Text            =   "1000"
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "How many times should The File Shredder overwrite these / this file(s)?"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Please refer to HELP (modHelp.bas) for documentation.
'-----------------------------------------------------

Private Sub Command1_Click() 'OK

    On Error GoTo ErrSub
    
    If Val(txtNumberOfTimes.Text) = 0 Then
    'if over writing number is 0
    
    MsgBox ("Cannot overwrite 0 times." & vbCrLf & "Please use another number."), vbCritical + vbOKOnly, "Error"
    'show an error box as we need a positive integer
    
    Else
    'if it is not 0, i.e. it's an OK number
    
    NumberOfTimes = Val(txtNumberOfTimes.Text)
    'set variable to text box
    
    Unload Me
    'close form
    
    End If
    
ErrSub:
    
    '6 = Overflow, number too big
    '0 = Always generates
    
    If Err.Number = 6 Then: MsgBox ("Number is too large!"), vbCritical + vbOKOnly, "Error"
    If Err.Number = 0 Then: Resume Next
    
End Sub

    Private Sub Command2_Click() 'cancel
    Unload Me
    End Sub

    Private Sub Form_Load()
    txtNumberOfTimes.Text = NumberOfTimes
    End Sub
