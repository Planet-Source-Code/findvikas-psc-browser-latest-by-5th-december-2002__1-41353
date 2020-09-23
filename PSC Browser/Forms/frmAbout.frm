VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MyApp"
   ClientHeight    =   4515
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   7695
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7.964
   ScaleMode       =   0  'User
   ScaleWidth      =   13.574
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtAbout 
      Appearance      =   0  'Flat
      ForeColor       =   &H00FF6400&
      Height          =   2025
      Left            =   525
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Text            =   "frmAbout.frx":058A
      Top             =   1200
      Width           =   6645
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   315
      Left            =   6180
      TabIndex        =   0
      Top             =   3630
      Width           =   960
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":0990
      ForeColor       =   &H00FF6400&
      Height          =   675
      Left            =   2205
      TabIndex        =   5
      Top             =   3360
      Width           =   3285
   End
   Begin VB.Label lblBuild 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Build On 05 Dec. 2002"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H005827DC&
      Height          =   240
      Left            =   2587
      TabIndex        =   3
      Top             =   810
      Width           =   2520
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PSC Browser"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H005827DC&
      Height          =   405
      Left            =   2647
      TabIndex        =   1
      Top             =   120
      Width           =   2310
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version #.#.#"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H005827DC&
      Height          =   240
      Left            =   3082
      TabIndex        =   2
      Top             =   510
      Width           =   1560
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdOK_Click()
On Error GoTo Debugger
    
        Me.Hide
        
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmAbout.cmdOK_Click", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If

End Sub



Private Sub Form_Activate()

On Error GoTo Debugger
    
    Me.Caption = "About " & App.Title
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblBuild = "Build On 03 Dec. 2002"
    
    Show
    
    Me.Move Me.Left - 100, Me.Top + 100: Delay 50: Beep
    Me.Move Me.Left + 100, Me.Top - 100: Delay 50: Beep
    Me.Move Me.Left - 100, Me.Top + 100: Delay 50: Beep
    Me.Move Me.Left + 100, Me.Top - 100: Delay 50: Beep
    Me.Move Me.Left - 100, Me.Top + 100: Delay 50: Beep
    Me.Move Me.Left + 100, Me.Top - 100: Delay 50: Beep
    Me.Move Me.Left - 100, Me.Top + 100: Delay 50: Beep
    Me.Move Me.Left + 100, Me.Top - 100: Delay 50: Beep
        
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmAbout.Form_Load", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
    
End Sub

Private Sub Text1_Change()

End Sub

