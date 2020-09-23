VERSION 5.00
Begin VB.Form frmCode 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Code Windows"
   ClientHeight    =   5235
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   Icon            =   "frmCode.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   6030
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtCode 
      BackColor       =   &H00D8E9EC&
      ForeColor       =   &H00FF6400&
      Height          =   4665
      Left            =   60
      MaxLength       =   65300
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   60
      Width           =   5895
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4740
      TabIndex        =   1
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   3510
      TabIndex        =   0
      Top             =   4800
      Width           =   1215
   End
End
Attribute VB_Name = "frmCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CancelButton_Click()

On Error GoTo Debugger

    Me.Hide
    
Debugger:
    If Not Err.Number = 0 And Not Err.Number = -2147467259 Then
        Select Case ErrLog(Err.Number, Error$, "frmCode.CancelButton_Click", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If

End Sub

Private Sub Form_Load()

On Error GoTo Debugger

    txtCode.Text = sTextCode
    
Debugger:
    If Not Err.Number = 0 And Not Err.Number = -2147467259 Then
        Select Case ErrLog(Err.Number, Error$, "frmCode.Form_Load", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If

    
End Sub


Private Sub OKButton_Click()

On Error GoTo Debugger

    sTextCode = txtCode.Text
    Me.Hide
    
Debugger:
    If Not Err.Number = 0 And Not Err.Number = -2147467259 Then
        Select Case ErrLog(Err.Number, Error$, "frmCode.OKButton_Click", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If

End Sub
