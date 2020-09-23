VERSION 5.00
Begin VB.Form frmConfig 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuration"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7395
   Icon            =   "frmConfig.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   7395
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtUpLoadURL 
      BackColor       =   &H00D8E9EC&
      ForeColor       =   &H00FF6400&
      Height          =   315
      Left            =   1770
      TabIndex        =   6
      Top             =   2250
      Width           =   5500
   End
   Begin VB.TextBox txtLogOnReturnURL 
      BackColor       =   &H00D8E9EC&
      ForeColor       =   &H00FF6400&
      Height          =   315
      Left            =   1770
      TabIndex        =   5
      Top             =   1895
      Width           =   5500
   End
   Begin VB.TextBox txtLogOnURL 
      BackColor       =   &H00D8E9EC&
      ForeColor       =   &H00FF6400&
      Height          =   315
      Left            =   1770
      TabIndex        =   4
      Top             =   1540
      Width           =   5500
   End
   Begin VB.CommandButton btnDefaults 
      Caption         =   "&Defaults"
      Height          =   345
      Left            =   4290
      TabIndex        =   9
      Top             =   2730
      Width           =   975
   End
   Begin VB.CommandButton btnApply 
      Caption         =   "&Apply"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   345
      Left            =   5310
      TabIndex        =   7
      Top             =   2730
      Width           =   975
   End
   Begin VB.CommandButton btnClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   345
      Left            =   6300
      TabIndex        =   8
      Top             =   2730
      Width           =   975
   End
   Begin VB.TextBox txtCodeOfDayURL 
      BackColor       =   &H00D8E9EC&
      ForeColor       =   &H00FF6400&
      Height          =   315
      Left            =   1770
      TabIndex        =   3
      Top             =   1185
      Width           =   5500
   End
   Begin VB.TextBox txtForgotPwdURL 
      BackColor       =   &H00D8E9EC&
      ForeColor       =   &H00FF6400&
      Height          =   315
      Left            =   1770
      TabIndex        =   2
      Top             =   830
      Width           =   5500
   End
   Begin VB.TextBox txtNewUserURL 
      BackColor       =   &H00D8E9EC&
      ForeColor       =   &H00FF6400&
      Height          =   315
      Left            =   1770
      TabIndex        =   1
      Top             =   475
      Width           =   5500
   End
   Begin VB.TextBox txtSearchURL 
      BackColor       =   &H00D8E9EC&
      ForeColor       =   &H00FF6400&
      Height          =   315
      Left            =   1770
      TabIndex        =   0
      Top             =   120
      Width           =   5500
   End
   Begin VB.Label lblUpLoad 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Upload URL:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   90
      TabIndex        =   16
      Top             =   2220
      Width           =   915
   End
   Begin VB.Label lblLogOnReturn 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Log On Return URL:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   90
      TabIndex        =   15
      Top             =   1890
      Width           =   1455
   End
   Begin VB.Label lblForgotPwd 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Forgot Password URL:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   90
      TabIndex        =   14
      Top             =   900
      Width           =   1560
   End
   Begin VB.Label lblNewUser 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New User URL:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   90
      TabIndex        =   13
      Top             =   570
      Width           =   1080
   End
   Begin VB.Label lblCodeOfDay 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Code Of Day URL:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   90
      TabIndex        =   12
      Top             =   1230
      Width           =   1320
   End
   Begin VB.Label lblSearch 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search URL:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   90
      TabIndex        =   11
      Top             =   240
      Width           =   870
   End
   Begin VB.Label lblLogOn 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Log On URL:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   90
      TabIndex        =   10
      Top             =   1560
      Width           =   945
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub btnApply_Click()

On Error GoTo Debugger
    
    Call SaveSetting(App.ProductName, dynamicURLSection, "SearchURL", txtSearchURL.Text)
    Call SaveSetting(App.ProductName, dynamicURLSection, "NewUserURL", txtNewUserURL.Text)
    Call SaveSetting(App.ProductName, dynamicURLSection, "ForgotPwdURL", txtForgotPwdURL.Text)
    Call SaveSetting(App.ProductName, dynamicURLSection, "CodeOfDayURL", txtCodeOfDayURL.Text)
    Call SaveSetting(App.ProductName, dynamicURLSection, "LogOnURL", txtLogOnURL.Text)
    Call SaveSetting(App.ProductName, dynamicURLSection, "LogOnReturnURL", txtLogOnReturnURL.Text)
    Call SaveSetting(App.ProductName, dynamicURLSection, "UpLoadURL", txtUpLoadURL.Text)
    
    dynamicSearchURL = txtSearchURL.Text
    dynamicNewUserURL = txtNewUserURL.Text
    dynamicForgotPwdURL = txtForgotPwdURL.Text
    dynamicCodeOfDayURL = txtCodeOfDayURL.Text
    dynamicLogOnURL = txtLogOnURL.Text
    dynamicLogOnReturnURL = txtLogOnReturnURL.Text
    dynamicUpLoadURL = txtUpLoadURL.Text
    
    btnApply.Enabled = False
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmConfig.btnApply_Click", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    

End Sub

Private Sub btnClose_Click()

On Error GoTo Debugger
    
    Me.Hide
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmConfig.btnClose_Click", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub

Private Sub btnDefaults_Click()

On Error GoTo Debugger
    
    Call SaveSetting(App.ProductName, dynamicURLSection, "SearchURL", constSearchURL)
    Call SaveSetting(App.ProductName, dynamicURLSection, "NewUserURL", constNewUserURL)
    Call SaveSetting(App.ProductName, dynamicURLSection, "ForgotPwdURL", constForgotPwdURL)
    Call SaveSetting(App.ProductName, dynamicURLSection, "CodeOfDayURL", constCodeOfDayURL)
    Call SaveSetting(App.ProductName, dynamicURLSection, "LogOnURL", constLogOnURL)
    Call SaveSetting(App.ProductName, dynamicURLSection, "LogOnReturnURL", constLogOnReturnURL)
    Call SaveSetting(App.ProductName, dynamicURLSection, "UpLoadURL", constUpLoadURL)
    
    txtSearchURL = constSearchURL
    txtNewUserURL = constNewUserURL
    txtForgotPwdURL = constForgotPwdURL
    txtCodeOfDayURL = constCodeOfDayURL
    txtLogOnURL = constLogOnURL
    txtLogOnReturnURL = constLogOnReturnURL
    txtUpLoadURL = constUpLoadURL
    
    btnApply.Enabled = False
    
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmConfig.btnDefaults_Click", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
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
    
    
    
    txtSearchURL = GetSetting(App.ProductName, dynamicURLSection, "SearchURL", constSearchURL)
    txtNewUserURL = GetSetting(App.ProductName, dynamicURLSection, "NewUserURL", constNewUserURL)
    txtForgotPwdURL = GetSetting(App.ProductName, dynamicURLSection, "ForgotPwdURL", constForgotPwdURL)
    txtCodeOfDayURL = GetSetting(App.ProductName, dynamicURLSection, "CodeOfDayURL", constCodeOfDayURL)
    txtLogOnURL = GetSetting(App.ProductName, dynamicURLSection, "LogOnURL", constLogOnURL)
    txtLogOnReturnURL = GetSetting(App.ProductName, dynamicURLSection, "LogOnReturnURL", constLogOnReturnURL)
    txtUpLoadURL = GetSetting(App.ProductName, dynamicURLSection, "UpLoadURL", constUpLoadURL)
    
    Show
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmConfig.Form_Load", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub

Private Sub txtCodeOfDayURL_Change()

On Error GoTo Debugger
    
    btnApply.Enabled = True
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmConfig.txtCodeOfDayURL_Change", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    

End Sub

Private Sub txtForgotPwdURL_Change()

On Error GoTo Debugger
    
    btnApply.Enabled = True
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmConfig.txtForgotPwdURL_Change", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    

End Sub

Private Sub txtLogOnReturnURL_Change()

On Error GoTo Debugger
    
    btnApply.Enabled = True

    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmConfig.txtLogOnReturnURL_Change", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub

Private Sub txtLogOnURL_Change()

On Error GoTo Debugger
    
    btnApply.Enabled = True
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmConfig.txtLogOnURL_Change", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    

End Sub

Private Sub txtNewUserURL_Change()

On Error GoTo Debugger
    
    btnApply.Enabled = True

    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmConfig.txtNewUserURL_Change", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub

Private Sub txtSearchURL_Change()

On Error GoTo Debugger
    
    btnApply.Enabled = True
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmConfig.txtSearchURL_Change", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub
