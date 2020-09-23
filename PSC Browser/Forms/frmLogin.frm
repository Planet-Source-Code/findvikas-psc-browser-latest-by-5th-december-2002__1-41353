VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLogin 
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login To PSC "
   ClientHeight    =   4065
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4650
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   271
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   310
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar status 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   12
      Top             =   3765
      Width           =   4650
      _ExtentX        =   8202
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   8149
            MinWidth        =   8149
            Text            =   "Status: OK"
            TextSave        =   "Status: OK"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000A&
      Caption         =   "New User Sign Up"
      Height          =   1095
      Left            =   210
      TabIndex        =   9
      Top             =   150
      Width           =   4185
      Begin VB.CommandButton btnNewUser 
         Caption         =   "Create &New User Account"
         Height          =   405
         Left            =   390
         TabIndex        =   0
         Top             =   390
         Width           =   3405
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Existing User Log On"
      Height          =   2265
      Left            =   210
      TabIndex        =   6
      Top             =   1410
      Width           =   4185
      Begin VB.CommandButton btnLogOn 
         Caption         =   "&Log On"
         Height          =   345
         Left            =   2205
         TabIndex        =   4
         Top             =   1755
         Width           =   900
      End
      Begin VB.CheckBox chkRememberPassword 
         Caption         =   "&Save Password"
         Height          =   345
         Left            =   690
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1755
         Width           =   1455
      End
      Begin VB.CommandButton btnCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   345
         Left            =   3135
         TabIndex        =   5
         Top             =   1755
         Width           =   900
      End
      Begin VB.TextBox txtPassword 
         BackColor       =   &H00D8E9EC&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF6400&
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   1305
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   720
         Width           =   2745
      End
      Begin VB.TextBox txtEmailID 
         BackColor       =   &H00D8E9EC&
         ForeColor       =   &H00FF6400&
         Height          =   345
         Left            =   1305
         TabIndex        =   1
         Top             =   330
         Width           =   2745
      End
      Begin VB.Label lblDirectLogin 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Direct Login [ The Hacker's Way ]"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   750
         MouseIcon       =   "frmLogin.frx":058A
         MousePointer    =   99  'Custom
         TabIndex        =   11
         Top             =   1455
         Width           =   2535
      End
      Begin VB.Label lblForgot 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Help Me! I can't remember my password! "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   750
         MouseIcon       =   "frmLogin.frx":06DC
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   1200
         Width           =   3165
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   8
         Top             =   810
         Width           =   735
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E-Mail ID:"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   7
         Top             =   420
         Width           =   690
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnCancel_Click()

On Error GoTo Debugger
    
    frmCommon.Inet.Cancel
    status.Panels(1).Text = "Status: OK"
    btnLogOn.Enabled = True
    btnLogOn.Caption = "&Log On"
    
    
    Me.Hide
    frmMain.Show
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmLogin.btnCancel_Click", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub

Private Sub btnLogOn_Click()

On Error GoTo Debugger
Dim Temp As String
        
    btnLogOn.Enabled = False
    btnLogOn.Caption = "Wait..."
    If txtEmailID = "" Then MsgBox "Please Enter Email ID", vbExclamation: txtEmailID.SetFocus: Exit Sub
    If InStr(txtEmailID, "@") = 0 Or InStr(txtEmailID, "@") = Len(txtEmailID) Then MsgBox "Please Enter A Valid Email ID", vbExclamation: txtEmailID.SetFocus: Exit Sub
    If txtPassword = "" Then MsgBox "Please Enter Password", vbExclamation: txtPassword.SetFocus: Exit Sub
    
    If SavePassword = "TRUE" Then
        SaveSetting App.ProductName, dynamicProfileSection, "EmailID", txtEmailID.Text
        SaveSetting App.ProductName, dynamicProfileSection, "Password", Enc(txtPassword.Text)
        SaveSetting App.ProductName, dynamicSettingSection, "SavePassword", "TRUE"
    End If

    
    Temp = "txtEmailAddress=" & EmailId & "&txtReturnURL=" & dynamicLogOnReturnURL & "&lngWId=&blnOutsideOfVBSubWeb=FALSE&txtPassword=" & Password & "&chkRememberPassword=" & SavePassword & "&cmOk=Ok&strPassKey="
    Temp = InetURL(dynamicLogOnURL, "POST ", Temp, "User-Agent: PSC Browser v " & AppVersion & vbNewLine & "Request Time: " & Now)
    
    If Not Temp = "" Then
        MsgBox "Success With Following Returns:" & vbNewLine & Temp, vbInformation
        btnLogOn.Enabled = True
        btnLogOn.Caption = "&Log On"
        Me.Hide
        frmMain.Show
    Else
        MsgBox "Logon Failure.", vbInformation
        btnLogOn.Enabled = True
        btnLogOn.Caption = "&Log On"
    End If
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmLogin.btnLogOn_Click", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub

Public Sub btnNewUser_Click()

On Error GoTo Debugger
    
        JumpURL dynamicNewUserURL
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmLogin.btnNewUser_Click", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub

Private Sub chkRememberPassword_Click()

On Error GoTo Debugger
    
If chkRememberPassword Then
    SavePassword = "TRUE"
Else
    SavePassword = "FALSE"
End If
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmLogin.chkRememberPassword_Click", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
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
    
    LoadUserSettings
    
    
    txtEmailID = GetSetting(App.ProductName, dynamicProfileSection, "EmailID", "")
    txtPassword = Dec(GetSetting(App.ProductName, dynamicProfileSection, "Password", ""))
    
    SavePassword = GetSetting(App.ProductName, dynamicSettingSection, "SavePassword", "FALSE")
    If SavePassword = "TRUE" Then chkRememberPassword.Value = 1
    
    
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmLogin.Form_Load", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub


Private Sub Form_Unload(Cancel As Integer)

On Error GoTo Debugger
    
    If SavePassword = "TRUE" Then
        
        If Not txtEmailID = "" Then SaveSetting App.ProductName, dynamicProfileSection, "EmailID", txtEmailID.Text
        If Not txtPassword = "" Then SaveSetting App.ProductName, dynamicProfileSection, "Password", Enc(txtPassword.Text)
        SaveSetting App.ProductName, dynamicSettingSection, "SavePassword", "TRUE"
    End If
    Me.Hide
    frmMain.Show
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmLogin.Form_Unload", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub


Private Sub lblDirectLogin_Click()

On Error GoTo Debugger
    
    frmDirectLogin.Show

Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmLogin.lblDirectLogin_Click", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub

Private Sub lblDirectLogin_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

On Error GoTo Debugger
    
    lblDirectLogin.ForeColor = vbRed
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmLogin.lblDirectLogin_MouseDown", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    

End Sub

Private Sub lblDirectLogin_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

On Error GoTo Debugger
    
    lblDirectLogin.ForeColor = vbBlue
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmLogin.lblDirectLogin_MouseUp", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
End Sub

Private Sub lblForgot_Click()

On Error GoTo Debugger
    
    JumpURL dynamicForgotPwdURL
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmLogin.lblForgot_Click", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub

Private Sub lblForgot_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

On Error GoTo Debugger
    
    lblForgot.ForeColor = vbRed
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmLogin.lblForgot_MouseDown", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub


Private Sub lblForgot_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

On Error GoTo Debugger
    
    lblForgot.ForeColor = vbBlue
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmLogin.lblForgot_MouseUp", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub

