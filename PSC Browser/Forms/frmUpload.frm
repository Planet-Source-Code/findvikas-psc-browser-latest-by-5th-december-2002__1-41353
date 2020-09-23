VERSION 5.00
Begin VB.Form frmUpload 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Code Upload - [ Currently Not Working Properly ]"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6480
   Icon            =   "frmUpload.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   6480
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnTextCode 
      Caption         =   "Enter Text Code"
      Height          =   345
      Left            =   1260
      TabIndex        =   17
      Top             =   6360
      Width           =   1605
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   5340
      TabIndex        =   19
      Top             =   6360
      Width           =   900
   End
   Begin VB.CommandButton btnUpLoad 
      Caption         =   "&UpLoad"
      Height          =   345
      Left            =   4410
      TabIndex        =   18
      Top             =   6360
      Width           =   900
   End
   Begin VB.CommandButton btnZipFile 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5940
      TabIndex        =   16
      Top             =   5880
      Width           =   300
   End
   Begin VB.CommandButton btnScreenShot 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5940
      TabIndex        =   14
      Top             =   5490
      Width           =   300
   End
   Begin VB.TextBox txtUploadLocalFileName 
      BackColor       =   &H00D8E9EC&
      ForeColor       =   &H00FF6400&
      Height          =   285
      Left            =   1260
      MaxLength       =   100
      TabIndex        =   15
      Top             =   5880
      Width           =   4635
   End
   Begin VB.TextBox txtScreenShot 
      BackColor       =   &H00D8E9EC&
      ForeColor       =   &H00FF6400&
      Height          =   285
      Left            =   1260
      MaxLength       =   100
      TabIndex        =   13
      Top             =   5490
      Width           =   4635
   End
   Begin VB.Frame Experience 
      BackColor       =   &H8000000A&
      Caption         =   "Experience Level"
      Height          =   1095
      Left            =   1260
      TabIndex        =   24
      Top             =   2220
      Width           =   3375
      Begin VB.OptionButton optDifficultyTypeId 
         BackColor       =   &H8000000A&
         Caption         =   "Intermediate (6 months - 1 year)"
         Height          =   195
         Index           =   3
         Left            =   270
         TabIndex        =   26
         Top             =   525
         Width           =   2685
      End
      Begin VB.OptionButton optDifficultyTypeId 
         BackColor       =   &H8000000A&
         Caption         =   "Advanced (1+ years) "
         Height          =   195
         Index           =   4
         Left            =   270
         TabIndex        =   25
         Top             =   780
         Width           =   2685
      End
      Begin VB.OptionButton optDifficultyTypeId 
         BackColor       =   &H8000000A&
         Caption         =   "Beginner (0-6 months experience)"
         Height          =   195
         Index           =   2
         Left            =   270
         TabIndex        =   11
         Top             =   270
         Width           =   2685
      End
   End
   Begin VB.Frame VBPlatform 
      Caption         =   "Compatiblity Platform(s)"
      Height          =   1245
      Left            =   1260
      TabIndex        =   23
      Top             =   900
      Width           =   4935
      Begin VB.CheckBox chkCompat 
         BackColor       =   &H8000000A&
         Caption         =   "VBA MS Access"
         Height          =   285
         Index           =   9
         Left            =   1665
         TabIndex        =   9
         Top             =   840
         Width           =   1515
      End
      Begin VB.CheckBox chkCompat 
         BackColor       =   &H8000000A&
         Caption         =   " VB 5.0 "
         Height          =   285
         Index           =   4
         Left            =   120
         TabIndex        =   5
         Top             =   555
         Width           =   1515
      End
      Begin VB.CheckBox chkCompat 
         BackColor       =   &H8000000A&
         Caption         =   "ASP"
         Height          =   285
         Index           =   8
         Left            =   120
         TabIndex        =   8
         Top             =   870
         Width           =   1515
      End
      Begin VB.CheckBox chkCompat 
         BackColor       =   &H8000000A&
         Caption         =   "VB 3.0"
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1515
      End
      Begin VB.CheckBox chkCompat 
         BackColor       =   &H8000000A&
         Caption         =   "VB Script"
         Height          =   285
         Index           =   7
         Left            =   3210
         TabIndex        =   7
         Top             =   525
         Width           =   1515
      End
      Begin VB.CheckBox chkCompat 
         BackColor       =   &H8000000A&
         Caption         =   "VB 4.0 (16-bit)"
         Height          =   285
         Index           =   2
         Left            =   1665
         TabIndex        =   3
         Top             =   210
         Width           =   1515
      End
      Begin VB.CheckBox chkCompat 
         BackColor       =   &H8000000A&
         Caption         =   "VB 4.0 (32-bit)"
         Height          =   285
         Index           =   3
         Left            =   3210
         TabIndex        =   4
         Top             =   210
         Width           =   1515
      End
      Begin VB.CheckBox chkCompat 
         BackColor       =   &H8000000A&
         Caption         =   " VB 6.0"
         Height          =   285
         Index           =   5
         Left            =   1665
         TabIndex        =   6
         Top             =   525
         Width           =   1515
      End
      Begin VB.CheckBox chkCompat 
         BackColor       =   &H8000000A&
         Caption         =   "VBA MS Excel"
         Height          =   285
         Index           =   10
         Left            =   3210
         TabIndex        =   10
         Top             =   840
         Width           =   1515
      End
   End
   Begin VB.TextBox txtPurpose 
      BackColor       =   &H00D8E9EC&
      ForeColor       =   &H00FF6400&
      Height          =   1995
      Left            =   1260
      MaxLength       =   4000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   3390
      Width           =   4965
   End
   Begin VB.ComboBox cboCategory 
      BackColor       =   &H00D8E9EC&
      ForeColor       =   &H00FF6400&
      Height          =   315
      ItemData        =   "frmUpload.frx":058A
      Left            =   1260
      List            =   "frmUpload.frx":05F5
      TabIndex        =   1
      Text            =   "(select a category)"
      Top             =   450
      Width           =   4905
   End
   Begin VB.TextBox txtCodeName 
      BackColor       =   &H00D8E9EC&
      ForeColor       =   &H00FF6400&
      Height          =   285
      Left            =   1260
      MaxLength       =   100
      TabIndex        =   0
      Top             =   60
      Width           =   4875
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Zipped Code:          Or                                    Text Code"
      Height          =   765
      Index           =   4
      Left            =   120
      TabIndex        =   28
      Top             =   5910
      Width           =   1005
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Screen Shot:"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   27
      Top             =   5520
      Width           =   930
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Purpose:"
      Height          =   195
      Index           =   2
      Left            =   420
      TabIndex        =   22
      Top             =   3420
      Width           =   630
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Category:"
      Height          =   195
      Index           =   1
      Left            =   420
      TabIndex        =   21
      Top             =   540
      Width           =   675
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Code Name:"
      Height          =   195
      Index           =   0
      Left            =   210
      TabIndex        =   20
      Top             =   90
      Width           =   885
   End
End
Attribute VB_Name = "frmUpload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnCancel_Click()

On Error GoTo Debugger
    
    frmCode.Hide
    Me.Hide
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmUpload.btnCancel_Click", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub


Private Sub btnScreenShot_Click()

On Error GoTo Debugger
    
    With frmCommon.cmd
        .DialogTitle = "Select A .jpg or .gif File To Upload On Server"
        .Filter = "Picture File|*.jpg;*.gif|All Files|*.*"
        .ShowOpen
        ScreenShot = .FileName
    End With
    txtScreenShot = ScreenShot
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmUpload.btnScreenShot_Click", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
    
End Sub

Private Sub btnTextCode_Click()

On Error GoTo Debugger
    
    frmCode.Show
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmUpload.btnTextCode_Click", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub

Private Sub btnUpLoad_Click()

On Error GoTo Debugger
    
If MsgBox("This Function Is Not Fully Functional May Be Just Be Cause I Dont Know The Exact URL Or Any Other Reason." & vbNewLine & "Do You Still Want To Continue", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub
    
    MsgBox "If You Know More Details How To Post Code On PSC Then Please Let Me Know", vbInformation
    

    If Len(txtCodeName) <= 0 Or Len(txtCodeName) > 100 Then
        MsgBox "Length Of Code Name Should Be In Between 1 to 100" & vbNewLine & "Your Length is " & Len(txtCodeName), vbInformation
        txtCodeName.SetFocus
        Exit Sub
    Else
        CodeName = txtCodeName
    End If
    
    If Len(txtPurpose) <= 0 Or Len(txtPurpose) > 4000 Then
        MsgBox "Length Of Purpose Should Be In Between 1 to 4000" & vbNewLine & "Your Length is " & Len(txtPurpose), vbInformation
        txtPurpose.SetFocus
        Exit Sub
    Else
        Purpose = txtPurpose
    End If
    
    If Len(frmCode.txtCode) > 65300 Then
        MsgBox "Length Of Code Should Be Less Than 65300" & vbNewLine & "Your Length is " & Len(frmCode.txtCode), vbInformation
        Exit Sub
    Else
        scode = frmCode.txtCode
    End If
    
    Purpose = txtPurpose.Text
    
    ScreenShot = txtScreenShot.Text
    
    UploadLocalFileName = txtUploadLocalFileName.Text
    
    CodeUpload
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmUpload.btnUpLoad_Click", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
    
End Sub

Private Sub btnZipFile_Click()

On Error GoTo Debugger
    
    With frmCommon.cmd
        .DialogTitle = "Select A .Zip or .tar or .gz File To Upload On Server"
        .Filter = "Archive File|*.zip;*.tar;*.gz|All Files|*.*"
        .ShowOpen
        UploadLocalFileName = .FileName
    End With
    txtUploadLocalFileName = UploadLocalFileName
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmUpload.btnZipFile_Click", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub

Private Sub cboCategory_Click()

On Error GoTo Debugger
    
    Category = cboCategory.ItemData(cboCategory.ListIndex)
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmUpload.cboCategory_Click", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub


Private Sub chkCompat_Click(Index As Integer)

On Error GoTo Debugger

Select Case Index
    
    Case 1
            Compat1 = 1
    
    Case 2
            Compat2 = 2
    
    Case 3
            Compat3 = 3
    
    Case 4
            Compat4 = 4
    
    Case 5
            Compat5 = 27
    
    Case 7
            Compat7 = 5
    
    Case 8
            Compat8 = 6
    
    Case 9
            Compat9 = 37
    
    Case 10
            Compat10 = 38
    
End Select
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmUpload.chkCompat(" & Index & ")_Click", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
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
    
    Unload frmCode
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmUpload.Form_Unload", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub

Private Sub optDifficultyTypeId_Click(Index As Integer)

On Error GoTo Debugger
    
    DifficultyId = Index
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmUpload.optDifficultyTypeId_Click", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub
