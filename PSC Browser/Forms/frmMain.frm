VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PSC Browser By V2 Softwares Pvt. Ltd."
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   10695
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   474
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   713
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnMoreLess 
      Caption         =   "Search"
      Height          =   345
      Left            =   3360
      TabIndex        =   33
      Top             =   1350
      Width           =   1905
   End
   Begin VB.Frame Experience 
      BackColor       =   &H80000000&
      Caption         =   "Difficulity Level Options"
      Height          =   1365
      Left            =   120
      TabIndex        =   5
      Top             =   4650
      Width           =   5145
      Begin VB.CheckBox chkDifficultyTypeId 
         BackColor       =   &H80000000&
         Caption         =   "Unranked"
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   9
         Top             =   270
         Width           =   2685
      End
      Begin VB.CheckBox chkDifficultyTypeId 
         BackColor       =   &H80000000&
         Caption         =   "Beginner (0-6 months experience)"
         Height          =   195
         Index           =   2
         Left            =   210
         TabIndex        =   8
         Top             =   510
         Width           =   2685
      End
      Begin VB.CheckBox chkDifficultyTypeId 
         BackColor       =   &H80000000&
         Caption         =   "Advanced (1+ years) "
         Height          =   195
         Index           =   4
         Left            =   210
         TabIndex        =   7
         Top             =   990
         Width           =   2685
      End
      Begin VB.CheckBox chkDifficultyTypeId 
         BackColor       =   &H80000000&
         Caption         =   "Intermediate (6 months - 1 year)"
         Height          =   195
         Index           =   3
         Left            =   210
         TabIndex        =   6
         Top             =   750
         Width           =   2685
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000000&
      Caption         =   "Code Type Options"
      Height          =   1485
      Left            =   120
      TabIndex        =   0
      Top             =   3120
      Width           =   5145
      Begin VB.CheckBox chkCode3rdPartyReview 
         BackColor       =   &H80000000&
         Caption         =   "3rd Party Review"
         Height          =   225
         Left            =   210
         TabIndex        =   4
         Top             =   1095
         Width           =   2475
      End
      Begin VB.CheckBox chkCodeTypeArticle 
         BackColor       =   &H80000000&
         Caption         =   "Articles / Tutorials"
         Height          =   225
         Left            =   210
         TabIndex        =   3
         Top             =   810
         Width           =   2475
      End
      Begin VB.CheckBox chkCodeTypeText 
         BackColor       =   &H80000000&
         Caption         =   "'Copy-and-paste' source code "
         Height          =   225
         Left            =   210
         TabIndex        =   2
         Top             =   525
         Width           =   2475
      End
      Begin VB.CheckBox chkCodeTypeZip 
         BackColor       =   &H80000000&
         Caption         =   "Zip Files "
         Height          =   225
         Left            =   210
         TabIndex        =   1
         Top             =   240
         Width           =   2475
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000000&
      Caption         =   "Misc. Options"
      Height          =   4005
      Left            =   5370
      TabIndex        =   15
      Top             =   2010
      Width           =   5145
      Begin VB.CheckBox chkAdmAccess 
         BackColor       =   &H80000000&
         Caption         =   "Is Super Admin Access On"
         Height          =   300
         Left            =   1830
         TabIndex        =   44
         Top             =   3345
         Width           =   2685
      End
      Begin VB.CheckBox chkEditCode 
         BackColor       =   &H80000000&
         Caption         =   "Edit Code [ Experimental ]"
         Height          =   300
         Left            =   1830
         TabIndex        =   43
         Top             =   3030
         Width           =   2685
      End
      Begin VB.CheckBox chkNewestCode 
         BackColor       =   &H80000000&
         Caption         =   "Search Newest Code"
         Height          =   300
         Left            =   1830
         TabIndex        =   42
         Top             =   2715
         Width           =   2685
      End
      Begin VB.CheckBox chkTopCode 
         BackColor       =   &H80000000&
         Caption         =   "Search Top Code"
         Height          =   300
         Left            =   1830
         TabIndex        =   41
         Top             =   2400
         Width           =   2685
      End
      Begin VB.CheckBox chkThoroughSearch 
         BackColor       =   &H80000000&
         Caption         =   "Scans actual code contents."
         Height          =   300
         Left            =   1830
         TabIndex        =   40
         Top             =   3660
         Width           =   2685
      End
      Begin VB.TextBox txtMaxEntries 
         BackColor       =   &H00D8E9EC&
         ForeColor       =   &H00FF6400&
         Height          =   315
         Left            =   1830
         TabIndex        =   38
         Text            =   "20"
         Top             =   2010
         Width           =   3075
      End
      Begin VB.TextBox txtEntriesEnd 
         BackColor       =   &H00D8E9EC&
         ForeColor       =   &H00FF6400&
         Height          =   345
         Left            =   3600
         TabIndex        =   36
         Text            =   "20"
         Top             =   1590
         Width           =   1305
      End
      Begin VB.TextBox txtEntriesStart 
         BackColor       =   &H00D8E9EC&
         ForeColor       =   &H00FF6400&
         Height          =   345
         Left            =   1830
         TabIndex        =   35
         Text            =   "1"
         Top             =   1590
         Width           =   1305
      End
      Begin VB.ComboBox cboReset 
         BackColor       =   &H00D8E9EC&
         ForeColor       =   &H00FF6400&
         Height          =   315
         ItemData        =   "frmMain.frx":058A
         Left            =   1830
         List            =   "frmMain.frx":0594
         TabIndex        =   19
         Text            =   "TRUE"
         Top             =   1200
         Width           =   3075
      End
      Begin VB.ComboBox cboWorldDropDown 
         BackColor       =   &H00D8E9EC&
         ForeColor       =   &H00FF6400&
         Height          =   315
         ItemData        =   "frmMain.frx":05A5
         Left            =   1830
         List            =   "frmMain.frx":05AF
         TabIndex        =   18
         Text            =   "TRUE"
         Top             =   825
         Width           =   3075
      End
      Begin VB.ComboBox cboSortOrder 
         BackColor       =   &H00D8E9EC&
         ForeColor       =   &H00FF6400&
         Height          =   315
         ItemData        =   "frmMain.frx":05C0
         Left            =   1830
         List            =   "frmMain.frx":05D0
         TabIndex        =   16
         Text            =   "Alphabetical"
         Top             =   450
         Width           =   3075
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Entries Per Page:"
         Height          =   195
         Left            =   330
         TabIndex        =   39
         Top             =   2040
         Width           =   1230
      End
      Begin VB.Label Label9 
         Caption         =   "to"
         Height          =   195
         Left            =   3300
         TabIndex        =   37
         Top             =   1680
         Width           =   135
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Show Entries From #"
         Height          =   195
         Left            =   120
         TabIndex        =   34
         Top             =   1650
         Width           =   1470
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reset All Variables:"
         Height          =   195
         Left            =   180
         TabIndex        =   21
         Top             =   1320
         Width           =   1365
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "World Drop Down:"
         Height          =   195
         Left            =   240
         TabIndex        =   20
         Top             =   930
         Width           =   1320
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sorting Order:"
         Height          =   195
         Left            =   570
         TabIndex        =   17
         Top             =   570
         Width           =   975
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Nedded To Query The PSC Database"
      Height          =   1155
      Left            =   120
      TabIndex        =   22
      Top             =   60
      Width           =   5145
      Begin VB.TextBox txtCriteria 
         BackColor       =   &H00D8E9EC&
         ForeColor       =   &H00FF6400&
         Height          =   315
         Left            =   1920
         TabIndex        =   24
         Text            =   """PSC Browser"""
         Top             =   630
         Width           =   2775
      End
      Begin VB.ComboBox cboLanguage 
         BackColor       =   &H00D8E9EC&
         ForeColor       =   &H00FF6400&
         Height          =   315
         ItemData        =   "frmMain.frx":0636
         Left            =   1920
         List            =   "frmMain.frx":0658
         TabIndex        =   23
         Text            =   "Please Select A Language..."
         Top             =   240
         Width           =   3075
      End
      Begin VB.Image btnGo 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   4740
         Picture         =   "frmMain.frx":06E4
         Top             =   660
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Language:"
         Height          =   195
         Left            =   300
         TabIndex        =   26
         Top             =   330
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search For:"
         Height          =   195
         Left            =   240
         TabIndex        =   25
         Top             =   690
         Width           =   825
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000000&
      Caption         =   "Author Search Options"
      Height          =   1095
      Left            =   120
      TabIndex        =   10
      Top             =   1980
      Width           =   5145
      Begin VB.TextBox txtAuthorID 
         BackColor       =   &H00D8E9EC&
         ForeColor       =   &H00FF6400&
         Height          =   315
         Left            =   1890
         TabIndex        =   14
         Text            =   "3317434212"
         Top             =   630
         Width           =   3075
      End
      Begin VB.TextBox txtAuthorName 
         BackColor       =   &H00D8E9EC&
         ForeColor       =   &H00FF6400&
         Height          =   315
         Left            =   1890
         TabIndex        =   13
         Text            =   "newbornhacker"
         Top             =   270
         Width           =   3075
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Author ID:"
         Height          =   195
         Left            =   420
         TabIndex        =   12
         Top             =   690
         Width           =   720
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Author Name:"
         Height          =   195
         Left            =   180
         TabIndex        =   11
         Top             =   330
         Width           =   975
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "By V2 Softwares Pvt. Ltd."
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   2
      Left            =   5910
      TabIndex        =   29
      Top             =   1290
      Width           =   4125
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PSC Browser"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   585
      Index           =   0
      Left            =   6090
      TabIndex        =   27
      Top             =   120
      Width           =   3465
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PSC Browser"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF6400&
      Height          =   600
      Index           =   1
      Left            =   6120
      TabIndex        =   28
      Top             =   150
      Width           =   3465
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "By V2 Softwares Pvt. Ltd."
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF6400&
      Height          =   330
      Index           =   3
      Left            =   5880
      TabIndex        =   30
      Top             =   1260
      Width           =   4125
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version #:#:####"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF6400&
      Height          =   240
      Index           =   5
      Left            =   6930
      TabIndex        =   32
      Top             =   840
      Width           =   1920
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version #:#:####"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   240
      Index           =   4
      Left            =   6900
      TabIndex        =   31
      Top             =   810
      Width           =   1920
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuChangeUser 
         Caption         =   "Change &User / Login"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuCreateNewUser 
         Caption         =   "Create &New User"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCodeTicker 
         Caption         =   "Latest Code &Ticker"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUpload 
         Caption         =   "U&pload Code To Site"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLogViewer 
         Caption         =   "Log &Viewer"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuConfig 
         Caption         =   "Con&figuration"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuPSCLinks 
      Caption         =   "PSC Links"
      Begin VB.Menu mnuCodeOfDay 
         Caption         =   "Code Of The &Day"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuFeedback 
         Caption         =   "Feed&back To Site"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuLink 
         Caption         =   "&Link to the Site"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuAwards 
         Caption         =   "&Awards By Site"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuAdvertising 
         Caption         =   "A&dvertising To Site"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuPrivacy 
         Caption         =   "Privacy P&olicy Of Site"
         Shortcut        =   ^Q
      End
      Begin VB.Menu mnuTerms 
         Caption         =   "Te&rms and Conditions"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAboutSite 
         Caption         =   "About The &Site"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuOnlineHelp 
         Caption         =   "&Online Help"
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu mnuEmailHelp 
         Caption         =   "Get Help By &E-Mail"
         Shortcut        =   +{F1}
      End
      Begin VB.Menu mnuBugReport 
         Caption         =   "Report A Bu&g"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLiveUpdate 
         Caption         =   "Live Update"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuSep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnGo_Click()

On Error GoTo Debugger
    
    JumpURL CreateURL
        
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmMain.btnGo_Click", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    

End Sub

Private Sub btnGo_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

On Error GoTo Debugger
    
    btnGo.Move btnGo.Left + 15, btnGo.Top + 15
    btnGo.BorderStyle = 1
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmMain.btnGo_MouseDown", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub

Private Sub btnGo_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

On Error GoTo Debugger
    
    btnGo.Move btnGo.Left - 15, btnGo.Top - 15
    btnGo.BorderStyle = 0
    
    
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmMain.btnGo_MouseUp", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub

Public Sub btnMoreLess_Click()

On Error GoTo Debugger
    
If btnMoreLess.Caption = "&Advanced Search" Then
    btnMoreLess.Caption = "&Basic Search"
    Me.Height = 6825
    Me.Move (Screen.Width / 2) - (Me.Width / 2), (Screen.Height / 2) - (Me.Height / 2)
    AdvancedSearch = True
Else
        btnMoreLess.Caption = "&Advanced Search"
        Me.Height = 2550
        Me.Move (Screen.Width / 2) - (Me.Width / 2), (Screen.Height / 2) - (Me.Height / 2)
        AdvancedSearch = False
End If
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmMain.btnMoreLess_Click", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub

Private Sub cboLanguage_Click()

On Error GoTo Debugger
    
    LanguageID = cboLanguage.ListIndex + 1
    If LanguageID <= 0 Or LanguageID > 10 Then LanguageID = -1
    
    FinalURL = dynamicSearchURL & "lngWId=" & LanguageID
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmMain.cboLanguage_Click", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub

Private Sub cboLanguage_KeyPress(KeyAscii As Integer)

On Error GoTo Debugger
    
    If KeyAscii = vbKeyReturn Then btnGo_Click

Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmMain.cboLanguage_KeyPress", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
End Sub

Private Sub cboLanguage_LostFocus()

On Error GoTo Debugger
    
    LanguageID = cboLanguage.ListIndex + 1
    If LanguageID <= 0 Or LanguageID > 10 Then LanguageID = -1
    
    FinalURL = dynamicSearchURL & "lngWId=" & LanguageID
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmMain.cboLanguage_LostFocus", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    

End Sub

Private Sub cboReset_Click()

On Error GoTo Debugger
    
    ResetVariables = cboReset.List(cboReset.ListIndex)
    If ResetVariables = "" Then ResetVariables = "FALSE"
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmMain.cboReset_Click", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    

End Sub

Private Sub cboReset_KeyPress(KeyAscii As Integer)

On Error GoTo Debugger
    
If KeyAscii = vbKeyReturn Then btnGo_Click
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmMain.cboReset_KeyPress", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub

Private Sub cboReset_LostFocus()

On Error GoTo Debugger
    
    ResetVariables = cboReset.List(cboReset.ListIndex)
    If ResetVariables = "" Then ResetVariables = "TRUE"
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmMain.cboReset_LostFocus", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    

End Sub

Private Sub cboSortOrder_Click()

On Error GoTo Debugger
    
Select Case cboSortOrder.ListIndex
    
    Case 0
        SortOrder = "Alphabetical"
        
    Case 1
        SortOrder = "DateDescending"
        
    Case 2
        SortOrder = "DateAscending"
        
    Case 3
        SortOrder = "CountDescending"
        

End Select


If SortOrder = "" Then SortOrder = "Alphabetical"
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmMain.cboSortOrder_Click", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    

End Sub

Private Sub cboSortOrder_KeyPress(KeyAscii As Integer)

On Error GoTo Debugger
    
If KeyAscii = vbKeyReturn Then btnGo_Click
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmMain.cboSortOrder_KeyPress", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub

Private Sub cboSortOrder_LostFocus()

On Error GoTo Debugger
    
Select Case cboSortOrder.ListIndex
    
    Case 0
        SortOrder = "Alphabetical"
        
    Case 1
        SortOrder = "DateDescending"
        
    Case 2
        SortOrder = "DateAscending"
        
    Case 3
        SortOrder = "CountDescending"
        

End Select

If SortOrder = "" Then SortOrder = "Alphabetical"
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmMain.cboSortOrder_LostFocus", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    

End Sub

Private Sub cboWorldDropDown_Click()

On Error GoTo Debugger
    
    DropDown = cboWorldDropDown.List(cboWorldDropDown.ListIndex)
    If DropDown = "" Then DropDown = "TRUE"
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmMain.cboWorldDropDown_Click", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    

End Sub

Private Sub cboWorldDropDown_KeyPress(KeyAscii As Integer)

On Error GoTo Debugger
    
If KeyAscii = vbKeyReturn Then btnGo_Click
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmMain.cboWorldDropDown_KeyPress", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub

Private Sub cboWorldDropDown_LostFocus()

On Error GoTo Debugger
    
    DropDown = cboWorldDropDown.List(cboWorldDropDown.ListIndex)
    If DropDown = "" Then DropDown = "TRUE"
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmMain.cboWorldDropDown_LostFocus", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
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
    
    
    
    
    
    Me.Height = 2550


    Label3(4).Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    Label3(5).Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision

    LanguageID = cboLanguage.ListIndex + 1
    If LanguageID <= 0 Or LanguageID > 10 Then LanguageID = -1
    
    If Not txtCriteria.Text = "" Then TextToSearch = txtCriteria.Text
    
    DropDown = cboWorldDropDown.List(cboWorldDropDown.ListIndex)
    If DropDown = "" Then DropDown = "TRUE"

    MaxEntry = Val(txtMaxEntries.Text)
    If Not IsNumeric(MaxEntry) Or Not MaxEntry > 0 Then MaxEntry = 10

    SortOrder = cboSortOrder.List(cboSortOrder.ListIndex)
    If SortOrder = "" Then SortOrder = "Alphabetical"

    ResetVariables = cboReset.List(cboReset.ListIndex)
    If ResetVariables = "" Then ResetVariables = "TRUE"
    
    
    
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmMain.Form_Load", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
                    Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
    
End Sub




Private Sub mnuCodeOfDay_Click()

On Error GoTo Debugger
    
    JumpURL constCodeOfDayURL
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmMain.mnuCodeOfDay_Click", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub

Private Sub mnuCodeTicker_Click()

On Error GoTo Debugger
        
        frmCodeTicker.Show
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmMain.mnuCodeTicker_Click", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub

Private Sub mnuLiveUpdate_Click()
    frmLiveUpdate.Show
End Sub

Private Sub mnuLogViewer_Click()
    frmLogViewer.Show
End Sub

Private Sub mnuPrivacy_Click()

On Error GoTo Debugger
    
    JumpURL constPrivacyURL
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmMain.mnuPrivacy_Click", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub
Private Sub mnuLink_Click()

On Error GoTo Debugger
    
    JumpURL constLinkToSiteURL
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmMain.mnuLink_Click", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub
Private Sub mnuFeedback_Click()

On Error GoTo Debugger
    
    JumpURL constFeedBackURL
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmMain.mnuFeedback_Click", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub


Private Sub mnuTerms_Click()

On Error GoTo Debugger
    
    JumpURL constTermsURL
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmMain.mnuTerms_Click", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

On Error GoTo Debugger
    
'    If Button = 2 Then PopupMenu frmBrowser.mnuHelp
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmMain.Form_MouseDown", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
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
    
    Call SaveSetting(App.ProductName, dynamicURLSection, "SearchURL", dynamicSearchURL)
    Call SaveSetting(App.ProductName, dynamicURLSection, "NewUserURL", dynamicNewUserURL)
    Call SaveSetting(App.ProductName, dynamicURLSection, "ForgotPwdURL", dynamicForgotPwdURL)
    Call SaveSetting(App.ProductName, dynamicURLSection, "CodeOfDayURL", dynamicCodeOfDayURL)
    Call SaveSetting(App.ProductName, dynamicURLSection, "LogOnURL", dynamicLogOnURL)
    Call SaveSetting(App.ProductName, dynamicURLSection, "LogOnReturnURL", dynamicLogOnReturnURL)

    If Not Quit Then Cancel = True
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmMain.Form_Unload", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub

Private Sub mnuAbout_Click()

On Error GoTo Debugger
    
    frmAbout.Show
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmMain.mnuAbout_Click", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub

Private Sub mnuBugReport_Click()
On Error GoTo Debugger
    
    frmBugs.Show
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmMain.mnuBugReport_Click", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub

Private Sub mnuChangeUser_Click()

On Error GoTo Debugger
    
Me.Hide
frmLogin.Show
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmMain.mnuChangeUser_Click", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub

Private Sub mnuConfig_Click()

On Error GoTo Debugger
    
    frmConfig.Show
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmMain.mnuConfig_Click", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub

Private Sub mnuCreateNewUser_Click()

On Error GoTo Debugger
    
    frmLogin.btnNewUser_Click
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmMain.mnuCreateNewUser_Click", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub

Private Sub mnuEmailHelp_Click()

On Error GoTo Debugger
    
    Call ShellExecute(vbEmpty, "OPEN", "mailto:PlanetVBCode@yahoo.com?subject=Help On PSC Browser", vbEmpty, vbEmpty, vbEmpty)
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmMain.mnuEmailHelp_Click", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub

Private Sub mnuExit_Click()

On Error GoTo Debugger
    
Me.Hide
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmMain.mnuExit_Click", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub


Private Sub mnuOnlineHelp_Click()

On Error GoTo Debugger
    
    JumpURL constOnlineHelpURL
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmMain.mnuOnlineHelp_Click", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub

Private Sub mnuUpload_Click()

On Error GoTo Debugger
    
    frmUpload.Show
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmMain.mnuUpload_Click", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
End Sub



Private Sub txtCriteria_Change()

On Error GoTo Debugger
    
If Not txtCriteria.Text = "" Then TextToSearch = txtCriteria.Text
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmMain.txtCriteria_Change", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub

Private Sub txtCriteria_KeyPress(KeyAscii As Integer)

On Error GoTo Debugger
    
If KeyAscii = vbKeyReturn Then btnGo_Click
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmMain.txtCriteria_KeyPress", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub

Private Sub txtCriteria_LostFocus()

On Error GoTo Debugger
    
If Not txtCriteria.Text = "" Then TextToSearch = txtCriteria.Text
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmMain.txtCriteria_LostFocus", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    

End Sub

Private Sub txtMaxEntries_Change()

On Error GoTo Debugger
    
MaxEntry = Val(txtMaxEntries.Text)
If Not IsNumeric(MaxEntry) Or Not MaxEntry > 0 Then MaxEntry = 10
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmMain.txtMaxEntries_Change", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    

End Sub

Private Sub txtMaxEntries_KeyPress(KeyAscii As Integer)

On Error GoTo Debugger
    
If KeyAscii = vbKeyReturn Then btnGo_Click
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmMain.txtMaxEntries_KeyPress", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub

Private Sub txtMaxEntries_LostFocus()

On Error GoTo Debugger
    
MaxEntry = Val(txtMaxEntries.Text)
If Not IsNumeric(MaxEntry) Or Not MaxEntry > 0 Then MaxEntry = 10

    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmMain.txtMaxEntries_LostFocus", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub
Private Sub mnuAboutSite_Click()

On Error GoTo Debugger
    
    JumpURL constAboutSiteURL
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmMain.mnuAboutSite_Click", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub

Private Sub mnuAdvertising_Click()

On Error GoTo Debugger
    
    JumpURL constAdvertisingURL
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmMain.mnuAdvertising_Click", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub
Private Sub mnuAwards_Click()

On Error GoTo Debugger
    
    JumpURL constAwardsURL
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmMain.mnuAwards_Click", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub


