VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmBrowser 
   ClientHeight    =   5595
   ClientLeft      =   3345
   ClientTop       =   3435
   ClientWidth     =   7710
   Icon            =   "frmBrowser.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   373
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   514
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ProgressBar pbar 
      Align           =   2  'Align Bottom
      Height          =   195
      Left            =   0
      TabIndex        =   5
      Top             =   5400
      Visible         =   0   'False
      Width           =   7710
      _ExtentX        =   13600
      _ExtentY        =   344
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   7710
      _ExtentX        =   13600
      _ExtentY        =   1005
      ButtonWidth     =   1191
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Back"
            Key             =   "Back"
            Object.ToolTipText     =   "Back"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Forward"
            Key             =   "Forward"
            Object.ToolTipText     =   "Forward"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Stop"
            Key             =   "Stop"
            Object.ToolTipText     =   "Stop"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refresh"
            Key             =   "Refresh"
            Object.ToolTipText     =   "Refresh"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Home"
            Key             =   "Home"
            Object.ToolTipText     =   "Home"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Search"
            Key             =   "Search"
            Object.ToolTipText     =   "Search"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin SHDocVwCtl.WebBrowser brwWebBrowser 
      Height          =   3735
      Left            =   30
      TabIndex        =   1
      Top             =   1215
      Width           =   5400
      ExtentX         =   9525
      ExtentY         =   6588
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
      AutoArrange     =   -1  'True
      NoClientEdge    =   -1  'True
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Timer timTimer 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   3870
      Top             =   3840
   End
   Begin VB.PictureBox picAddress 
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   0
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   436
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   540
      Width           =   6540
      Begin VB.ComboBox cboAddress 
         BackColor       =   &H00D8E9EC&
         ForeColor       =   &H00FF6400&
         Height          =   315
         ItemData        =   "frmBrowser.frx":058A
         Left            =   720
         List            =   "frmBrowser.frx":058C
         TabIndex        =   0
         Text            =   "http://www.Planet-Source-Code.com"
         Top             =   30
         Width           =   3795
      End
      Begin VB.Image btnGo 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   4620
         Picture         =   "frmBrowser.frx":058E
         Stretch         =   -1  'True
         Top             =   90
         Width           =   240
      End
      Begin VB.Label lblAddress 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Address:"
         Height          =   195
         Left            =   60
         TabIndex        =   2
         Tag             =   "&Address:"
         Top             =   90
         Width           =   615
      End
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   2670
      Top             =   2325
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":0B18
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":180C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":2500
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":31F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":3EE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":4BDC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000C&
      BorderStyle     =   6  'Inside Solid
      X1              =   0
      X2              =   454
      Y1              =   29
      Y2              =   29
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000000&
      BorderWidth     =   2
      X1              =   0
      X2              =   454
      Y1              =   30
      Y2              =   30
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
         Shortcut        =   ^W
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuStop 
         Caption         =   "Sto&p                Esc"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "&Refresh"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuSource 
         Caption         =   "Sour&ce"
      End
   End
   Begin VB.Menu mnuFavorites 
      Caption         =   "&Favorites"
      Begin VB.Menu mnuAddToFavorites 
         Caption         =   "&Add To Favorites"
      End
      Begin VB.Menu mnuOrganizeFavorites 
         Caption         =   "&Organize Favorites"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuEmpty 
         Caption         =   ""
         Index           =   0
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public StartingAddress As String
Dim mbDontNavigateNow As Boolean

Private Sub brwWebBrowser_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)

On Error GoTo Debugger
        
    Dim FileNo As Integer
    
    FileNo = FreeFile

    LastURL = URL
    
    If Not cboAddress.Text = URL Then
        cboAddress.AddItem URL
        cboAddress.Text = URL
    End If
    
    Open App.Path & URLLogFileName For Append As #FileNo
        Print #FileNo, URL & Delimeter & Date & Delimeter & Time
    Close #FileNo
   Me.Caption = "Please Wait...."

Debugger:
    If Not Err.Number = 0 And Not Err.Number = -2147467259 Then
        Select Case ErrLog(Err.Number, Error$, "frmBrowser.brwWebBrowser_BeforeNavigate2", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub


Private Sub brwWebBrowser_DownloadBegin()

On Error GoTo Debugger
        
    pbar.Value = 0
    pbar.Visible = True

Debugger:
    If Not Err.Number = 0 And Not Err.Number = -2147467259 Then
        Select Case ErrLog(Err.Number, Error$, "frmBrowser.brwWebBrowser_DownloadBegin", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub

Private Sub brwWebBrowser_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)

On Error GoTo Debugger
        
        If Progress > 0 And ProgressMax > 0 Then
        pbar.Max = ProgressMax
    Else
        Exit Sub
    End If
    
    pbar.Min = 0
    pbar.Value = Progress
    Caption = Int((Progress * 100) / ProgressMax) & " % Complete"

Debugger:
    If Not Err.Number = 0 And Not Err.Number = -2147467259 Then
        Select Case ErrLog(Err.Number, Error$, "frmBrowser.brwWebBrowser_ProgressChange", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub

Private Sub btnGo_Click()

On Error GoTo Debugger
    
    cboAddress_Click
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmBrowser.btnGo_Click", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
                 Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub

Private Sub btnGo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error GoTo Debugger
    
    btnGo.Move btnGo.Left + 1, btnGo.Top + 1
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmBrowser.btnGo_MouseDown", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub

Private Sub btnGo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error GoTo Debugger
    
    btnGo.Move btnGo.Left - 1, btnGo.Top - 1
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmBrowser.btnGo_MouseUp", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

On Error GoTo Debugger
    
    
    Select Case KeyAscii
    Case vbKeyF5
        brwWebBrowser.Refresh
    Case vbKeyEscape
        brwWebBrowser.Stop
    
    End Select
    
    
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmBrowser.Form_KeyPress", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
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
    
    Dim ctr As Integer
    Dim i As Integer
    Dim mnuname
    Dim mnu
    
    Me.Show
    tbToolBar.Refresh
    
    cboAddress.Move 50, lblAddress.Top + lblAddress.Height + 15

    If Len(StartingAddress) > 0 Then
        cboAddress.Text = StartingAddress
        cboAddress.AddItem cboAddress.Text
        'try to navigate to the starting address
        timTimer.Enabled = True
        brwWebBrowser.Navigate StartingAddress
    End If
    
    ctr = GetSetting(App.ProductName, dynamicFavoriteURLSection, "Count", 0)
    
    For i = 1 To ctr
        
        mnuname = GetSetting(App.ProductName, dynamicFavoriteSection, i)
        Set mnu = Me.mnuEmpty
        Load mnu(i)
        mnu(i).Visible = True
        mnu(i).Caption = mnuname
        mnu(i).Enabled = True
    
    Next
    Form_Resize


    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmBrowser.Form_Load", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub



Private Sub brwWebBrowser_DownloadComplete()

On Error GoTo Debugger
    
    Me.Caption = brwWebBrowser.LocationName
    pbar.Value = 0
    pbar.Visible = False
    
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmBrowser.brwWebBrowser_DownloadComplete", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub

Private Sub brwWebBrowser_NavigateComplete2(ByVal pDisp As Object, URL As Variant)

On Error GoTo Debugger
    
    Dim i As Integer
    Dim bFound As Boolean
    Me.Caption = brwWebBrowser.LocationName
    For i = 0 To cboAddress.ListCount - 1
        If cboAddress.List(i) = brwWebBrowser.LocationURL Then
            bFound = True
            Exit For
        End If
    Next i
    mbDontNavigateNow = True
    If bFound Then
        cboAddress.RemoveItem i
    End If
    cboAddress.AddItem brwWebBrowser.LocationURL, 0
    cboAddress.ListIndex = 0
    mbDontNavigateNow = False
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmBrowser.brwWebBrowser_NavigateComplete2", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub

Private Sub cboAddress_Click()

On Error GoTo Debugger
    
    If mbDontNavigateNow Then Exit Sub
    timTimer.Enabled = True
    brwWebBrowser.Navigate cboAddress.Text
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmBrowser.cboAddress_Click", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub

Private Sub cboAddress_KeyPress(KeyAscii As Integer)

On Error GoTo Debugger
    
    
    Select Case KeyAscii
    Case vbKeyReturn
        cboAddress_Click
    Case vbKeyEscape
        brwWebBrowser.Stop
    
    End Select
    
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmBrowser.cboAddress_KeyPress", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub

Private Sub Form_Resize()


On Error GoTo Debugger
    
    
    If Not Me.WindowState = vbMinimized Then
    
        Line1.X1 = 0
        Line1.X2 = ScaleWidth
        Line1.Y1 = 30
        Line1.Y2 = 30
    
        Line2.X1 = 0
        Line2.X2 = ScaleWidth
        Line2.Y1 = 29
        Line2.Y2 = 29
    
        picAddress.Move 5, 35, ScaleWidth - btnGo.Width + 3
        lblAddress.Move 3, (picAddress.Height / 2) - (lblAddress.Height / 2)
        cboAddress.Move lblAddress.Left + lblAddress.Width + 5, lblAddress.Top - 2, picAddress.ScaleWidth - (lblAddress.Left + lblAddress.Width + 25)
        btnGo.Move cboAddress.Left + cboAddress.Width + 3, cboAddress.Top, 16, 16
    
        brwWebBrowser.Move 5, picAddress.Top + picAddress.Height + 3, ScaleWidth - 5, ScaleHeight - (brwWebBrowser.Top + 5)
    
    End If
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmBrowser.Form_Resize", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub



Private Sub mnuAddToFavorites_Click()

On Error GoTo Debugger
    
    Dim mnu
    
    Dim mnuname As String
    Dim ctr As Integer
    
    ctr = GetSetting(App.ProductName, dynamicFavoriteURLSection, "Count", 0) + 1
    
    mnuname = InputBox("Name: ", , brwWebBrowser.LocationName)
    
    If mnuname = "" Then Exit Sub
    
    Set mnu = Me.mnuEmpty
    
    Load mnu(ctr)
    mnu(ctr).Visible = True
    mnu(ctr).Caption = mnuname
    mnu(ctr).Enabled = True
    
    Call SaveSetting(App.ProductName, dynamicFavoriteSection, ctr, mnuname)
    Call SaveSetting(App.ProductName, dynamicFavoriteURLSection, "Count", ctr)
    Call SaveSetting(App.ProductName, dynamicFavoriteURLSection, ctr, LastURL)
    

Debugger:
    If Not Err.Number = 0 And Not Err.Number = -2147467259 Then
        Select Case ErrLog(Err.Number, Error$, "frmBrowser.mnuAddToFavorites_Click", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
    
End Sub



Private Sub mnuClose_Click()

On Error GoTo Debugger
    
    Me.Hide
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmBrowser.mnuClose_Click", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub

Private Sub mnuEmpty_Click(Index As Integer)

On Error GoTo Debugger
    
    Call JumpURL(GetSetting(App.ProductName, dynamicFavoriteURLSection, Index, "about:Page Not Found"))

Debugger:
    If Not Err.Number = 0 And Not Err.Number = -2147467259 Then
        Select Case ErrLog(Err.Number, Error$, "frmBrowser.mnuEmpty_Click", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub

Private Sub mnuOpen_Click()

On Error GoTo Debugger
    
    With frmCommon.cmd
        .DialogTitle = "Select A File To Open"
        .Filter = "Any File|*.*"
        .ShowOpen
        brwWebBrowser.Navigate2 (.FileName)
    End With

Debugger:
    If Not Err.Number = 0 And Not Err.Number = -2147467259 Then
        Select Case ErrLog(Err.Number, Error$, "frmBrowser.mnuOpen_Click", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub



Private Sub mnuRefresh_Click()

On Error GoTo Debugger
    
    brwWebBrowser.Refresh
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmBrowser.mnuRefresh_Click", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub

Private Sub mnuSource_Click()

On Error GoTo Debugger
    
Dim tmpData As String

    tmpData = frmCommon.Inet.OpenURL(LastURL)
    MsgBox tmpData
'    TO DO [ ADD CODE FOR SOURCE VIEWING ]
    
    
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmBrowser.mnuSource_Click", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub



Private Sub picAddress_KeyPress(KeyAscii As Integer)

On Error GoTo Debugger
    
    Select Case KeyAscii
    Case vbKeyF5
        brwWebBrowser.Refresh
    Case vbKeyEscape
        brwWebBrowser.Stop
    
    End Select
    
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmBrowser.picAddress_KeyPress", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    

End Sub

Private Sub timTimer_Timer()

On Error GoTo Debugger
    
    If brwWebBrowser.Busy = False Then
        timTimer.Enabled = False
        Me.Caption = brwWebBrowser.LocationName
    End If
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmBrowser.timTimer_Timer", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As Button)

On Error GoTo Debugger
    
     
    timTimer.Enabled = True
     
    Select Case Button.Key
        Case "Back"
            brwWebBrowser.GoBack
        Case "Forward"
            brwWebBrowser.GoForward
        Case "Refresh"
            brwWebBrowser.Refresh
        Case "Home"
            brwWebBrowser.GoHome
        Case "Search"
            frmMain.Show
            
        Case "Stop"
            timTimer.Enabled = False
            brwWebBrowser.Stop
            Me.Caption = brwWebBrowser.LocationName
    End Select

    
Debugger:
    If Not Err.Number = 0 And Not Err.Number = -2147467259 Then
        Select Case ErrLog(Err.Number, Error$, "frmBrowser.tbToolBar_ButtonClick", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub

