VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLogViewer 
   BackColor       =   &H8000000A&
   Caption         =   "Log Viewer"
   ClientHeight    =   7290
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8700
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   486
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   580
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList imglistBIG 
      Left            =   4080
      Top             =   3360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogViewer.frx":0000
            Key             =   "BIG_ErrorIcon"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogViewer.frx":0454
            Key             =   "BIG_InetIcon"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogViewer.frx":08A8
            Key             =   "BIG_URLIcon"
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox cboOptions 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmLogViewer.frx":0CFC
      Left            =   75
      List            =   "frmLogViewer.frx":0D0F
      TabIndex        =   0
      Text            =   "Select A Valid Option"
      Top             =   75
      Width           =   6765
   End
   Begin MSComctlLib.ImageList imglistSMALL 
      Left            =   3000
      Top             =   3330
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogViewer.frx":0D6C
            Key             =   "SMALL_ErrorIcon"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogViewer.frx":1308
            Key             =   "SMALL_InetIcon"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogViewer.frx":16A4
            Key             =   "SMALL_URLIcon"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6345
      Left            =   90
      TabIndex        =   1
      Top             =   870
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   11192
      View            =   3
      Arrange         =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      PictureAlignment=   5
      TextBackground  =   -1  'True
      _Version        =   393217
      Icons           =   "imglistBIG"
      SmallIcons      =   "imglistSMALL"
      ColHdrIcons     =   "imglistBIG"
      ForeColor       =   -2147483635
      BackColor       =   14215660
      BorderStyle     =   1
      Appearance      =   1
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmLogViewer.frx":1C40
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "URL"
         Object.Width           =   38100
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Date"
         Object.Width           =   13229
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Time"
         Object.Width           =   13229
      EndProperty
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000C&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   0
      X2              =   600
      Y1              =   35
      Y2              =   35
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   0
      X1              =   0
      X2              =   600
      Y1              =   36
      Y2              =   36
   End
   Begin VB.Menu popup 
      Caption         =   "popup"
      Begin VB.Menu mnuViewHolder 
         Caption         =   "&View"
         Begin VB.Menu mnuView 
            Caption         =   "Lar&ge Icons"
            Index           =   0
         End
         Begin VB.Menu mnuView 
            Caption         =   "S&mall Icons"
            Index           =   1
         End
         Begin VB.Menu mnuView 
            Caption         =   "&List"
            Index           =   2
         End
         Begin VB.Menu mnuView 
            Caption         =   "&Details"
            Index           =   3
         End
      End
      Begin VB.Menu mnuSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuArrangeIcons 
         Caption         =   "Arrange Icons"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "Re&fresh"
      End
   End
End
Attribute VB_Name = "frmLogViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cboOptions_Click()
    
On Error GoTo Debugger

    Select Case cboOptions.ListIndex
    
        Case 0              'error log
            ViewLog ("ERROR")
        Case 1              'inet log
            ViewLog ("INET")
        Case 2              'url log
            ViewLog ("URL")
        Case 3              'clear log
            If MsgBox("This Will Delete All Log Files" & vbNewLine & "Are You Sure You Want To Continue", vbQuestion + vbYesNo) = vbNo Then Exit Sub
            Kill App.Path & ErrorLogFileName
            Kill App.Path & InetURLLogFileName
            Kill App.Path & URLLogFileName
        Case 4              'exit
            Me.Hide
    
    End Select

Debugger:
    If Not Err.Number = 0 And Not Err.Number = 53 Then
        Select Case ErrLog(Err.Number, Error$, "frmLogViewer.cboOptions_Click", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
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

    ScreenPosition = GetSetting(App.ProductName, dynamicSettingSection & "\Forms\frmLogViewer", "ScreenPosition", "")
    If ScreenPosition = "" Then Exit Sub
    Me.Left = Val(Left(ScreenPosition, InStr(ScreenPosition, ",") - 1)): ScreenPosition = Right(ScreenPosition, Len(ScreenPosition) - InStr(ScreenPosition, ","))
    Me.Top = Val(Left(ScreenPosition, InStr(ScreenPosition, ",") - 1)): ScreenPosition = Right(ScreenPosition, Len(ScreenPosition) - InStr(ScreenPosition, ","))
    Me.Width = Val(Left(ScreenPosition, InStr(ScreenPosition, ",") - 1)): ScreenPosition = Right(ScreenPosition, Len(ScreenPosition) - InStr(ScreenPosition, ","))
    Me.Height = Val(ScreenPosition): ScreenPosition = ""

Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmLogViewer.Form_Load", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
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

    cboOptions.Move 5, 5, ScaleWidth - 10
    Line1(0).X1 = 0: Line1(0).X2 = ScaleWidth: Line1(0).Y1 = 41: Line1(0).Y2 = 41
    Line1(1).X1 = 0: Line1(1).X2 = ScaleWidth: Line1(1).Y1 = 40: Line1(1).Y2 = 40
    ListView1.Move 5, 45, ScaleWidth - 10, ScaleHeight - 50
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmLogViewer.Form_Resize", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
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

ScreenPosition = Me.Left & "," & Me.Top & "," & Me.Width & "," & Me.Height
Call SaveSetting(App.ProductName, dynamicSettingSection & "\Forms\frmLogViewer", "ScreenPosition", ScreenPosition)
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmLogViewer.Form_Unload", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
End Sub

Private Sub ListView1_DblClick()

On Error GoTo Debugger

Select Case cboOptions.ListIndex
    
    Case 0
        If Not ListView1.SelectedItem = "" Then MsgBox "Error     : " & ListView1.SelectedItem & vbNewLine & _
                                                       "Type    : " & ListView1.SelectedItem.SubItems(1) & vbNewLine & _
                                                       "Source : " & ListView1.SelectedItem.SubItems(2)
                                                
    Case 1
        If Not ListView1.SelectedItem = "" Then MsgBox "URL     : " & ListView1.SelectedItem & vbNewLine & _
                                                       "HTTPCommand    : " & ListView1.SelectedItem.SubItems(1) & vbNewLine & _
                                                       "Inputs : " & ListView1.SelectedItem.SubItems(2) & vbNewLine & _
                                                       "Input Headers : " & ListView1.SelectedItem.SubItems(3)

        
    Case 2
        If Not ListView1.SelectedItem = "" Then MsgBox "Url    : " & ListView1.SelectedItem & vbNewLine & _
                                                "Date : " & ListView1.SelectedItem.SubItems(1) & vbNewLine & _
                                                "Time : " & ListView1.SelectedItem.SubItems(2)
                                                
                                                
        
End Select

Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmLogViewer.ListView1_DblClick", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    

End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

On Error GoTo Debugger

If Button = 2 Then PopupMenu popup

Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmLogViewer.ListView1_MouseDown", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub

Private Sub mnuArrangeIcons_Click()
    
On Error GoTo Debugger

    ListView1.Arrange = lvwAutoLeft
    ListView1.Arrange = lvwAutoTop

Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmLogViewer.mnuArrangeIcons_Click", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub

Private Sub mnuView_Click(Index As Integer)

On Error GoTo Debugger

Select Case Index
    Case 0
        ListView1.View = lvwIcon
    
    Case 1
        ListView1.View = lvwSmallIcon

    Case 2
        ListView1.View = lvwList
        
    Case 3
        ListView1.View = lvwReport
            
End Select

Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmLogViewer.mnuView_Click", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub
