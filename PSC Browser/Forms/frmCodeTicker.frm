VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmCodeTicker 
   BackColor       =   &H00FF6400&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Latest Code Ticker"
   ClientHeight    =   3015
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   3495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   201
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   233
   Begin SHDocVwCtl.WebBrowser brwCode 
      Height          =   2400
      Left            =   30
      TabIndex        =   0
      Top             =   120
      Width           =   4635
      ExtentX         =   8176
      ExtentY         =   4233
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   -1  'True
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "frmCodeTicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  


Private Sub brwCode_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)

On Error GoTo Debugger

If Progress <= 0 Or ProgressMax <= 0 Then Exit Sub

    Me.Caption = "Latest Code Ticker - [ " & Int((Progress * 100) / ProgressMax) & " % Done" & " ] "

Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmCodeTicker.brwCode_ProgressChange", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
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

    brwCode.Navigate constCodeTickerURL
    
    ScreenPosition = GetSetting(App.ProductName, dynamicSettingSection & "\Forms\frmCodeTicker", "ScreenPosition", "")
    If ScreenPosition = "" Then Exit Sub
    Me.Left = Val(Left(ScreenPosition, InStr(ScreenPosition, ",") - 1)): ScreenPosition = Right(ScreenPosition, Len(ScreenPosition) - InStr(ScreenPosition, ","))
    Me.Top = Val(Left(ScreenPosition, InStr(ScreenPosition, ",") - 1)): ScreenPosition = Right(ScreenPosition, Len(ScreenPosition) - InStr(ScreenPosition, ","))
    Me.Width = Val(Left(ScreenPosition, InStr(ScreenPosition, ",") - 1)): ScreenPosition = Right(ScreenPosition, Len(ScreenPosition) - InStr(ScreenPosition, ","))
    Me.Height = Val(ScreenPosition): ScreenPosition = ""
    
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmCodeTicker.Form_Load", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
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

brwCode.Move 5, 5, ScaleWidth - 10, ScaleHeight - 10

Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmCodeTicker.Form_Resize", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
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
ScreenPosition = Me.Left & "," & Me.Top & "," & Me.Width & "," & Me.Height
Call SaveSetting(App.ProductName, dynamicSettingSection & "\Forms\frmCodeTicker", "ScreenPosition", ScreenPosition)
End Sub
