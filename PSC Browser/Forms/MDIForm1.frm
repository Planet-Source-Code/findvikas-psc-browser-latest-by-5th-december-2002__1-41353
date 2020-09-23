VERSION 5.00
Begin VB.MDIForm MDIBrowser 
   BackColor       =   &H80000003&
   Caption         =   "PSC Browser By V2 Softwares Pvt. Ltd."
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuNewWindow 
      Caption         =   "&Show The Browser"
   End
   Begin VB.Menu mnuSearchEngine 
      Caption         =   "&Show The Search Engine"
   End
End
Attribute VB_Name = "MDIBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub mnuNewWindow_Click()

On Error GoTo Debugger
    
    Dim nw As frmBrowser
    Set nw = New frmBrowser
    nw.Show
        
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "MDIBrowser.mnuNewWindow_Click", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    

End Sub

Private Sub mnuSearchEngine_Click()

On Error GoTo Debugger
    
    frmMain.Show
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "MDIBrowser.mnuSearchEngine_Click", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub
