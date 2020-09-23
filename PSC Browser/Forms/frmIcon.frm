VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmCommon 
   Caption         =   "Just For Common Stuff"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmIcon.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin InetCtlsObjects.Inet Inet 
      Left            =   720
      Top             =   330
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog cmd 
      Left            =   2100
      Top             =   1350
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmCommon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error GoTo Debugger
    Me.Hide
    
Debugger:
    If Not Err.Number = 0 And Not Err.Number = -2147467259 Then
        Select Case ErrLog(Err.Number, Error$, "frmCommon.Form_Load", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If


End Sub

Private Sub Inet_StateChanged(ByVal State As Integer)

On Error GoTo Debugger

InetState = State

        
        
    
    
Debugger:
    If Not Err.Number = 0 And Not Err.Number = -2147467259 Then
        Select Case ErrLog(Err.Number, Error$, "frmCommon.Inet_StateChanged", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If


End Sub
