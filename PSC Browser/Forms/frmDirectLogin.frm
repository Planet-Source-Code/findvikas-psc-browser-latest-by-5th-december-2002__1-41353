VERSION 5.00
Begin VB.Form frmDirectLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Direct Login [ The Hacker's Way ]"
   ClientHeight    =   975
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3750
   Icon            =   "frmDirectLogin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   65
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtID 
      Height          =   345
      Left            =   1290
      TabIndex        =   1
      Top             =   135
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   315
      Left            =   2010
      TabIndex        =   2
      Top             =   570
      Width           =   780
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   315
      Left            =   2820
      TabIndex        =   3
      Top             =   570
      Width           =   780
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Identification &ID:"
      Height          =   195
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   210
      Width           =   1155
   End
End
Attribute VB_Name = "frmDirectLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdCancel_Click()

On Error GoTo Debugger
    
    Me.Hide
    frmLogin.Show
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmDirectLogin.cmdCancel_Click", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub

Private Sub cmdOK_Click()

On Error GoTo Debugger
    
    
    Dim tmpData As String
        
    cmdOK.Enabled = False
    cmdCancel.Enabled = False
    txtID.Enabled = False
    
    If txtID.Text = "" Or Val(txtID) = 0 Then
        MsgBox "Please Enter A Valid ID!", vbExclamation, App.ProductName & " - [ Login ]"
        cmdOK.Enabled = True
        cmdCancel.Enabled = True
        txtID.Enabled = True
    Else
    
        tmpData = InetURL(constDirectLoginURL, "POST", "lngWId=&lngIdentId=" & txtID.Text)
    
        If tmpData = "" Then
            MsgBox "Cannot Log You On", vbExclamation, App.ProductName & " - [ Login ]"
            cmdOK.Enabled = True
            cmdCancel.Enabled = True
            txtID.Enabled = True
        Else
            MsgBox "LogOn Successful" & vbNewLine & "Return Contents Are As Follows: " & vbNewLine & tmpData, vbInformation
            Me.Hide
        End If
    
    End If
    
        
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmDirectLogin.cmdOK_Click", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub

