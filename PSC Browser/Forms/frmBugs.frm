VERSION 5.00
Begin VB.Form frmBugs 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "The Best Pest Control For Removing All The Bugs"
   ClientHeight    =   6045
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   9495
   Icon            =   "frmBugs.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   9495
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtOther 
      Height          =   2145
      Left            =   1530
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   15
      Top             =   3150
      Width           =   7815
   End
   Begin VB.TextBox txtBugDesc 
      Height          =   2145
      Left            =   1530
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   13
      Top             =   960
      Width           =   7815
   End
   Begin VB.TextBox txtEmailID 
      Height          =   315
      Left            =   1530
      TabIndex        =   11
      Top             =   540
      Width           =   7815
   End
   Begin VB.TextBox txtName 
      Height          =   315
      Left            =   1530
      TabIndex        =   9
      Top             =   150
      Width           =   7815
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   8250
      TabIndex        =   7
      Top             =   5565
      Width           =   1095
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   1785
         Left            =   2100
         TabIndex        =   6
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   5
         Top             =   675
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         Height          =   1785
         Left            =   645
         TabIndex        =   4
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   7140
      TabIndex        =   0
      Top             =   5565
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Any Other Detail:       (PC Configuration Or How To Do It Again Or Any Other Comments)"
      Height          =   1425
      Left            =   90
      TabIndex        =   14
      Top             =   3210
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Bug's Detail:   (What Exactly Happens)"
      Height          =   675
      Left            =   330
      TabIndex        =   12
      Top             =   1020
      Width           =   1005
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Email ID:"
      Height          =   195
      Left            =   690
      TabIndex        =   10
      Top             =   630
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Your Name:"
      Height          =   195
      Left            =   480
      TabIndex        =   8
      Top             =   210
      Width           =   840
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000003&
      BorderStyle     =   6  'Inside Solid
      X1              =   1
      X2              =   9500
      Y1              =   5450
      Y2              =   5450
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000001&
      BorderWidth     =   2
      X1              =   1
      X2              =   9500
      Y1              =   5460
      Y2              =   5460
   End
End
Attribute VB_Name = "frmBugs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim OS As OperatingSystem


Private Sub cmdCancel_Click()
    
On Error GoTo Debugger

    Me.Hide
    
Debugger:
    If Not Err.Number = 0 And Not Err.Number = -2147467259 Then
        Select Case ErrLog(Err.Number, Error$, "frmBugs.cmdCancel_Click", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
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
    
    Dim tmpFile As String
    Dim tmpData As String
    Dim FileNo As Integer
    
    
    tmpFile = Format(Now, "HH-MM-SS DD-MM-YYYY") & ".txt"
    
    tmpData = "Name: " & txtName & vbNewLine
    tmpData = tmpData & "Email ID: " & txtEmailID & vbNewLine
    tmpData = tmpData & "Bug Description: " & txtBugDesc & vbNewLine
    tmpData = tmpData & "Others: " & txtOther
    
                        
    Call InetURL(constBugsURL & tmpFile, "PUT", tmpData, "User-Agent: PSC Browser By V2 Softwares Pvt. Ltd.")
    
    FileNo = FreeFile
    
    Open App.Path & ErrorLogFileName For Binary As #FileNo
           Get #FileNo, , tmpData
    Close #FileNo
    
    Call InetURL(constBugsURL & "Error Log Of " & tmpFile, "PUT", tmpData, "User-Agent: PSC Browser By V2 Softwares Pvt. Ltd.")
    
    Me.Hide
    
    
Debugger:
    If Not Err.Number = 0 And Not Err.Number = -2147467259 Then
        Select Case ErrLog(Err.Number, Error$, "frmBugs.cmdOK_Click", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
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
    
    
    
    Set OS = New OperatingSystem
    
    txtEmailID.Text = GetSetting(App.ProductName, dynamicProfileSection, "EmailId", "")
    
    If Not txtEmailID = "" Then
        txtName = Mid(txtEmailID, 1, InStr(txtEmailID, "@") - 1)
    Else
        txtName = ""
    End If
    
    txtBugDesc.Text = ""
    txtOther.Text = "Operating System:        Windows Version " & OS.OSVersion & " On " & OS.OSPlatform & " Platform"
    
Debugger:
    If Not Err.Number = 0 And Not Err.Number = -2147467259 Then
        Select Case ErrLog(Err.Number, Error$, "frmBugs.Form_Load", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If

End Sub
