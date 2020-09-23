VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLiveUpdate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Live Update"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   6495
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5220
      TabIndex        =   3
      Top             =   2430
      Width           =   1155
   End
   Begin VB.CommandButton btnUpdate 
      Caption         =   "&Update"
      Height          =   375
      Left            =   4050
      TabIndex        =   2
      Top             =   2430
      Width           =   1155
   End
   Begin MSComctlLib.ProgressBar pbar 
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   6225
      _ExtentX        =   10980
      _ExtentY        =   344
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status: Press Update Button To Change You Status"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   2100
      Width           =   3675
   End
   Begin VB.Label lblProgress 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Progress: 0 % Done"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   1860
      Width           =   1395
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmLiveUpdate.frx":0000
      Height          =   1335
      Left            =   270
      TabIndex        =   1
      Top             =   60
      Width           =   5970
   End
End
Attribute VB_Name = "frmLiveUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FileToSave As String
Dim FileNo As Integer

Private Sub btnCancel_Click()
        
On Error GoTo Debugger

        Me.Hide
        frmCommon.Inet.Cancel
        
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmLiveUpdate.btnCancel_Click", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If

End Sub

Private Sub btnUpdate_Click()

On Error GoTo Debugger

Dim BinaryData() As Byte
Dim tmpURL As String
Dim starttime, endtime
    
      btnUpdate.Enabled = False
      btnCancel.Enabled = False
      
      lblStatus = "Status: " & "Connecting..."
      
      tmpURL = constLiveUpdateURL & "/psc_redirect.htm"
      
      Call frmCommon.Inet.Execute(tmpURL, "GET")    ' get the filename to download from server
      
      While frmCommon.Inet.StillExecuting
        DoEvents            'wait for get response
      Wend
      
      tmpURL = constLiveUpdateURL & frmCommon.Inet.GetChunk(255, icString) ' change the url to new url as server redirect us telling the latest codes filename
      
      If tmpURL = constLiveUpdateURL Then
        lblStatus = "Status: No Internet Connection."
        Call MsgBox("Probably You Are Not Connected To Internet. Try Again Later", vbInformation)
        GoTo exit_sub
      End If
      
      Call frmCommon.Inet.Execute(tmpURL, "GET")        'get the new file the actual source code file
      
      While frmCommon.Inet.StillExecuting
        DoEvents            'wait for get response
      Wend
      
      
      If InStr(frmCommon.Inet.GetHeader, "404") Then MsgBox "File Has Not Been Been Found Or Server. Try Online Help on Contact The Developer", vbExclamation: Exit Sub  ' to ensure that the updated version is present
      If Not frmCommon.Inet.GetHeader("Content-length:") = "" Then
            pbar.Max = Val(frmCommon.Inet.GetHeader("Content-length:"))
            pbar.Min = 0
            pbar.Value = 1
      Else
            MsgBox "Possibly You Are Not Connected To Internet. Check You Internet Connection First Then Retry Later", vbInformation
            GoTo exit_sub
      End If
      
    With frmCommon.cmd
        .DialogTitle = "Where To Save The Latest Sources"
        .Filter = "Any File|*.*"
        .CancelError = False
        .ShowOpen
        FileToSave = .FileName
    End With
    If FileToSave = "" Then GoTo exit_sub
    FileNo = FreeFile
    
    Open FileToSave For Binary As #FileNo
      
      Do
          starttime = TimeValue(Time)
          BinaryData() = frmCommon.Inet.GetChunk(14336, icByteArray)    'get 14 KB at a time
          endtime = TimeValue(Time)
          
          Put #FileNo, , BinaryData
          If pbar.Value + 14336 < pbar.Max Then
            pbar.Value = pbar.Value + 14336
          Else
            pbar.Value = pbar.Max
          End If
          lblProgress.Caption = "Progress: " & Int((pbar.Value * 100) / pbar.Max) & " % Done"
          lblStatus = "Status: Getting Latest File(s)."
      Loop Until Not frmCommon.Inet.StillExecuting And pbar.Value = pbar.Max
      
    
    Close #FileNo
    lblStatus = "Status: Done."
    MsgBox "Congrats You Have Successfully Update To Latest Sources", vbInformation
    pbar.Value = 0
exit_sub:
    btnUpdate.Enabled = True
    btnCancel.Enabled = True
      
    'You Can uncomment it for openning the downloaded file after successful download
    'Call ShellExecute(Me.hwnd, "OPEN", FileToSave, "", "", 0)
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "frmLiveUpdate.btnUpdate_Click", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Sub

