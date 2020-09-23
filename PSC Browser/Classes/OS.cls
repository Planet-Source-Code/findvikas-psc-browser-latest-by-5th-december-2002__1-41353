VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OperatingSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type



Public Function OSVersion() As String

On Error GoTo Debugger

Dim OSINFO As OSVERSIONINFO
    OSINFO.dwOSVersionInfoSize = Len(OSINFO)
    Call GetVersionEx(OSINFO)
    OSVersion = OSINFO.dwMajorVersion & "." & OSINFO.dwMinorVersion & " Build " & OSINFO.dwBuildNumber
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "OperatingSystem.OSVersion", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Function
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Function

Public Function OSPlatform() As String

On Error GoTo Debugger

Dim OSINFO As OSVERSIONINFO
    OSINFO.dwOSVersionInfoSize = Len(OSINFO)
    Call GetVersionEx(OSINFO)
    
    Select Case OSINFO.dwPlatformId
        Case 1
            OSPlatform = "Intel"
            
        Case Else
            OSPlatform = "Unknown"
            
    End Select
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "OperatingSystem.OSPlatform", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Function
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
    
End Function


Private Sub Class_Initialize()

On Error GoTo Debugger

     'Set OSINFO = OSVERSIONINFO

Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "OperatingSystem.Class_Initialize", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Sub
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
End Sub
