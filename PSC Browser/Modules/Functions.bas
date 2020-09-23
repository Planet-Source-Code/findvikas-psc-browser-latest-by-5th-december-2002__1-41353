Attribute VB_Name = "basFunctions"
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)



Global Const constSearchURL = "http://www.planet-source-code.com/vb/scripts/includes/TranslateFormVarsToURLString.asp?txtURL=/vb/scripts/BrowseCategoryOrSearchResults.asp?"
Global Const constAdvancedSearchURL = "http://www.planet-source-code.com/vb/scripts/includes/TranslateFormVarsToURLString.asp?txtURL=/vb/scripts/search.asp?"
Global Const constNewUserURL = "http://www.planet-source-code.com/vb/scripts/GetUserId.asp?lngWId=&txtReturnURL=http://planet-source-code.com/vb/authentication/Login.asp?"
Global Const constForgotPwdURL = "https://www.exhedra.com/ads/authentication/ForgotPassword.asp?lngWId=&txtReturnURL=http://planet-source-code.com/vb/authentication/Login.asp?"
Global Const constCodeOfDayURL = "http://www.planet-source-code.com/vb/CodeOfTheDay/signup.asp?lngWId=-1"
Global Const constLogOnURL = "https://www.exhedra.com/ads/authentication/LoginAction.asp?"
Global Const constLogOnReturnURL = "http://www.planet-source-code.com/vb/authentication/Login.asp?&txtReturnURL=/vb/authors/existing_author_choices.asp?lngWId=1"
Global Const constUpLoadURL = "http://www.planet-source-code.com/vb/Authors/Submit/SubmitAction.asp?"
Global Const constCodeTickerURL = "http://www.Planet-Source-Code.com/vb/linktous/ScrollingCode.asp?lngWId=-1"

Global Const constDirectLoginURL = "http://www.planet-source-code.com/vb/authentication/Wc.asp"

Global Const constOnlineHelpURL = "http://geocities.com/PlanetVBCode/PSCBrowser/Feedback/Help.htm"
Global Const constBugsURL = "http://geocities.com/PlanetVBCode/PSCBrowser/Bugs"

Global Const constFeedBackURL = "http://www.planet-source-code.com/vb/feedback/feedback.asp?lngWId=-1"
Global Const constLinkToSiteURL = "http://www.planet-source-code.com/vb/LinkToUs/default.asp?lngWId=-1"
Global Const constAwardsURL = "http://www.planet-source-code.com/vb/about/Awards.asp?lngWId=-1"
Global Const constAdvertisingURL = "http://www.planet-source-code.com/vb/advertisement/scripts/rates.asp?lngWId=-1"
Global Const constPrivacyURL = "http://www.planet-source-code.com/vb/scripts/privacy.asp?lngWId=-1"
Global Const constTermsURL = "http://www.planet-source-code.com/vb/scripts/TermsAndConditions.asp?lngWId=-1"
Global Const constAboutSiteURL = "http://www.planet-source-code.com/vb/about/AboutTheSiteAndAuthor.asp?lngWId=-1"

Global Const constLiveUpdateURL = "http://127.0.0.1/latest"

Global Const URLLogFileName = "\UrlLog.txt"
Global Const InetURLLogFileName = "\InetLog.txt"
Global Const ErrorLogFileName = "\ErrorLog.txt"

Global Const Delimeter = "ß"

Global dynamicSearchURL As String
Global dynamicNewUserURL  As String
Global dynamicForgotPwdURL  As String
Global dynamicCodeOfDayURL  As String
Global dynamicLogOnURL  As String
Global dynamicLogOnReturnURL As String
Global dynamicUpLoadURL  As String

Global dynamicSection As String
Global dynamicOptionsSection  As String
Global dynamicSettingSection As String
Global dynamicURLSection As String
Global dynamicProfileSection As String
Global dynamicFavoriteSection As String
Global dynamicFavoriteURLSection As String

'txtEmailAddress=ivicky@indiatimes.com
'&txtReturnURL=http://www.planet-source-code.com/vb/authentication/Login.asp?txtReturnURL=/vb/scripts/ShowCode.asp?txtCodeId=22595&lngWId=1
'&lngWId=
'&blnOutsideOfVBSubWeb=False
'&txtPassword=spyeyes
'&chkRememberPassword=TRUE
'&cmOk=Ok
'&strPassKey=

'SearchURL Variables
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Global LanguageID As Integer    'lngWId (1-10)
Global TextToSearch As String   'txtCriteria (A-Z Anything)
Global DropDown As String       'blnWorldDropDownUsed (TRUE , FALSE)
Global SortOrder As String      'optSort (Alphabetical , DateDescending)
Global MaxEntry  As Integer     'txtMaxNumberOfEntriesPerPage (1-99)
Global ResetVariables As String 'blnResetAllVariables (TRUE , FALSE)


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'LogOnURL Variables
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Global sEmailID As String       'txtEmailAddress (Email Address Of The User)
Global sReturnURL As String     'txtReturnURL (URL To Jump After Login)
Global LangID As Integer        'lngWId (1-10)
Global Outside As String        'blnOutsideOfVBSubWeb (TRUE , FALSE)
Global sPassword  As String     'txtPassword (PAssword Of The User)
Global SavePassword As String   'chkRememberPassword (TRUE , FALSE)
Global Const Ok As String = "Ok" 'cmOk
Global Const PassKey As String = "" 'strPassKey
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'UpLoadURL Variables
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Global CodeName As String '* 100
Global DifficultyId As Integer
Global Compat1 As Byte
Global Compat2 As Byte
Global Compat3 As Byte
Global Compat4 As Byte
Global Compat5 As Byte
Global Compat7 As Byte
Global Compat8 As Byte
Global Compat9 As Byte
Global Compat10 As Byte
Global Category As Integer
Global Purpose As String '* 4000
Global sTextCode As String '* 65300
Global ScreenShot As String
Global UploadLocalFileName As String

Global FinalURL As String
Global AppVersion As String
Global InetState As Integer
Global AdvancedSearch As Boolean        'false=basic    true=advanced
Global IsValidLogin As Boolean
Global LastURL As String
Global ScreenPosition As String

Function JumpURL(URL As String) As Boolean

On Error GoTo Debugger
    
    DoEvents
    Dim nw
    Set nw = New frmBrowser
    'nw.Load
    nw.Show
    nw.cboAddress.Text = URL
    Call nw.brwWebBrowser.Navigate(URL)
    JumpURL = True
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "basFunctions.JumpURL", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Function
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
    
End Function

Function CreateURL() As String


On Error GoTo Debugger
    
    Dim tmpURL As String
    Dim CodeDifficulity As String
    
DoEvents


    If LanguageID <= 0 Or LanguageID >= 10 Then LanguageID = 1
    
If AdvancedSearch = True Then


    tmpURL = dynamicSearchURL & "lngWId=" & LanguageID & "&txtCriteria=" & TextToSearch & "&blnWorldDropDownUsed=" & DropDown & _
                "&optSort=" & SortOrder & "&txtMaxNumberOfEntriesPerPage=" & MaxEntry & _
                "&blnResetAllVariables=" & ResetVariables & _
                "&chkCodeTypeZip=" & TRUE_FALSE(frmMain.chkCodeTypeZip.Value) & _
                "&chkCodeTypeText=" & TRUE_FALSE(frmMain.chkCodeTypeText.Value) & _
                "&chkCodeTypeArticle=" & TRUE_FALSE(frmMain.chkCodeTypeArticle.Value) & _
                "&chkCode3rdPartyReview=" & TRUE_FALSE(frmMain.chkCode3rdPartyReview.Value) & _
                "&chkThoroughSearch=" & TRUE_FALSE(frmMain.chkThoroughSearch) & _
                "&mblnIsSuperAdminAccessOn=" & TRUE_FALSE(frmMain.chkAdmAccess.Value) & _
                "&blnTopCode=" & TRUE_FALSE(frmMain.chkTopCode.Value) & _
                "&blnNewestCode=" & TRUE_FALSE(frmMain.chkNewestCode.Value)

    If Not frmMain.txtAuthorID = "" Or Not frmMain.txtAuthorName = "" Then
        tmpURL = tmpURL & "&blnAuthorSearch=TRUE&lngAuthorId=" & Val(frmMain.txtAuthorID) & "&strAuthorName=" & frmMain.txtAuthorName
    Else
        tmpURL = tmpURL & "&blnAuthorSearch=FALSE"
    End If
        tmpURL = tmpURL & "&blnEditCode=" & TRUE_FALSE(frmMain.chkEditCode.Value) & "&intFirstRecordOnPage=" & Val(frmMain.txtEntriesStart.Text) & "&intLastRecordOnPage=" & Val(frmMain.txtEntriesEnd.Text)
        
          If frmMain.chkDifficultyTypeId(1).Value And CodeDifficulity = "" Then CodeDifficulity = "1,"
                
          If frmMain.chkDifficultyTypeId(2).Value And Not CodeDifficulity = "" Then
                CodeDifficulity = CodeDifficulity & "+2,"
          Else
                CodeDifficulity = "2"
          End If
          
          If frmMain.chkDifficultyTypeId(3).Value And Not CodeDifficulity = "" Then
                CodeDifficulity = CodeDifficulity & "+3,"
          Else
                CodeDifficulity = "3"
          End If
          
          If frmMain.chkDifficultyTypeId(4).Value And Not CodeDifficulity = "" Then
                CodeDifficulity = CodeDifficulity & "+4,"
          Else
                CodeDifficulity = "4"
          End If
          
          tmpURL = tmpURL & "&chkCodeDifficulty=" & CodeDifficulity & "&cmSearch=Search"

Else

    tmpURL = dynamicSearchURL & "lngWId=" & LanguageID & "&txtCriteria=" & TextToSearch & _
                "&optSort=Alphabetical&txtMaxNumberOfEntriesPerPage=10" & _
                "&blnResetAllVariables=FALSE&B1=Quick+Search"
End If
            
            
CreateURL = Replace(tmpURL, " ", "+")

Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "basFunctions.CreateURL", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Function
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Function


Function CodeUpload()

On Error GoTo Debugger
    
Dim tmpURL As String

tmpURL = dynamicUpLoadURL & "txtCodeName=" & CodeName & _
        "&optDifficultyTypeId=" & DifficultyId & _
        "&chkCompat1=" & Compat1 & _
        "&chkCompat2=" & Compat2 & _
        "&chkCompat3=" & Compat3 & _
        "&chkCompat4=" & Compat4 & _
        "&chkCompat5=" & Compat5 & _
        "&chkCompat7=" & Compat7 & _
        "&chkCompat8=" & Compat8 & _
        "&chkCompat9=" & Compat9 & _
        "&chkCompat10=" & Compat10 & _
        "&cmbCategory=" & Category & _
        "&txtPurpose=" & Purpose & _
        "&txtCode=" & frmCode.txtCode.Text & _
        "&txtScreenShot=" & ScreenShot & "&lngScreenshotAction=1" & _
        "&txtUploadLocalFileName=" & UploadLocalFileName & _
        "&cmGo=Go"
JumpURL Trim(tmpURL)
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "basFunctions.CodeUpload", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Function
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Function

Function Delay(ByVal MilliSeconds As Long)

On Error GoTo Debugger
    
DoEvents
Sleep (MilliSeconds)
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "basFunctions.Delay", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Function
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Function




Function InetURL(URL As String, Optional HTTPCommand As String = "POST", Optional Inputs As String = "", Optional InputHeader As String = "") As String

On Error GoTo Debugger
    
    
    Dim nLength As Long
    Dim FileNo As Integer
    
    FileNo = FreeFile
    
    With frmCommon.Inet
             .Execute URL, HTTPCommand, Inputs, InputHeader
             Do
                DoEvents
             Loop Until .StillExecuting = False
             
             Debug.Print .GetHeader
             nLength = Val(.GetHeader("Content-Length"))
             If nLength = 0 Then Exit Function
             
             InetURL = .GetChunk(nLength, icByteArray)
             
             '& vbCr & .ResponseInfo
    End With
    
    Open App.Path & InetURLLogFileName For Append As #FileNo
        Print #FileNo, URL & Delimeter & HTTPCommand & Delimeter & Inputs & Delimeter & InputHeader
    Close #FileNo
   
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "basFunctions.InetURL", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Function
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
             
End Function

Function Quit() As Boolean

On Error GoTo Debugger
    

    If MsgBox("Are You Sure You Want To Quit", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
        SaveUserSettings
        Unload frmAbout
        Unload frmBrowser
        Unload frmBugs
        Unload frmCode
        Unload frmCommon
        Unload frmConfig
        Unload frmDirectLogin
        Unload frmLogin
        Unload frmMain
        Unload frmUpload
        Unload MDIBrowser
        End
    
    End If
    
    Quit = False
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "basFunctions.Quit", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Function
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
    
End Function

Function ErrLog(nNumber As Integer, Optional sMsg As String = "", Optional sSource As String = "", Optional ErrorType As VbMsgBoxStyle = vbExclamation, Optional LogToFile As Boolean = True) As Integer

On Error GoTo Debugger
    
Dim FileNo As Integer
Dim errReturn As Integer
    
    FileNo = FreeFile
    
    errReturn = MsgBox("Error: " & sMsg & " (" & nNumber & ")", ErrorType, App.ProductName & " - [ " & sSource & " ]")
    ErrLog = errReturn
    If LogToFile = False Then Exit Function
    
    Open App.Path & ErrorLogFileName For Append As #FileNo
        Print #FileNo, sMsg & Delimeter & nNumber & Delimeter & ErrorType & Delimeter & sSource
    Close #FileNo
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "basFunctions.ErrLog", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Function
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
    
End Function

Function LoadUserSettings()

On Error GoTo Debugger
    
    AppVersion = App.Major & "." & App.Minor & ".X"
    AdvancedSearch = False
    FinalURL = dynamicSearchURL & "lngWId=" & LanguageID
    
    dynamicSection = "Version\" & AppVersion
    dynamicOptionsSection = dynamicSection & "\Options"
    dynamicSettingSection = dynamicSection & "\Settings"
    dynamicURLSection = dynamicSettingSection & "\URLs"
    dynamicProfileSection = dynamicSettingSection & "\UserProfiles"
    dynamicFavoriteSection = dynamicOptionsSection & "\Favorite"
    dynamicFavoriteURLSection = dynamicOptionsSection & "\Favorite\URLs"
    
    dynamicSearchURL = GetSetting(App.ProductName, dynamicURLSection, "SearchURL", constSearchURL)
    dynamicNewUserURL = GetSetting(App.ProductName, dynamicURLSection, "NewUserURL", constNewUserURL)
    dynamicForgotPwdURL = GetSetting(App.ProductName, dynamicURLSection, "ForgotPwdURL", constForgotPwdURL)
    dynamicCodeOfDayURL = GetSetting(App.ProductName, dynamicURLSection, "CodeOfDayURL", constCodeOfDayURL)
    dynamicLogOnURL = GetSetting(App.ProductName, dynamicURLSection, "LogOnURL", constLogOnURL)
    dynamicLogOnReturnURL = GetSetting(App.ProductName, dynamicURLSection, "LogOnReturnURL", constLogOnReturnURL)
    dynamicUpLoadURL = GetSetting(App.ProductName, dynamicURLSection, "UpLoadURL", constUpLoadURL)
    
    If GetSetting(App.ProductName, dynamicSettingSection & "\Forms\frmMain", "Search Type", "Advanced") = "Advanced" Then
        frmMain.btnMoreLess.Caption = "&Basic Search"
        frmMain.Height = 6825
        AdvancedSearch = True
    Else
        frmMain.btnMoreLess.Caption = "&Advanced Search"
        frmMain.Height = 2550
        AdvancedSearch = False
    End If
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "basFunctions.LoadUserSettings", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Function
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Function

Function SaveUserSettings()

On Error GoTo Debugger
    
    
    Call SaveSetting(App.ProductName, dynamicURLSection, "SearchURL", dynamicSearchURL)
    Call SaveSetting(App.ProductName, dynamicURLSection, "NewUserURL", dynamicNewUserURL)
    Call SaveSetting(App.ProductName, dynamicURLSection, "ForgotPwdURL", dynamicForgotPwdURL)
    Call SaveSetting(App.ProductName, dynamicURLSection, "CodeOfDayURL", dynamicCodeOfDayURL)
    Call SaveSetting(App.ProductName, dynamicURLSection, "LogOnURL", dynamicLogOnURL)
    Call SaveSetting(App.ProductName, dynamicURLSection, "LogOnReturnURL", dynamicLogOnReturnURL)
    Call SaveSetting(App.ProductName, dynamicURLSection, "UpLoadURL", dynamicUpLoadURL)
    
    If frmMain.btnMoreLess.Caption = "&Basic Search" Then
        Call SaveSetting(App.ProductName, dynamicSettingSection & "\Forms\frmMain", "Search Type", "Advanced")
    Else
        Call SaveSetting(App.ProductName, dynamicSettingSection & "\Forms\frmMain", "Search Type", "Basic")
    End If
    
Debugger:
    If Not Err.Number = 0 Then
        Select Case ErrLog(Err.Number, Error$, "basFunctions.SaveUserSettings", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Function
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
    
End Function


Function ViewLog(LogType As String)

On Error GoTo Debugger

Dim tmp As String
Dim tmp1 As String
Dim FileNo As Integer
Dim ctr As Long

FileNo = FreeFile
ctr = 1
    Select Case LCase(LogType)
        
        Case "error"
            frmLogViewer.ListView1.ListItems.Clear
            
            Open App.Path & ErrorLogFileName For Input As #FileNo
            frmLogViewer.ListView1.ColumnHeaders(1).Text = "Error Description And Number"
            frmLogViewer.ListView1.ColumnHeaders(2).Text = "Error Type"
            frmLogViewer.ListView1.ColumnHeaders(3).Text = "Error Source"
            If frmLogViewer.ListView1.ColumnHeaders.Count = 4 Then frmLogViewer.ListView1.ColumnHeaders.Remove (4)
            
                
                While (Not EOF(FileNo))
                    Input #FileNo, tmp
                    tmp1 = Left(tmp, InStr(tmp, Delimeter) - 1)
                    tmp = Mid(tmp, InStr(tmp, Delimeter) + 1, Len(tmp) - InStr(tmp, Delimeter))
                    tmp1 = tmp1 & " (" & Left(tmp, InStr(tmp, Delimeter) - 1) & ")"
                    tmp = Mid(tmp, InStr(tmp, Delimeter) + 1, Len(tmp) - InStr(tmp, Delimeter))
                    
                    frmLogViewer.ListView1.ListItems.Add ctr, , tmp1, "BIG_ErrorIcon", "SMALL_ErrorIcon"
                    frmLogViewer.ListView1.ListItems(ctr).SubItems(1) = Left(tmp, InStr(tmp, Delimeter) - 1)
                    frmLogViewer.ListView1.ListItems(ctr).SubItems(2) = Mid(tmp, InStr(tmp, Delimeter) + 1, Len(tmp) - InStr(tmp, Delimeter))
                    ctr = ctr + 1
                Wend
                
            Close #FileNo
            
        Case "inet"
            frmLogViewer.ListView1.ListItems.Clear
            
            Open App.Path & InetURLLogFileName For Input As #FileNo
            
            frmLogViewer.ListView1.ColumnHeaders(1).Text = "URL"
            frmLogViewer.ListView1.ColumnHeaders(2).Text = "HTTP Command"
            frmLogViewer.ListView1.ColumnHeaders(3).Text = "Inputs"
            If frmLogViewer.ListView1.ColumnHeaders.Count = 3 Then
                frmLogViewer.ListView1.ColumnHeaders.Add(4).Text = "Input Headers"
            Else
                frmLogViewer.ListView1.ColumnHeaders(4).Text = "Input Headers"
            End If
                
                While (Not EOF(FileNo))
                    Input #FileNo, tmp
                    tmp1 = Left(tmp, InStr(tmp, Delimeter) - 1)
                    tmp = Mid(tmp, InStr(tmp, Delimeter) + 1, Len(tmp) - InStr(tmp, Delimeter))
                    frmLogViewer.ListView1.ListItems.Add ctr, , tmp1, "BIG_InetIcon", "SMALL_InetIcon"
                    
                    tmp1 = Left(tmp, InStr(tmp, Delimeter) - 1)
                    tmp = Mid(tmp, InStr(tmp, Delimeter) + 1, Len(tmp) - InStr(tmp, Delimeter))
                    
                    frmLogViewer.ListView1.ListItems(ctr).SubItems(1) = tmp1
                    frmLogViewer.ListView1.ListItems(ctr).SubItems(2) = Left(tmp, InStr(tmp, Delimeter) - 1)
                    frmLogViewer.ListView1.ListItems(ctr).SubItems(3) = Mid(tmp, InStr(tmp, Delimeter) + 1, Len(tmp) - InStr(tmp, Delimeter))
                    ctr = ctr + 1
                Wend
                
            Close #FileNo
        
        
        Case "url"
            frmLogViewer.ListView1.ListItems.Clear
            
            Open App.Path & URLLogFileName For Input As #FileNo
            
            frmLogViewer.ListView1.ColumnHeaders(1).Text = "URL"
            frmLogViewer.ListView1.ColumnHeaders(2).Text = "Date"
            frmLogViewer.ListView1.ColumnHeaders(3).Text = "Time"
            If frmLogViewer.ListView1.ColumnHeaders.Count = 4 Then frmLogViewer.ListView1.ColumnHeaders.Remove (4)
                
                While (Not EOF(FileNo))
                    Input #FileNo, tmp
                    tmp1 = Left(tmp, InStr(tmp, Delimeter) - 1)
                    tmp = Mid(tmp, InStr(tmp, Delimeter) + 1, Len(tmp) - InStr(tmp, Delimeter))
                    frmLogViewer.ListView1.ListItems.Add ctr, , tmp1, "BIG_URLIcon", "SMALL_URLIcon"
                    frmLogViewer.ListView1.ListItems(ctr).SubItems(1) = Left(tmp, InStr(tmp, Delimeter) - 1)
                    frmLogViewer.ListView1.ListItems(ctr).SubItems(2) = Mid(tmp, InStr(tmp, Delimeter) + 1, Len(tmp) - InStr(tmp, Delimeter))
                    ctr = ctr + 1
                Wend
                
            Close #FileNo
            
    End Select

Debugger:
    If Not Err.Number = 0 And Not Err.Number = 53 Then
        Select Case ErrLog(Err.Number, Error$, "basFunctions.ViewLog", vbExclamation + vbAbortRetryIgnore + vbDefaultButton1)
            Case vbAbort
                Exit Function
            Case vbRetry
                Resume
            Case vbIgnore
                Resume Next
        End Select
    End If
    
End Function

Function TRUE_FALSE(nNumber As Integer) As String
If nNumber Then
    TRUE_FALSE = "TRUE"
Else
    TRUE_FALSE = "FALSE"
End If
End Function
