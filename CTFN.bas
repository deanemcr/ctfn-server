Attribute VB_Name = "CTFNModule"
'===============================================================================
' CTFN() -- Excel User Defined Function for Pressrisk/CTFN M&A Data
'
' Returns M&A news, Active Situations, and DMA data directly in the cell.
' Calls the CTFN server REST API on Render (no local server needed).
' Works on both Windows (MSXML2) and Mac (curl via libc).
'
' SYNTAX:
'   =CTFN(parameter, [ticker], [search], [limit], [pages], [extra])
'
' EXAMPLES:
'   =CTFN("news_headline", "ROKU")          -> Latest CTFN headline for ROKU
'   =CTFN("situations", "Disney")           -> Active Situations mentioning Disney
'   =CTFN("latest_news")                    -> Most recent CTFN headlines
'   =CTFN("dma", "STIZ")                   -> DMA filings for STIZ
'   =CTFN("has_coverage", "NSC")            -> "News + Situation" / "No coverage"
'   =CTFN("news_search", , "antitrust", 5)  -> 5 news items about antitrust
'   =CTFN("dma_enforceability", "AAPL")     -> Enforceability rating
'   =CTFN("situation_count", "Disney")       -> Number of situations for Disney
'
' SETUP (for end users):
'   1. Open the CTFN.xlam file you received
'   2. In Excel: File > Options > Add-ins > Go > check "CTFN" > OK
'   3. Type =CTFN("headline","ROKU") in any cell
'   4. You will be prompted to log in (one time only)
'   5. That's it -- everything works automatically from now on.
'===============================================================================

Option Explicit

' ---------------------------------------------------------------------------
' Configuration
' ---------------------------------------------------------------------------
Private Const SERVER_URL As String = "https://ctfn.onrender.com"
'Private Const SERVER_URL As String = "http://127.0.0.1:8002"

Private Const HTTP_SEP As String = "|||"
Private Const CRED_FILENAME As String = ".ctfn_auth"

' ---------------------------------------------------------------------------
' Module-level state
' ---------------------------------------------------------------------------
Private mAuthToken As String
Private mAuthUser As String
Private mAuthPass As String
Private mAutoLoginDone As Boolean

' ---------------------------------------------------------------------------
' Mac shell support via libc (used only on Mac)
' ---------------------------------------------------------------------------
#If Mac Then
    #If VBA7 Then
        Private Declare PtrSafe Function popen Lib "libc.dylib" _
            (ByVal command As String, ByVal mode As String) As LongPtr
        Private Declare PtrSafe Function pclose Lib "libc.dylib" _
            (ByVal file As LongPtr) As Long
        Private Declare PtrSafe Function fread Lib "libc.dylib" _
            (ByVal outStr As String, ByVal size As Long, _
             ByVal items As Long, ByVal stream As LongPtr) As Long
        Private Declare PtrSafe Function feof Lib "libc.dylib" _
            (ByVal file As LongPtr) As Long
    #Else
        Private Declare Function popen Lib "libc.dylib" _
            (ByVal command As String, ByVal mode As String) As Long
        Private Declare Function pclose Lib "libc.dylib" _
            (ByVal file As Long) As Long
        Private Declare Function fread Lib "libc.dylib" _
            (ByVal outStr As String, ByVal size As Long, _
             ByVal items As Long, ByVal stream As Long) As Long
        Private Declare Function feof Lib "libc.dylib" _
            (ByVal file As Long) As Long
    #End If

Private Function ShellRun(ByVal cmd As String) As String
    #If VBA7 Then
        Dim fp As LongPtr
    #Else
        Dim fp As Long
    #End If

    fp = popen(cmd, "r")
    If fp = 0 Then
        ShellRun = ""
        Exit Function
    End If

    Dim result As String
    Dim chunk As String
    Dim bytesRead As Long

    result = ""
    Do While feof(fp) = 0
        chunk = Space$(4096)
        bytesRead = fread(chunk, 1, 4096, fp)
        If bytesRead > 0 Then
            result = result & Left$(chunk, bytesRead)
        End If
    Loop

    pclose fp
    ShellRun = result
End Function
#End If

Private Function ShellEscape(ByVal s As String) As String
    ShellEscape = Replace(s, "'", "'\''" & "'")
End Function

' ---------------------------------------------------------------------------
' Auto_Open -- runs automatically when the add-in loads.
' If saved credentials exist, logs in silently.
' If no saved credentials, prompts the user to log in.
' ---------------------------------------------------------------------------
Public Sub Auto_Open()
    If Len(mAuthToken) > 0 Then Exit Sub
    If mAutoLoginDone Then Exit Sub
    mAutoLoginDone = True

    ' Try saved credentials first (silent -- no dialogs)
    If LoadCredentials() Then
        SilentLogin
        If Len(mAuthToken) > 0 Then Exit Sub
    End If

    ' No saved credentials -- ask the user to log in now
    Dim ans As VbMsgBoxResult
    ans = MsgBox("Welcome to CTFN!" & vbCrLf & vbCrLf & _
                 "You need to log in to use =CTFN() formulas." & vbCrLf & _
                 "This is a one-time setup." & vbCrLf & vbCrLf & _
                 "Log in now?", vbYesNo + vbInformation, "CTFN")
    If ans = vbYes Then
        CTFN_LOGIN
    End If
End Sub

' ---------------------------------------------------------------------------
' Credential file path
' ---------------------------------------------------------------------------
Private Function CredFilePath() As String
#If Mac Then
    CredFilePath = Environ("HOME") & "/" & CRED_FILENAME
#Else
    CredFilePath = Environ("APPDATA") & "\" & CRED_FILENAME
#End If
End Function

' ---------------------------------------------------------------------------
' Save / Load / Delete credentials
' ---------------------------------------------------------------------------
Private Sub SaveCredentials(ByVal username As String, ByVal password As String)
    Dim fPath As String
    Dim fNum As Integer
    fPath = CredFilePath()

    On Error GoTo SaveErr
    fNum = FreeFile
    Open fPath For Output As #fNum
    Print #fNum, ObfuscateString(username & vbLf & password)
    Close #fNum
    Exit Sub
SaveErr:
    If fNum > 0 Then Close #fNum
End Sub

Private Function LoadCredentials() As Boolean
    Dim fPath As String
    Dim fNum As Integer
    Dim payload As String
    Dim decoded As String
    Dim nlPos As Long

    fPath = CredFilePath()
    LoadCredentials = False

    On Error Resume Next
    Dim attr As Long
    attr = GetAttr(fPath)
    If Err.Number <> 0 Then
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    On Error GoTo LoadErr
    fNum = FreeFile
    Open fPath For Input As #fNum
    Line Input #fNum, payload
    Close #fNum

    decoded = DeobfuscateString(payload)
    nlPos = InStr(1, decoded, vbLf)
    If nlPos > 0 Then
        mAuthUser = Left(decoded, nlPos - 1)
        mAuthPass = Mid(decoded, nlPos + 1)
        LoadCredentials = True
    End If
    Exit Function
LoadErr:
    If fNum > 0 Then Close #fNum
End Function

Private Sub DeleteCredentials()
    On Error Resume Next
    Kill CredFilePath()
    On Error GoTo 0
    mAuthPass = ""
End Sub

' ---------------------------------------------------------------------------
' XOR obfuscation (prevents casual reading of saved credentials)
' ---------------------------------------------------------------------------
Private Function ObfuscateString(ByVal s As String) As String
    Dim key As String: key = "CTFNExcelKey2026"
    Dim i As Long
    Dim result As String: result = ""
    For i = 1 To Len(s)
        result = result & Right("0" & Hex(Asc(Mid(s, i, 1)) Xor _
                 Asc(Mid(key, ((i - 1) Mod Len(key)) + 1, 1))), 2)
    Next i
    ObfuscateString = result
End Function

Private Function DeobfuscateString(ByVal s As String) As String
    Dim key As String: key = "CTFNExcelKey2026"
    Dim i As Long
    Dim result As String: result = ""
    For i = 1 To Len(s) Step 2
        result = result & Chr(CLng("&H" & Mid(s, i, 2)) Xor _
                 Asc(Mid(key, (((i - 1) / 2) Mod Len(key)) + 1, 1)))
    Next i
    DeobfuscateString = result
End Function

' ---------------------------------------------------------------------------
' ScheduleLoginPrompt -- attempts to schedule CTFN_LOGIN via OnTime
' Called from UDF context where direct dialogs are forbidden.
' ---------------------------------------------------------------------------
Private Sub ScheduleLoginPrompt()
    On Error Resume Next
    Application.OnTime Now + TimeSerial(0, 0, 1), "CTFN_LOGIN"
    On Error GoTo 0
End Sub

' ---------------------------------------------------------------------------
' Silent login (no dialogs -- uses stored credentials)
' ---------------------------------------------------------------------------
Private Function SilentLogin() As Boolean
    Dim jsonBody As String
    Dim rawResponse As String
    Dim pipePos As Long
    Dim httpStatus As String
    Dim responseBody As String
    Dim token As String

    SilentLogin = False
    If Len(mAuthUser) = 0 Or Len(mAuthPass) = 0 Then Exit Function

    jsonBody = "{""username"":""" & Replace(mAuthUser, """", "") & _
               """,""password"":""" & Replace(mAuthPass, """", "") & """}"

    rawResponse = HttpPost_CTFN(SERVER_URL & "/login", jsonBody)

    pipePos = InStr(1, rawResponse, "|")
    If pipePos = 0 Then Exit Function

    httpStatus = Left(rawResponse, pipePos - 1)
    responseBody = Mid(rawResponse, pipePos + 1)

    If httpStatus = "200" Then
        token = JsonGetString(responseBody, "token")
        If Len(token) > 0 Then
            mAuthToken = token
            SilentLogin = True
        End If
    Else
        DeleteCredentials
        mAuthUser = ""
        mAuthPass = ""
    End If
End Function

' ---------------------------------------------------------------------------
' EnsureAuth -- called before every API request
' ---------------------------------------------------------------------------
Private Function EnsureAuth() As Boolean
    If Len(mAuthToken) > 0 Then
        EnsureAuth = True
        Exit Function
    End If

    ' Try loading saved credentials if we haven't
    If Len(mAuthPass) = 0 And Not mAutoLoginDone Then
        mAutoLoginDone = True
        If LoadCredentials() Then
            If SilentLogin() Then
                EnsureAuth = True
                Exit Function
            End If
        End If
    End If

    ' Token may have expired -- try re-login with saved credentials
    If Len(mAuthPass) > 0 And Len(mAuthToken) = 0 Then
        If SilentLogin() Then
            EnsureAuth = True
            Exit Function
        End If
    End If

    EnsureAuth = (Len(mAuthToken) > 0)
End Function

' ---------------------------------------------------------------------------
' Main UDF
' ---------------------------------------------------------------------------
Public Function CTFN(ByVal param As String, _
                     Optional ByVal ticker As String = "", _
                     Optional ByVal search As String = "", _
                     Optional ByVal limit As Long = 10, _
                     Optional ByVal pages As String = "1", _
                     Optional ByVal extra As String = "") As Variant

    On Error GoTo ErrHandler

    Dim url As String
    Dim response As String
    Dim result As Variant

    ' Auto-login if needed
    EnsureAuth

    ' No credentials at all -- tell the user what to do
    If Len(mAuthToken) = 0 And Len(mAuthPass) = 0 Then
        CTFN = "Not logged in. Go to Developer > Macros > CTFN_LOGIN"
        Exit Function
    End If

    ' Build the API URL
    url = SERVER_URL & "/api/ctfn?" & _
          "param=" & EncodeURL_CTFN(Trim(param)) & _
          "&ticker=" & EncodeURL_CTFN(Trim(ticker)) & _
          "&search=" & EncodeURL_CTFN(Trim(search)) & _
          "&limit=" & CStr(limit) & _
          "&pages=" & EncodeURL_CTFN(Trim(pages))

    If Len(extra) > 0 Then
        url = url & "&extra=" & EncodeURL_CTFN(Trim(extra))
    End If

    ' Make the HTTP GET request
    response = HttpGet_CTFN(url)

    ' If 401, try re-login with saved credentials and retry once
    If response = "#AUTH" Then
        If Len(mAuthPass) > 0 Then
            mAuthToken = ""
            If SilentLogin() Then
                response = HttpGet_CTFN(url)
            End If
        End If
    End If

    ' Still auth error -- schedule login prompt
    If response = "#AUTH" Then
        ScheduleLoginPrompt
        CTFN = "Logging in..."
        Exit Function
    End If

    If Left(response, 5) = "#ERR:" Then
        CTFN = response
        Exit Function
    End If

    result = ParseCTFNResponse(response)
    CTFN = result
    Exit Function

ErrHandler:
    CTFN = "#ERR: " & Err.Description
End Function

' ---------------------------------------------------------------------------
' HTTP GET (cross-platform)
' ---------------------------------------------------------------------------
Private Function HttpGet_CTFN(ByVal url As String) As String
#If Mac Then
    Dim cmd As String
    cmd = "curl -s -w '" & HTTP_SEP & "%{http_code}' " & _
          "'" & url & "' " & _
          "-H 'Accept: application/json'"
    If Len(mAuthToken) > 0 Then
        cmd = cmd & " -H 'Authorization: Bearer " & mAuthToken & "'"
    End If

    Dim raw As String: raw = ShellRun(cmd)
    Dim sepPos As Long: sepPos = InStrRev(raw, HTTP_SEP)
    If sepPos > 0 Then
        Dim body As String: body = Left$(raw, sepPos - 1)
        Dim statusCode As String: statusCode = Trim$(Mid$(raw, sepPos + Len(HTTP_SEP)))
        Select Case statusCode
            Case "200": HttpGet_CTFN = body
            Case "401": HttpGet_CTFN = "#AUTH"
            Case Else:  HttpGet_CTFN = "#ERR: HTTP " & statusCode
        End Select
    Else
        HttpGet_CTFN = "#ERR: Server not reachable"
    End If
#Else
    Dim http As Object
    On Error GoTo HttpError
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", url, False
    http.setRequestHeader "Accept", "application/json"
    If Len(mAuthToken) > 0 Then
        http.setRequestHeader "Authorization", "Bearer " & mAuthToken
    End If
    http.send
    Select Case http.Status
        Case 200: HttpGet_CTFN = http.responseText
        Case 401: HttpGet_CTFN = "#AUTH"
        Case Else: HttpGet_CTFN = "#ERR: HTTP " & http.Status & " - " & http.statusText
    End Select
    Set http = Nothing
    Exit Function
HttpError:
    HttpGet_CTFN = "#ERR: Server not reachable"
    If Not http Is Nothing Then Set http = Nothing
#End If
End Function

' ---------------------------------------------------------------------------
' HTTP POST (cross-platform) -- returns "statusCode|responseBody"
' ---------------------------------------------------------------------------
Private Function HttpPost_CTFN(ByVal url As String, ByVal jsonBody As String) As String
#If Mac Then
    Dim cmd As String
    cmd = "curl -s -w '" & HTTP_SEP & "%{http_code}' " & _
          "-X POST '" & url & "' " & _
          "-H 'Content-Type: application/json' " & _
          "-H 'Accept: application/json'"
    If Len(mAuthToken) > 0 Then
        cmd = cmd & " -H 'Authorization: Bearer " & mAuthToken & "'"
    End If
    cmd = cmd & " -d '" & ShellEscape(jsonBody) & "'"

    Dim raw As String: raw = ShellRun(cmd)
    Dim sepPos As Long: sepPos = InStrRev(raw, HTTP_SEP)
    If sepPos > 0 Then
        Dim body As String: body = Left$(raw, sepPos - 1)
        Dim statusCode As String: statusCode = Trim$(Mid$(raw, sepPos + Len(HTTP_SEP)))
        HttpPost_CTFN = statusCode & "|" & body
    Else
        HttpPost_CTFN = "0|#ERR: Server not reachable"
    End If
#Else
    Dim http As Object
    On Error GoTo PostError
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Accept", "application/json"
    If Len(mAuthToken) > 0 Then
        http.setRequestHeader "Authorization", "Bearer " & mAuthToken
    End If
    http.send jsonBody
    HttpPost_CTFN = http.Status & "|" & http.responseText
    Set http = Nothing
    Exit Function
PostError:
    HttpPost_CTFN = "0|#ERR: Server not reachable"
    If Not http Is Nothing Then Set http = Nothing
#End If
End Function

' ---------------------------------------------------------------------------
' JSON response parser
' ---------------------------------------------------------------------------
Private Function ParseCTFNResponse(ByVal jsonStr As String) As Variant
    Dim valStart As Long
    Dim valStr As String

    valStart = InStr(1, jsonStr, """value"":")
    If valStart = 0 Then
        ParseCTFNResponse = "#ERR: Invalid response"
        Exit Function
    End If
    valStart = valStart + 8
    Do While valStart <= Len(jsonStr) And Mid(jsonStr, valStart, 1) = " "
        valStart = valStart + 1
    Loop

    valStr = Mid(jsonStr, valStart)
    If Right(Trim(valStr), 1) = "}" Then
        valStr = Left(Trim(valStr), Len(Trim(valStr)) - 1)
    End If
    valStr = Trim(valStr)

    If valStr = "null" Then
        ParseCTFNResponse = "N/A"
    ElseIf Left(valStr, 1) = """" Then
        valStr = Mid(valStr, 2)
        If Right(valStr, 1) = """" Then valStr = Left(valStr, Len(valStr) - 1)
        ParseCTFNResponse = valStr
    ElseIf IsNumeric(valStr) Then
        ParseCTFNResponse = CDbl(valStr)
    Else
        ParseCTFNResponse = valStr
    End If
End Function

' ---------------------------------------------------------------------------
' URL encoding
' ---------------------------------------------------------------------------
Private Function EncodeURL_CTFN(ByVal str As String) As String
    Dim i As Long, ch As String, result As String
    result = ""
    For i = 1 To Len(str)
        ch = Mid(str, i, 1)
        Select Case ch
            Case "A" To "Z", "a" To "z", "0" To "9", "-", "_", ".", "~"
                result = result & ch
            Case " "
                result = result & "%20"
            Case Else
                result = result & "%" & Right("0" & Hex(Asc(ch)), 2)
        End Select
    Next i
    EncodeURL_CTFN = result
End Function

' ---------------------------------------------------------------------------
' JSON string extractor
' ---------------------------------------------------------------------------
Private Function JsonGetString(ByVal json As String, ByVal key As String) As String
    Dim search As String, pos As Long, startPos As Long, endPos As Long
    search = """" & key & """:"
    pos = InStr(1, json, search)
    If pos = 0 Then Exit Function

    startPos = pos + Len(search)
    Do While startPos <= Len(json) And Mid(json, startPos, 1) = " "
        startPos = startPos + 1
    Loop

    If Mid(json, startPos, 1) = """" Then
        startPos = startPos + 1
        endPos = InStr(startPos, json, """")
        If endPos > 0 Then JsonGetString = Mid(json, startPos, endPos - startPos)
    Else
        endPos = startPos
        Do While endPos <= Len(json)
            Dim c As String: c = Mid(json, endPos, 1)
            If c = "," Or c = "}" Or c = " " Then Exit Do
            endPos = endPos + 1
        Loop
        JsonGetString = Mid(json, startPos, endPos - startPos)
    End If
End Function

' ===========================================================================
' PUBLIC MACROS
' ===========================================================================

' ---------------------------------------------------------------------------
' CTFN_LOGIN -- prompts for credentials and saves them permanently
' ---------------------------------------------------------------------------
Public Sub CTFN_LOGIN()
    Dim username As String
    Dim password As String
    Dim jsonBody As String
    Dim rawResponse As String
    Dim httpStatus As String
    Dim responseBody As String
    Dim pipePos As Long
    Dim token As String

    ' Already logged in?
    If Len(mAuthToken) > 0 And Len(mAuthPass) > 0 Then
        Dim ans As VbMsgBoxResult
        ans = MsgBox("Already logged in as " & mAuthUser & "." & vbCrLf & vbCrLf & _
                     "Log in as a different user?", vbYesNo + vbQuestion, "CTFN")
        If ans = vbNo Then Exit Sub
        mAuthToken = ""
        mAuthUser = ""
        mAuthPass = ""
    End If

    username = InputBox("Enter your CTFN username (email):", "CTFN Login")
    If Len(Trim(username)) = 0 Then
        MsgBox "Login cancelled.", vbInformation, "CTFN"
        Exit Sub
    End If

    password = InputBox("Enter your password:", "CTFN Login")
    If Len(Trim(password)) = 0 Then
        MsgBox "Login cancelled.", vbInformation, "CTFN"
        Exit Sub
    End If

    jsonBody = "{""username"":""" & Replace(username, """", "") & _
               """,""password"":""" & Replace(password, """", "") & """}"

    rawResponse = HttpPost_CTFN(SERVER_URL & "/login", jsonBody)

    pipePos = InStr(1, rawResponse, "|")
    If pipePos = 0 Then
        MsgBox "Could not reach the CTFN server." & vbCrLf & vbCrLf & _
               "Check your internet connection and try again.", _
               vbExclamation, "CTFN Login"
        Exit Sub
    End If

    httpStatus = Left(rawResponse, pipePos - 1)
    responseBody = Mid(rawResponse, pipePos + 1)

    If httpStatus = "200" Then
        token = JsonGetString(responseBody, "token")
        If Len(token) > 0 Then
            mAuthToken = token
            mAuthUser = Trim(username)
            mAuthPass = Trim(password)
            SaveCredentials mAuthUser, mAuthPass

            MsgBox "Logged in as " & mAuthUser & "." & vbCrLf & vbCrLf & _
                   "You won't need to log in again.", _
                   vbInformation, "CTFN"
            Application.CalculateFullRebuild
        Else
            MsgBox "Login succeeded but no token received." & vbCrLf & responseBody, _
                   vbExclamation, "CTFN Login"
        End If
    ElseIf httpStatus = "401" Then
        MsgBox "Invalid username or password.", vbExclamation, "CTFN Login"
    Else
        MsgBox "Login failed (HTTP " & httpStatus & ")." & vbCrLf & vbCrLf & responseBody, _
               vbExclamation, "CTFN Login"
    End If
End Sub

' ---------------------------------------------------------------------------
' CTFN_LOGOUT -- clears session AND saved credentials
' ---------------------------------------------------------------------------
Public Sub CTFN_LOGOUT()
    If Len(mAuthToken) > 0 Then
        HttpPost_CTFN SERVER_URL & "/logout", "{}"
    End If
    mAuthToken = ""
    mAuthUser = ""
    mAuthPass = ""
    mAutoLoginDone = False
    DeleteCredentials
    MsgBox "Logged out and saved credentials cleared." & vbCrLf & vbCrLf & _
           "Run CTFN_LOGIN to log in again.", vbInformation, "CTFN"
End Sub

' ---------------------------------------------------------------------------
' CTFN_REFRESH -- recalculate all CTFN formulas
' ---------------------------------------------------------------------------
Public Sub CTFN_REFRESH()
    If Len(mAuthToken) = 0 Then EnsureAuth
    If Len(mAuthToken) = 0 Then
        CTFN_LOGIN
        If Len(mAuthToken) = 0 Then Exit Sub
    End If
    Application.CalculateFullRebuild
    MsgBox "All CTFN formulas refreshed.", vbInformation, "CTFN"
End Sub

' ---------------------------------------------------------------------------
' CTFN_STATUS -- check server and auth state
' ---------------------------------------------------------------------------
Public Sub CTFN_STATUS()
    Dim response As String, authInfo As String

    If Len(mAuthToken) = 0 Then EnsureAuth
    response = HttpGet_CTFN(SERVER_URL & "/health")

    If Len(mAuthToken) > 0 Then
        authInfo = "Logged in as: " & mAuthUser & " (credentials saved)"
    ElseIf Len(mAuthPass) > 0 Then
        authInfo = "Saved credentials found for: " & mAuthUser
    Else
        authInfo = "Not logged in"
    End If

    If Left(response, 5) = "#ERR:" Or response = "#AUTH" Then
        MsgBox "Server: NOT reachable" & vbCrLf & _
               "URL: " & SERVER_URL & vbCrLf & _
               authInfo, vbExclamation, "CTFN Status"
    Else
        MsgBox "Server: Running" & vbCrLf & _
               "URL: " & SERVER_URL & vbCrLf & _
               authInfo, vbInformation, "CTFN Status"
    End If
End Sub
