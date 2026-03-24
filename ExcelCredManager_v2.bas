Attribute VB_Name = "ExcelCredManager"
' =======================================================
' ExcelCredManager_v2.bas  (v4 — Pure VBA, quality fixes)
'
' Fixes over v3:
'   1. ScanAuth error handler restores Application state on failure
'   2. Stale ws reference reset before InitReport check
'   3. File property access extracted to ProcessTokenFile with own error handler
'   4. GoTo out of For Each replaced with Exit For
'   5. cmdkey parser handles both "Target:" and "Target Name:" variants
'   6. ReadFileText uses ADODB.Stream for UTF-8, falls back to FSO
'   7. ScanWindowsIdentity error handling on WScript.Network creation
'   8. ReadRegValue caches WScript.Shell via Static
'   9. EnumKey result checked with IsArray, not just IsNull
'  10. Report column width governed by MAX_COLS constant
'
' Import: Alt+F11 > File > Import File
' Save workbook as .xlsm
' =======================================================

Option Explicit

Private Const SHEET_NAME As String = "AuthReport"
Private Const STALE_DAYS As Long = 1
Private Const MAX_COLS   As Long = 8

' Registry
Private Const HKCU                  As Long = &H80000001
Private Const OFFICE_IDENTITY_KEY   As String = "Software\Microsoft\Office\16.0\Common\Identity"
Private Const OFFICE_IDENTITIES_KEY As String = "Software\Microsoft\Office\16.0\Common\Identity\Identities"

' Report state (module-level so section subs can share)
Private ws   As Worksheet
Private row_ As Long

' ── Helpers ──────────────────────────────────────────────────────

Private Sub WriteSection(title As String)
    row_ = row_ + 1
    ws.Range(ws.Cells(row_, 1), ws.Cells(row_, MAX_COLS)).Interior.Color = RGB(30, 30, 30)
    With ws.Cells(row_, 1)
        .Value = title
        .Font.Bold = True
        .Font.Size = 12
        .Font.Color = RGB(255, 255, 255)
    End With
    row_ = row_ + 1
End Sub

Private Sub WriteHeader(ParamArray cols() As Variant)
    Dim c As Long
    For c = 0 To UBound(cols)
        With ws.Cells(row_, c + 1)
            .Value = cols(c)
            .Font.Bold = True
            .Interior.Color = RGB(220, 220, 220)
        End With
    Next c
    row_ = row_ + 1
End Sub

Private Sub WriteRow(ParamArray cols() As Variant)
    Dim c As Long
    For c = 0 To UBound(cols)
        ws.Cells(row_, c + 1).Value = cols(c)
    Next c
    row_ = row_ + 1
End Sub

' [FIX #8] Static caches WScript.Shell across calls
Private Function ReadRegValue(key As String, valueName As String) As String
    Static wsh As Object
    On Error Resume Next
    If wsh Is Nothing Then Set wsh = CreateObject("WScript.Shell")
    ReadRegValue = wsh.RegRead("HKCU\" & key & "\" & valueName)
    If Err.Number <> 0 Then ReadRegValue = ""
    Err.Clear
    On Error GoTo 0
End Function

' [FIX #6] ADODB.Stream for UTF-8, fallback to FSO
Private Function ReadFileText(filePath As String) As String
    ReadFileText = ""

    ' Try ADODB.Stream (proper UTF-8 support)
    On Error Resume Next
    Dim stm As Object: Set stm = CreateObject("ADODB.Stream")
    If Err.Number = 0 Then
        stm.Type = 2         ' adTypeText
        stm.Charset = "utf-8"
        stm.Open
        stm.LoadFromFile filePath
        If Err.Number = 0 Then
            ReadFileText = stm.ReadText(-1)  ' adReadAll
            stm.Close
            Err.Clear
            On Error GoTo 0
            Exit Function
        End If
        Err.Clear
        On Error Resume Next
        stm.Close
        Err.Clear
    End If
    Err.Clear

    ' Fallback: FSO (ASCII — strip UTF-8 BOM if present)
    Dim fso As Object, ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(filePath, 1, False)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    Dim raw As String: raw = ts.ReadAll
    ts.Close
    On Error GoTo 0

    ' Strip UTF-8 BOM (EF BB BF appears as chr 239,187,191 in ASCII read)
    If Len(raw) >= 3 Then
        If Left(raw, 3) = Chr(239) & Chr(187) & Chr(191) Then
            raw = Mid(raw, 4)
        End If
    End If
    ReadFileText = raw
End Function

Private Function ExtractJsonField(json As String, fieldName As String) As String
    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.Pattern = """" & fieldName & """\s*:\s*""?([^"",}\s]+)""?"
    re.IgnoreCase = True
    Dim matches As Object: Set matches = re.Execute(json)
    If matches.Count > 0 Then
        ExtractJsonField = matches(0).SubMatches(0)
    Else
        ExtractJsonField = ""
    End If
End Function

Private Function UnixToDate(unixStr As String) As Date
    If Len(unixStr) = 0 Then UnixToDate = 0: Exit Function
    On Error Resume Next
    Dim epoch As Double: epoch = CDbl(unixStr)
    If Err.Number <> 0 Then UnixToDate = 0: Err.Clear: On Error GoTo 0: Exit Function
    On Error GoTo 0
    ' Handle millisecond timestamps
    If epoch > 9999999999# Then epoch = epoch / 1000
    UnixToDate = DateAdd("s", epoch, #1/1/1970#)
End Function

Private Function TokenStatus(lastWrite As Date, Optional expiry As Date = 0) As String
    Dim now_ As Date: now_ = Now
    If expiry > 0 Then
        If expiry < now_ Then
            TokenStatus = "Expired"
        ElseIf expiry < DateAdd("h", 1, now_) Then
            TokenStatus = "Expiring Soon"
        Else
            TokenStatus = "Valid"
        End If
    Else
        Dim ageDays As Double: ageDays = now_ - lastWrite
        If ageDays > STALE_DAYS Then
            TokenStatus = "Stale (" & Round(ageDays, 1) & "d)"
        Else
            TokenStatus = "Recent"
        End If
    End If
End Function

Private Sub ColorStatus(r As Long, c As Long)
    Dim cell As Range: Set cell = ws.Cells(r, c)
    Dim status As String: status = cell.Value
    Select Case True
        Case InStr(status, "Expired") > 0
            cell.Interior.Color = RGB(255, 200, 200)
            cell.Font.Color = RGB(180, 0, 0)
        Case InStr(status, "Stale") > 0 Or InStr(status, "Expiring") > 0
            cell.Interior.Color = RGB(255, 240, 180)
            cell.Font.Color = RGB(160, 100, 0)
        Case InStr(status, "Valid") > 0 Or InStr(status, "Recent") > 0
            cell.Interior.Color = RGB(200, 240, 200)
            cell.Font.Color = RGB(0, 120, 0)
    End Select
End Sub

' ── Section: Windows Identity ──────────────────────────────────

' [FIX #7] Error handling on WScript.Network creation
Private Sub ScanWindowsIdentity()
    WriteSection "Windows Identity"
    WriteHeader "Property", "Value"

    WriteRow "Username", Environ("USERNAME")
    WriteRow "Domain", Environ("USERDOMAIN")
    WriteRow "DNS Domain", Environ("USERDNSDOMAIN")
    WriteRow "Computer", Environ("COMPUTERNAME")
    WriteRow "User Profile", Environ("USERPROFILE")

    Dim net As Object
    On Error Resume Next
    Set net = CreateObject("WScript.Network")
    If Err.Number = 0 Then
        WriteRow "Network Username", net.UserName
        WriteRow "Network Domain", net.UserDomain
    Else
        WriteRow "Network Info", "(WScript.Network unavailable)"
    End If
    Err.Clear
    On Error GoTo 0

    WriteRow "Office Display Name", Application.UserName
End Sub

' ── Section: Office 365 Identity ──────────────────────────────

Private Sub ScanOfficeIdentity()
    WriteSection "Office 365 Identity"

    ' Top-level identity registry values
    WriteHeader "Property", "Value"
    Dim topVals As Variant, v As Long
    topVals = Array("ADAccountType", "FederationProvider", "DefaultADUser", "ConnectedAccountType")
    For v = 0 To UBound(topVals)
        Dim regVal As String: regVal = ReadRegValue(OFFICE_IDENTITY_KEY, CStr(topVals(v)))
        If Len(regVal) > 0 Then WriteRow topVals(v), regVal
    Next v

    ' Enumerate cached identities via WMI StdRegProv
    Dim oReg As Object
    On Error Resume Next
    Set oReg = GetObject("winmgmts:\\.\root\default:StdRegProv")
    If Err.Number <> 0 Then
        Err.Clear: On Error GoTo 0
        WriteRow "(WMI StdRegProv unavailable — cannot enumerate identities)"
        Exit Sub
    End If
    On Error GoTo 0

    row_ = row_ + 1
    WriteHeader "Email", "Display Name", "Provider ID", "Sign-In State"

    ' [FIX #9] Check IsArray, not just IsNull
    Dim subKeys As Variant, lRet As Long
    lRet = oReg.EnumKey(HKCU, OFFICE_IDENTITIES_KEY, subKeys)

    If lRet <> 0 Or IsEmpty(subKeys) Or IsNull(subKeys) Then
        WriteRow "(no cached identities found)", "", "", ""
        Exit Sub
    End If
    If Not IsArray(subKeys) Then
        WriteRow "(no cached identities found)", "", "", ""
        Exit Sub
    End If

    Dim i As Long
    For i = 0 To UBound(subKeys)
        Dim idKey As String:    idKey = OFFICE_IDENTITIES_KEY & "\" & subKeys(i)
        Dim email As String:    email = ReadRegValue(idKey, "EmailAddress")
        Dim friendly As String: friendly = ReadRegValue(idKey, "FriendlyName")
        Dim provider As String: provider = ReadRegValue(idKey, "ProviderId")
        Dim signIn As String:   signIn = ReadRegValue(idKey, "SignInState")

        If Len(email) > 0 Or Len(friendly) > 0 Then
            WriteRow email, friendly, provider, signIn
        End If
    Next i
End Sub

' ── Section: Token Stores ──────────────────────────────────────

Private Sub ScanTokenStores()
    WriteSection "Token Stores"
    WriteHeader "Store", "File", "Last Modified", "Age (days)", "Expiry", "Status"

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim localApp As String: localApp = Environ("LOCALAPPDATA")

    If Len(localApp) = 0 Then
        WriteRow "(LOCALAPPDATA not set)", "", "", "", "", ""
        Exit Sub
    End If

    Dim found As Long: found = 0

    Dim storeNames(3) As String, storePaths(3) As String
    storeNames(0) = "OneAuth":       storePaths(0) = localApp & "\Microsoft\OneAuth\accounts"
    storeNames(1) = "TokenBroker":   storePaths(1) = localApp & "\Microsoft\TokenBroker\Accounts"
    storeNames(2) = "PowerQuery":    storePaths(2) = localApp & "\Microsoft\Power Query"
    storeNames(3) = "IdentitySvc":   storePaths(3) = localApp & "\Microsoft\.IdentityService"

    Dim s As Long
    For s = 0 To 3
        If fso.FolderExists(storePaths(s)) Then
            ScanFolderForTokens fso, fso.GetFolder(storePaths(s)), storeNames(s), found, 0
        Else
            WriteRow storeNames(s), "(not found)", storePaths(s), "", "", "N/A"
        End If
    Next s

    If found = 0 Then
        WriteRow "(no token files found in any store)", "", "", "", "", ""
    End If
End Sub

' [FIX #3, #4] Restructured — file access extracted to ProcessTokenFile;
'              no GoTo; clean Exit For on enumeration errors
Private Sub ScanFolderForTokens(fso As Object, folder As Object, storeName As String, _
                                 ByRef found As Long, depth As Long)
    If depth > 2 Then Exit Sub

    ' Enumerate files with error protection
    Dim f As Object
    On Error Resume Next
    For Each f In folder.Files
        If Err.Number <> 0 Then Err.Clear: Exit For
        On Error GoTo 0
        ProcessTokenFile fso, f, storeName, found
        On Error Resume Next
    Next f
    On Error GoTo 0

    ' Recurse subfolders
    Dim sf As Object
    On Error Resume Next
    For Each sf In folder.SubFolders
        If Err.Number <> 0 Then Err.Clear: Exit For
        On Error GoTo 0
        ScanFolderForTokens fso, sf, storeName, found, depth + 1
        On Error Resume Next
    Next sf
    On Error GoTo 0
End Sub

' [FIX #3] Isolated error handler per file — one bad file cannot crash the scan
Private Sub ProcessTokenFile(fso As Object, f As Object, storeName As String, ByRef found As Long)
    On Error GoTo SkipFile

    Dim ext As String: ext = LCase(fso.GetExtensionName(f.Name))
    If ext <> "json" And ext <> "cache" And ext <> "" Then Exit Sub

    Dim age As Double:       age = Now - f.DateLastModified
    Dim expiryStr As String: expiryStr = ""
    Dim expiryDate As Date:  expiryDate = 0

    If ext = "json" Then
        Dim content As String: content = ReadFileText(f.Path)
        If Len(content) > 0 Then
            Dim fields As Variant
            fields = Array("extended_expires_on", "expires_on", "expiry_time", "exp", "expirationTime")
            Dim fi As Long
            For fi = 0 To UBound(fields)
                Dim val As String: val = ExtractJsonField(content, CStr(fields(fi)))
                If Len(val) > 0 Then
                    expiryDate = UnixToDate(val)
                    If expiryDate > 0 Then
                        expiryStr = Format(expiryDate, "yyyy-mm-dd hh:nn:ss")
                        Exit For
                    End If
                End If
            Next fi
        End If
    End If

    Dim status As String: status = TokenStatus(f.DateLastModified, expiryDate)
    Dim startRow As Long: startRow = row_

    WriteRow storeName, f.Name, _
             Format(f.DateLastModified, "yyyy-mm-dd hh:nn:ss"), _
             Round(age, 1), expiryStr, status
    ColorStatus startRow, 6
    found = found + 1
    Exit Sub

SkipFile:
    ' File inaccessible (locked, deleted, permissions) — skip silently
End Sub

' ── Section: Workbook Connections ─────────────────────────────

Private Sub ScanConnections()
    WriteSection "Workbook Data Connections"

    If ThisWorkbook.Connections.Count = 0 Then
        WriteRow "(no data connections in this workbook)"
        Exit Sub
    End If

    WriteHeader "Name", "Type", "Description", "Connection String", "Auth Hint"

    Dim conn As WorkbookConnection
    For Each conn In ThisWorkbook.Connections
        Dim connType As String
        Select Case conn.Type
            Case 1:  connType = "OLE DB"
            Case 2:  connType = "ODBC"
            Case 3:  connType = "XML Map"
            Case 4:  connType = "Text"
            Case 5:  connType = "Web"
            Case 6:  connType = "Data Feed"
            Case 7:  connType = "Model"
            Case Else: connType = "Type " & conn.Type
        End Select

        Dim connStr As String: connStr = ""
        On Error Resume Next
        connStr = conn.OLEDBConnection.Connection
        If Len(connStr) = 0 Then connStr = conn.ODBCConnection.Connection
        Err.Clear
        On Error GoTo 0

        Dim authHint As String: authHint = ParseAuthHint(connStr)

        ' Truncate long connection strings
        If Len(connStr) > 200 Then connStr = Left(connStr, 200) & "..."

        WriteRow conn.Name, connType, conn.Description, connStr, authHint
    Next conn
End Sub

Private Function ParseAuthHint(connStr As String) As String
    Dim lc As String: lc = LCase(connStr)
    Dim hint As String: hint = ""

    If InStr(lc, "integrated security=sspi") > 0 Or InStr(lc, "trusted_connection=yes") > 0 Then
        hint = "Windows Auth (SSPI)"
    ElseIf InStr(lc, "authentication=activedirectoryinteractive") > 0 Then
        hint = "AAD Interactive"
    ElseIf InStr(lc, "authentication=activedirectoryintegrated") > 0 Then
        hint = "AAD Integrated"
    ElseIf InStr(lc, "authentication=activedirectorypassword") > 0 Then
        hint = "AAD Password"
    ElseIf InStr(lc, "authentication=activedirectoryserviceprincipal") > 0 Then
        hint = "AAD Service Principal"
    ElseIf InStr(lc, "user id=") > 0 Or InStr(lc, "uid=") > 0 Then
        hint = "SQL Auth (user/pass)"
    ElseIf InStr(lc, "oauth") > 0 Or InStr(lc, "token") > 0 Then
        hint = "OAuth/Token"
    End If

    If InStr(lc, "microsoft.mashup") > 0 Then
        If Len(hint) > 0 Then hint = hint & " + "
        hint = hint & "Power Query"
    ElseIf InStr(lc, "sqloledb") > 0 Or InStr(lc, "sqlncli") > 0 Or InStr(lc, "msoledbsql") > 0 Then
        If Len(hint) = 0 Then hint = "SQL Server"
    End If

    ParseAuthHint = hint
End Function

' ── Section: Credential Manager ───────────────────────────────

Private Sub ScanCredentialManager()
    WriteSection "Credential Manager (Office/SharePoint/SQL)"

    Dim tempFile As String
    tempFile = Environ("TEMP") & "\xl_cmdkey_out.txt"

    On Error Resume Next
    Dim exitCode As Long
    exitCode = CreateObject("WScript.Shell").Run( _
        "cmd /c cmdkey /list > """ & tempFile & """ 2>&1", 0, True)
    If Err.Number <> 0 Then
        Err.Clear: On Error GoTo 0
        WriteRow "(cmd.exe may be restricted — cmdkey unavailable)"
        Exit Sub
    End If
    On Error GoTo 0

    Dim content As String: content = ReadFileText(tempFile)

    ' Clean up
    On Error Resume Next: Kill tempFile: On Error GoTo 0

    If Len(content) = 0 Then
        WriteRow "(no output from cmdkey /list)"
        Exit Sub
    End If

    WriteHeader "Target", "Type", "User"

    Dim keywords As Variant
    keywords = Array("MicrosoftOffice", "Microsoft_OC", "Office16", "Office15", _
                     "PowerBI", "SharePoint", "OneDrive", "microsoftonline", _
                     "login.windows.net", "AADToken", "MicrosoftSqlServer")

    Dim lines() As String: lines = Split(content, vbCrLf)
    Dim curTarget As String, curType As String, curUser As String
    curTarget = "": curType = "": curUser = ""
    Dim found As Long: found = 0

    Dim ln As Long
    For ln = 0 To UBound(lines)
        Dim t As String: t = Trim(lines(ln))

        ' [FIX #5] Handle both "Target:" and "Target Name:" variants
        If Left(t, 7) = "Target:" Then
            curTarget = Trim(Mid(t, 8))
            curType = "": curUser = ""
        ElseIf Left(t, 12) = "Target Name:" Then
            curTarget = Trim(Mid(t, 13))
            curType = "": curUser = ""
        ElseIf Left(t, 5) = "Type:" Then
            curType = Trim(Mid(t, 6))
        ElseIf Left(t, 5) = "User:" Then
            curUser = Trim(Mid(t, 6))
            Dim k As Long
            For k = 0 To UBound(keywords)
                If InStr(1, curTarget, keywords(k), vbTextCompare) > 0 Then
                    WriteRow curTarget, curType, curUser
                    found = found + 1
                    Exit For
                End If
            Next k
        End If
    Next ln

    If found = 0 Then
        WriteRow "(no Office/SharePoint/SQL credentials found)", "", ""
    End If
End Sub

' ── Report Setup & Finalize ──────────────────────────────────

' [FIX #2] Reset ws to Nothing before lookup to prevent stale reference
Private Sub InitReport()
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = SHEET_NAME
    Else
        ws.Cells.Clear
    End If

    row_ = 1
    With ws.Cells(row_, 1)
        .Value = "Auth Environment Report"
        .Font.Bold = True
        .Font.Size = 14
    End With
    ws.Cells(row_, 3).Value = "Generated: " & Format(Now, "yyyy-mm-dd hh:nn:ss")
    row_ = row_ + 1
End Sub

' [FIX #10] Uses MAX_COLS constant
Private Sub FinalizeReport()
    ws.Columns("A:" & Chr(64 + MAX_COLS)).AutoFit
    Dim c As Long
    For c = 1 To MAX_COLS
        If ws.Columns(c).ColumnWidth > 60 Then ws.Columns(c).ColumnWidth = 60
    Next c
    ws.Activate
    ws.Range("A1").Select
End Sub

' ── Public Button Macros ─────────────────────────────────────

' BUTTON 1 — Full Auth Scan
' [FIX #1] On Error GoTo Cleanup restores Application state on any failure
Public Sub ScanAuth()
    On Error GoTo Cleanup
    Application.ScreenUpdating = False
    Application.Cursor = xlWait

    InitReport
    ScanWindowsIdentity
    ScanOfficeIdentity
    ScanTokenStores
    ScanConnections
    ScanCredentialManager
    FinalizeReport

Cleanup:
    Dim scanErr As Long: scanErr = Err.Number
    Dim scanMsg As String: scanMsg = Err.Description
    Application.Cursor = xlDefault
    Application.ScreenUpdating = True
    If scanErr <> 0 Then
        MsgBox "Scan failed: " & scanMsg, vbCritical, "Error"
    Else
        MsgBox "Auth scan complete. Review the " & SHEET_NAME & " sheet.", _
               vbInformation, "Scan Complete"
    End If
End Sub

' BUTTON 2 — Refresh Data Connections (triggers re-auth prompts)
Public Sub RefreshConnections()
    If ThisWorkbook.Connections.Count = 0 Then
        MsgBox "No data connections in this workbook." & vbLf & _
               "Open your data workbook and run this there.", vbInformation
        Exit Sub
    End If

    Dim conn As WorkbookConnection
    Dim refreshed As Long: refreshed = 0
    Dim errored As Long:   errored = 0

    For Each conn In ThisWorkbook.Connections
        On Error Resume Next
        conn.Refresh
        If Err.Number <> 0 Then errored = errored + 1 Else refreshed = refreshed + 1
        Err.Clear
        On Error GoTo 0
    Next conn

    MsgBox "Refresh done." & vbLf & _
           "OK: " & refreshed & "  Errors: " & errored & vbLf & vbLf & _
           "Login prompt = re-auth working.", vbInformation, "Refresh"
End Sub
