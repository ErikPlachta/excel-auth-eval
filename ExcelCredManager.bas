Attribute VB_Name = "ExcelCredManager"
' =======================================================
' ExcelCredManager.bas  (v2)
' Companion to ExcelCredManager.ps1 — must be in same folder as workbook.
' Import: Alt+F11 > File > Import File
' Save workbook as .xlsm.
' =======================================================

Private Const PS1_NAME    As String = "ExcelCredManager.ps1"
Private Const REPORT_NAME As String = "cred_report.csv"
Private Const SHEET_NAME  As String = "CredScan"

' Returns path to PS1 — expected beside the workbook
Private Function PS1Path() As String
    PS1Path = ThisWorkbook.Path & "\" & PS1_NAME
End Function

Private Function ReportPath() As String
    ReportPath = Environ("TEMP") & "\" & REPORT_NAME
End Function

' Shells PS1 synchronously, waits for completion
Private Sub RunPS(action As String)
    If Dir(PS1Path()) = "" Then
        MsgBox "ExcelCredManager.ps1 not found next to this workbook." & vbLf & _
               "Expected: " & PS1Path(), vbExclamation, "Missing Script"
        Exit Sub
    End If
    Dim cmd As String
    cmd = "powershell.exe -ExecutionPolicy Bypass -NonInteractive -File """ & PS1Path() & _
          """ -Action " & action & " -ReportPath """ & ReportPath() & """"
    CreateObject("WScript.Shell").Run cmd, 0, True
End Sub

' Loads CSV report into the CredScan sheet
Private Sub LoadReport()
    Dim ws As Worksheet
    On Error Resume Next: Set ws = ThisWorkbook.Sheets(SHEET_NAME): On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = SHEET_NAME
    Else
        ws.Cells.Clear
    End If

    If Dir(ReportPath()) = "" Then
        ws.Range("A1").Value = "No report generated — scan may have failed."
        Exit Sub
    End If

    Dim fNum As Integer: fNum = FreeFile
    Dim row As Long: row = 1
    Open ReportPath() For Input As #fNum
    Do While Not EOF(fNum)
        Dim rawLine As String
        Line Input #fNum, rawLine
        Dim cols() As String: cols = SplitCSV(rawLine)
        Dim c As Long
        For c = 0 To UBound(cols)
            ws.Cells(row, c + 1).Value = cols(c)
        Next c
        row = row + 1
    Loop
    Close #fNum

    ' Style header
    With ws.Range("A1:H1")
        .Font.Bold = True
        .Interior.Color = RGB(30, 30, 30)
        .Font.Color = RGB(255, 255, 255)
    End With

    ' Colour-code Status column (col 6)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim r As Long
    For r = 2 To lastRow
        Dim status As String: status = ws.Cells(r, 6).Value
        Dim cell As Range: Set cell = ws.Cells(r, 6)
        Select Case True
            Case InStr(status, "Expired") > 0 Or status = "Found"
                cell.Interior.Color = RGB(255, 200, 200)
                cell.Font.Color = RGB(180, 0, 0)
            Case InStr(status, "Stale") > 0 Or InStr(status, "Expiring") > 0
                cell.Interior.Color = RGB(255, 240, 180)
                cell.Font.Color = RGB(160, 100, 0)
            Case status = "Valid" Or InStr(status, "Recent") > 0
                cell.Interior.Color = RGB(200, 240, 200)
                cell.Font.Color = RGB(0, 120, 0)
        End Select
    Next r

    ws.Columns("A:H").AutoFit
    ws.Activate
End Sub

' ── Button macros ─────────────────────────────────────────────────────────────

' BUTTON 1 — Scan
Sub ScanCredentials()
    RunPS "Scan"
    LoadReport
    MsgBox "Scan complete. Review the CredScan sheet.", vbInformation, "Scan"
End Sub

' BUTTON 2 — Inspect (alias for Scan — PS1 reports all detail in one pass)
Sub InspectAuth()
    ScanCredentials
End Sub

' BUTTON 3 — Clear Auth
Sub ClearAuth()
    If MsgBox("Delete expired/stale tokens?" & vbLf & vbLf & _
              "Excel will require re-authentication on next data refresh.", _
              vbYesNo + vbExclamation, "Confirm Clear") = vbNo Then Exit Sub
    RunPS "Clear"   ' PS1 clears then re-scans automatically
    LoadReport
    MsgBox "Auth cleared. Refresh your data connections to re-authenticate.", _
           vbInformation, "Clear Complete"
End Sub

' BUTTON 4 — Refresh & Re-Auth
Sub RefreshAndReAuth()
    If ThisWorkbook.Connections.Count = 0 Then
        MsgBox "No data connections in this workbook." & vbLf & _
               "Open your data workbook and run this there, or use Data > Refresh All.", _
               vbInformation: Exit Sub
    End If
    Dim conn As WorkbookConnection
    Dim refreshed As Long: refreshed = 0
    Dim errored As Long: errored = 0
    For Each conn In ThisWorkbook.Connections
        On Error Resume Next
        conn.Refresh
        If Err.Number <> 0 Then errored = errored + 1 Else refreshed = refreshed + 1
        Err.Clear: On Error GoTo 0
    Next conn
    MsgBox "Refresh triggered." & vbLf & "OK: " & refreshed & "  Errors: " & errored & vbLf & vbLf & _
           "Login prompt = re-auth working correctly.", vbInformation, "Refresh"
End Sub

' ── Helper: minimal CSV line splitter (handles quoted fields) ─────────────────
Private Function SplitCSV(line As String) As String()
    Dim result(50) As String
    Dim n As Long: n = 0
    Dim inQ As Boolean: inQ = False
    Dim buf As String: buf = ""
    Dim i As Long
    For i = 1 To Len(line)
        Dim ch As String: ch = Mid(line, i, 1)
        If ch = """" Then
            If inQ And Mid(line, i + 1, 1) = """" Then
                buf = buf & """": i = i + 1  ' escaped quote
            Else
                inQ = Not inQ
            End If
        ElseIf ch = "," And Not inQ Then
            result(n) = buf: n = n + 1: buf = ""
        Else
            buf = buf & ch
        End If
    Next i
    result(n) = buf: n = n + 1
    ReDim Preserve result(n - 1)
    SplitCSV = result
End Function
