Attribute VB_Name = "modITRChecker"
'==============================================================================
' ITR Checker - Excel front-end for itr_status_checker.pyw
'==============================================================================
' What it does:
'   A button in Excel that reads your PAN / Password / AY columns, runs the
'   Python engine (which drives Edge through the Income-Tax portal), and writes
'   the results (Status, Acknowledgement No, Filing Date, Intimation, Verified)
'   into the columns you choose.
'
' It does NOT let Python touch this workbook (that would strip your macros and
' hit file-locks). Instead the macro exports the rows to a small job file, runs
' Python, and reads the results back itself.
'
' ONE-TIME SETUP:
'   1. Save this workbook as .xlsm (macro-enabled).
'   2. Alt+F11 -> File -> Import File... -> pick ITR_Checker.bas
'   3. Back in Excel: Alt+F8 -> run "ITR_Setup".
'      This creates an "ITR_Config" sheet and a "Run ITR Checker" button.
'   4. On the ITR_Config sheet, fill in:
'        - Python executable (usually just: python)
'        - Full path to itr_status_checker.pyw
'        - The data sheet name and the column letters for input & output
'   5. Click the "Run ITR Checker" button.
'==============================================================================
Option Explicit

Private Const CFG_SHEET As String = "ITR_Config"

'--- config cell addresses (column B on the ITR_Config sheet) ------------------
Private Const C_PYEXE As String = "B3"
Private Const C_PYW As String = "B4"
Private Const C_DATASHEET As String = "B5"
Private Const C_FIRSTROW As String = "B6"
Private Const C_LASTROW As String = "B7"
Private Const C_PAN As String = "B8"
Private Const C_PW As String = "B9"
Private Const C_AY As String = "B10"
Private Const C_STATUS As String = "B11"
Private Const C_ACK As String = "B12"
Private Const C_FILING As String = "B13"
Private Const C_INTIM As String = "B14"
Private Const C_VERIFIED As String = "B15"
Private Const C_NAME As String = "B16"
Private Const C_HEADLESS As String = "B17"


'==============================================================================
' ITR_Setup - create the config sheet + the Run button (run once)
'==============================================================================
Public Sub ITR_Setup()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CFG_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.Name = CFG_SHEET

        ws.Range("A1").Value = "ITR Checker - Configuration"
        ws.Range("A1").Font.Bold = True
        ws.Range("A1").Font.Size = 14

        WriteRow ws, "A3", "Python executable", C_PYEXE, "python"
        WriteRow ws, "A4", "Path to itr_status_checker.pyw", C_PYW, ""
        WriteRow ws, "A5", "Data sheet name", C_DATASHEET, ActiveSheetNameGuess()
        WriteRow ws, "A6", "First data row", C_FIRSTROW, "2"
        WriteRow ws, "A7", "Last data row (0 = auto)", C_LASTROW, "0"
        WriteRow ws, "A8", "PAN column (input)", C_PAN, "C"
        WriteRow ws, "A9", "Password column (input)", C_PW, "D"
        WriteRow ws, "A10", "AY column (input)", C_AY, "E"
        WriteRow ws, "A11", "Status column (output)", C_STATUS, "F"
        WriteRow ws, "A12", "Acknowledgement No column (output)", C_ACK, "G"
        WriteRow ws, "A13", "Filing Date column (output)", C_FILING, "H"
        WriteRow ws, "A14", "Intimation column (output)", C_INTIM, "I"
        WriteRow ws, "A15", "Verified column (output)", C_VERIFIED, "J"
        WriteRow ws, "A16", "Name column (optional)", C_NAME, ""
        WriteRow ws, "A17", "Headless (TRUE/FALSE)", C_HEADLESS, "FALSE"

        ws.Columns("A").ColumnWidth = 34
        ws.Columns("B").ColumnWidth = 46
        ws.Range("B3:B17").Interior.Color = RGB(255, 251, 224)
        ws.Range("B3:B17").Borders.LineStyle = xlContinuous
    End If

    EnsureButton ws
    ws.Activate
    MsgBox "Setup done." & vbCrLf & vbCrLf & _
           "Fill in the yellow cells on the '" & CFG_SHEET & "' sheet " & _
           "(especially the path to itr_status_checker.pyw and the column letters), " & _
           "then click the 'Run ITR Checker' button.", vbInformation, "ITR Checker"
End Sub


'==============================================================================
' ITR_Run - export rows, run Python, import results back into the sheet
'==============================================================================
Public Sub ITR_Run()
    Dim cfg As Worksheet
    Set cfg = CfgSheet()
    If cfg Is Nothing Then
        MsgBox "No config found. Run 'ITR_Setup' first.", vbExclamation, "ITR Checker"
        Exit Sub
    End If

    Dim pyExe As String, pyw As String, dataSheetName As String
    pyExe = GetCfg(cfg, C_PYEXE)
    If pyExe = "" Then pyExe = "python"
    pyw = GetCfg(cfg, C_PYW)
    dataSheetName = GetCfg(cfg, C_DATASHEET)

    If pyw = "" Or Dir(pyw) = "" Then
        MsgBox "Set a valid path to itr_status_checker.pyw in cell " & C_PYW & _
               " of the " & CFG_SHEET & " sheet.", vbExclamation, "ITR Checker"
        Exit Sub
    End If

    Dim ds As Worksheet
    On Error Resume Next
    Set ds = ThisWorkbook.Worksheets(dataSheetName)
    On Error GoTo 0
    If ds Is Nothing Then
        MsgBox "Data sheet '" & dataSheetName & "' not found (cell " & C_DATASHEET & ").", _
               vbExclamation, "ITR Checker"
        Exit Sub
    End If

    Dim panCol As Long, pwCol As Long, ayCol As Long
    panCol = ColNum(ds, GetCfg(cfg, C_PAN))
    pwCol = ColNum(ds, GetCfg(cfg, C_PW))
    ayCol = ColNum(ds, GetCfg(cfg, C_AY))
    If panCol = 0 Or pwCol = 0 Then
        MsgBox "Set the PAN and Password column letters (cells " & C_PAN & " and " & C_PW & ").", _
               vbExclamation, "ITR Checker"
        Exit Sub
    End If

    Dim firstRow As Long, lastRow As Long
    firstRow = Val(GetCfg(cfg, C_FIRSTROW))
    If firstRow < 1 Then firstRow = 2
    lastRow = Val(GetCfg(cfg, C_LASTROW))
    If lastRow < firstRow Then lastRow = ds.Cells(ds.Rows.Count, panCol).End(xlUp).Row

    Dim headless As String
    headless = LCase(GetCfg(cfg, C_HEADLESS))
    Dim headlessJson As String
    If headless = "true" Or headless = "yes" Or headless = "1" Then
        headlessJson = "true"
    Else
        headlessJson = "false"
    End If

    ' Build the rows JSON.
    Dim rowsJson As String, cnt As Long, r As Long
    Dim pan As String, pw As String, ay As String
    rowsJson = ""
    For r = firstRow To lastRow
        pan = Trim(CStr(ds.Cells(r, panCol).Value))
        pw = CStr(ds.Cells(r, pwCol).Value)
        If pan <> "" And pw <> "" And UCase(pan) <> "ABCDE1234F" Then
            ay = ""
            If ayCol > 0 Then ay = Trim(CStr(ds.Cells(r, ayCol).Value))
            If rowsJson <> "" Then rowsJson = rowsJson & ","
            rowsJson = rowsJson & "{""row"":" & r & _
                       ",""pan"":""" & JsonEsc(pan) & """" & _
                       ",""password"":""" & JsonEsc(pw) & """" & _
                       ",""ay"":""" & JsonEsc(ay) & """}"
            cnt = cnt + 1
        End If
    Next r

    If cnt = 0 Then
        MsgBox "No data rows found. Check the PAN/Password columns and the row range.", _
               vbExclamation, "ITR Checker"
        Exit Sub
    End If

    ' Temp file locations (workbook folder, or TEMP if unsaved).
    Dim base As String
    base = ThisWorkbook.Path
    If base = "" Then base = Environ$("TEMP")
    Dim jobPath As String, resPath As String, logPath As String
    jobPath = base & Application.PathSeparator & "itr_job.json"
    resPath = base & Application.PathSeparator & "itr_results.tsv"
    logPath = base & Application.PathSeparator & "itr_log.txt"

    Dim jobJson As String
    jobJson = "{""headless"":" & headlessJson & _
              ",""results"":""" & JsonEsc(resPath) & """" & _
              ",""log"":""" & JsonEsc(logPath) & """" & _
              ",""rows"":[" & rowsJson & "]}"

    ' Fresh results file each run.
    On Error Resume Next
    If Dir(resPath) <> "" Then Kill resPath
    On Error GoTo 0

    WriteUtf8 jobPath, jobJson

    If MsgBox(cnt & " row(s) will be processed." & vbCrLf & vbCrLf & _
              "An Edge window and a Python window will open. Keep them visible so " & _
              "you can solve any OTP/CAPTCHA. Excel will wait until it finishes." & vbCrLf & vbCrLf & _
              "Start now?", vbQuestion + vbYesNo, "ITR Checker") <> vbYes Then
        Exit Sub
    End If

    ' Run the engine and wait for it to finish.
    Dim cmd As String, rc As Long
    cmd = """" & pyExe & """ """ & pyw & """ --job """ & jobPath & """"
    Application.StatusBar = "ITR Checker: running " & cnt & " row(s)..."
    On Error GoTo RunErr
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    rc = wsh.Run(cmd, 1, True)      ' window normal, wait for return
    On Error GoTo 0

    ' Read results back into the sheet.
    Dim updated As Long
    updated = ITR_ImportResults(True)

    ' Cleanup (keep the log for troubleshooting).
    On Error Resume Next
    If Dir(jobPath) <> "" Then Kill jobPath
    If Dir(resPath) <> "" Then Kill resPath
    On Error GoTo 0

    Application.StatusBar = False
    MsgBox "Done. " & updated & " row(s) updated." & vbCrLf & _
           "Log: " & logPath, vbInformation, "ITR Checker"
    Exit Sub

RunErr:
    Application.StatusBar = False
    MsgBox "Could not start Python." & vbCrLf & _
           "Command: " & cmd & vbCrLf & vbCrLf & _
           "Check the 'Python executable' and the .pyw path in " & CFG_SHEET & ".", _
           vbCritical, "ITR Checker"
End Sub


'==============================================================================
' ITR_ImportResults - read itr_results.tsv and write into the output columns.
'   Can also be assigned to its own button (e.g. if you run Python separately).
'   Returns the number of rows updated.
'==============================================================================
Public Function ITR_ImportResults(Optional ByVal silent As Boolean = False) As Long
    Dim cfg As Worksheet
    Set cfg = CfgSheet()
    If cfg Is Nothing Then Exit Function

    Dim ds As Worksheet
    On Error Resume Next
    Set ds = ThisWorkbook.Worksheets(GetCfg(cfg, C_DATASHEET))
    On Error GoTo 0
    If ds Is Nothing Then Exit Function

    Dim base As String
    base = ThisWorkbook.Path
    If base = "" Then base = Environ$("TEMP")
    Dim resPath As String
    resPath = base & Application.PathSeparator & "itr_results.tsv"
    If Dir(resPath) = "" Then
        If Not silent Then MsgBox "No results file found (itr_results.tsv).", vbExclamation, "ITR Checker"
        Exit Function
    End If

    Dim cStatus As Long, cAck As Long, cFiling As Long, cIntim As Long, cVerified As Long
    cStatus = ColNum(ds, GetCfg(cfg, C_STATUS))
    cAck = ColNum(ds, GetCfg(cfg, C_ACK))
    cFiling = ColNum(ds, GetCfg(cfg, C_FILING))
    cIntim = ColNum(ds, GetCfg(cfg, C_INTIM))
    cVerified = ColNum(ds, GetCfg(cfg, C_VERIFIED))

    Dim raw As String
    raw = ReadUtf8(resPath)
    Dim lines() As String, i As Long, parts() As String, r As Long, n As Long
    raw = Replace(raw, vbCrLf, vbLf)
    raw = Replace(raw, vbCr, vbLf)
    lines = Split(raw, vbLf)

    For i = LBound(lines) To UBound(lines)
        If Len(Trim(lines(i))) > 0 Then
            parts = Split(lines(i), vbTab)
            If UBound(parts) >= 5 Then
                r = Val(parts(0))
                If r > 0 Then
                    PutText ds, r, cStatus, parts(1)
                    PutText ds, r, cAck, parts(2)
                    PutDate ds, r, cFiling, parts(3)
                    PutDate ds, r, cIntim, parts(4)
                    PutText ds, r, cVerified, parts(5)
                    n = n + 1
                End If
            End If
        End If
    Next i

    If Not silent Then MsgBox n & " row(s) imported.", vbInformation, "ITR Checker"
    ITR_ImportResults = n
End Function


'==============================================================================
' Helpers
'==============================================================================
Private Function CfgSheet() As Worksheet
    On Error Resume Next
    Set CfgSheet = ThisWorkbook.Worksheets(CFG_SHEET)
    On Error GoTo 0
End Function

Private Function GetCfg(ws As Worksheet, addr As String) As String
    GetCfg = Trim(CStr(ws.Range(addr).Value))
End Function

Private Sub WriteRow(ws As Worksheet, labelCell As String, label As String, valCell As String, dflt As String)
    ws.Range(labelCell).Value = label
    ws.Range(valCell).Value = dflt
End Sub

Private Function ActiveSheetNameGuess() As String
    On Error Resume Next
    ActiveSheetNameGuess = ThisWorkbook.Worksheets(1).Name
    On Error GoTo 0
End Function

Private Function ColNum(ws As Worksheet, letter As String) As Long
    letter = Trim(letter)
    If letter = "" Then Exit Function
    On Error Resume Next
    ColNum = ws.Range(letter & "1").Column
    On Error GoTo 0
End Function

Private Sub PutText(ws As Worksheet, r As Long, col As Long, v As String)
    If col > 0 Then ws.Cells(r, col).Value = v
End Sub

Private Sub PutDate(ws As Worksheet, r As Long, col As Long, v As String)
    If col = 0 Then Exit Sub
    If IsIsoDate(v) Then
        ws.Cells(r, col).Value = DateSerial(CInt(Left$(v, 4)), CInt(Mid$(v, 6, 2)), CInt(Mid$(v, 9, 2)))
        ws.Cells(r, col).NumberFormat = "dd-mm-yyyy"
    Else
        ws.Cells(r, col).NumberFormat = "General"
        ws.Cells(r, col).Value = v
    End If
End Sub

Private Function IsIsoDate(v As String) As Boolean
    If Len(v) <> 10 Then Exit Function
    If Mid$(v, 5, 1) <> "-" Or Mid$(v, 8, 1) <> "-" Then Exit Function
    If Not IsNumeric(Left$(v, 4)) Then Exit Function
    If Not IsNumeric(Mid$(v, 6, 2)) Then Exit Function
    If Not IsNumeric(Mid$(v, 9, 2)) Then Exit Function
    IsIsoDate = True
End Function

Private Function JsonEsc(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCr, "\r")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")
    JsonEsc = s
End Function

Private Sub WriteUtf8(path As String, content As String)
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2                 ' text
    st.Charset = "utf-8"
    st.Open
    st.WriteText content
    st.SaveToFile path, 2       ' 2 = overwrite
    st.Close
End Sub

Private Function ReadUtf8(path As String) As String
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile path
    ReadUtf8 = st.ReadText(-1)  ' -1 = read all
    st.Close
End Function

Private Sub EnsureButton(ws As Worksheet)
    Dim b As Object
    ' remove any old button we made, then add a fresh one
    On Error Resume Next
    For Each b In ws.Buttons
        If b.Caption = "Run ITR Checker" Then b.Delete
    Next b
    On Error GoTo 0
    Dim btn As Object
    Set btn = ws.Buttons.Add(ws.Range("D3").Left, ws.Range("D3").Top, 150, 40)
    btn.OnAction = "ITR_Run"
    btn.Caption = "Run ITR Checker"
    btn.Font.Size = 11
End Sub
