Attribute VB_Name = "HashInputs"
' Make a de-identified copy of every file in inputs\ in inputs-dev\.
'
' Columns marked "hash" in hash_columns.xlsx (IDs, emails, names, ...) are
' replaced by a keyed hash (HMAC-SHA256, first 16 hex characters). Every
' other cell is copied unchanged, so inputs-dev has the same files, sheets,
' column headers and column types as inputs. The same value always gets the
' same hash, so joins on ID or email still work on the hashed files.
'
' Setup (once):
'   1. In Excel press Alt+F11, then File > Import File... and pick this file.
'      (Import it into your Personal Macro Workbook or an .xlsm kept outside
'      the inputs folders.)
'   2. Run MakeHashKey. It creates hash_key.txt in BASE_DIR. Never share it,
'      commit it, or copy it into inputs-dev.
' Each time the inputs change:
'   Run HashInputsToDev.
'
' Column names:
'   hash_columns.xlsx (in BASE_DIR) lists every column name the macro knows,
'   each marked "hash" or "keep". It is created on the first run with the
'   roster and survey columns already filled in. When a file has a column
'   name that is not in the list, the macro shows it with a few example
'   values and asks whether to hash or keep it, then saves your answer, so
'   each new name is asked about only once. You can also open the file and
'   add or change rows yourself. Names ignore case; * is a wildcard (e.g.
'   "Quiz *" covers every column starting with "Quiz ").
'
' Safety:
'   - Files in inputs\ are opened read-only and never saved.
'   - Nothing is written until every column has an answer.
'   - CSV files are processed as text, not opened in Excel, so leading
'     zeros, long numbers and dates in the other columns are not changed.
'   - Hashed values that look like email addresses are lowercased first,
'     so Jane@PSU.edu and jane@psu.edu get the same hash.

Option Explicit

Private Const BASE_DIR As String = "C:\Users\billg\OneDrive - The Pennsylvania State University\104\104 Database -- Micro-analytics"
Private Const KEY_FILE As String = "hash_key.txt"
Private Const SETTINGS_FILE As String = "hash_columns.xlsx"
Private Const HASH_HEX_CHARS As Long = 16
Private Const MAX_SAMPLES As Long = 3

Private Const CP_UTF8 As Long = 65001
Private Const CP_ANSI As Long = 1252
Private Const MB_ERR_INVALID_CHARS As Long = 8
Private Const BCRYPT_ALG_HANDLE_HMAC_FLAG As Long = 8
Private Const BCRYPT_USE_SYSTEM_PREFERRED_RNG As Long = 2

Private Declare PtrSafe Function BCryptOpenAlgorithmProvider Lib "bcrypt.dll" (ByRef phAlgorithm As LongPtr, ByVal pszAlgId As LongPtr, ByVal pszImplementation As LongPtr, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function BCryptCloseAlgorithmProvider Lib "bcrypt.dll" (ByVal hAlgorithm As LongPtr, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function BCryptCreateHash Lib "bcrypt.dll" (ByVal hAlgorithm As LongPtr, ByRef phHash As LongPtr, ByVal pbHashObject As LongPtr, ByVal cbHashObject As Long, ByVal pbSecret As LongPtr, ByVal cbSecret As Long, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function BCryptHashData Lib "bcrypt.dll" (ByVal hHash As LongPtr, ByVal pbInput As LongPtr, ByVal cbInput As Long, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function BCryptFinishHash Lib "bcrypt.dll" (ByVal hHash As LongPtr, ByVal pbOutput As LongPtr, ByVal cbOutput As Long, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function BCryptDestroyHash Lib "bcrypt.dll" (ByVal hHash As LongPtr) As Long
Private Declare PtrSafe Function BCryptGenRandom Lib "bcrypt.dll" (ByVal hAlgorithm As LongPtr, ByVal pbBuffer As LongPtr, ByVal cbBuffer As Long, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long) As Long
Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, ByVal lpDefaultChar As LongPtr, ByVal lpUsedDefaultChar As LongPtr) As Long

Private mAlg As LongPtr
Private mKey() As Byte
Private mOpenWb As Workbook
Private mNames As Collection     ' column names (may contain *) from hash_columns.xlsx
Private mActions As Collection   ' "hash" or "keep", same order as mNames

' Rows written to hash_columns.xlsx when it is first created.
Private Function DefaultHashNames() As Variant
    DefaultHashNames = Array("ID", "PSU ID", "Campus ID", "Email", "sis_id", _
                             "Name", "First Name", "Last Name", "Pronouns")
End Function

Private Function DefaultKeepNames() As Variant
    DefaultKeepNames = Array("Notify", "Units", "Program and Plan", "Level", "Status Note", "Section", _
                             "100874763*")
End Function

' ---- Macros to run ----

Public Sub HashInputsToDev()
    Dim msg As String
    msg = HashFolder(BASE_DIR & "\inputs", BASE_DIR & "\inputs-dev", BASE_DIR & "\" & KEY_FILE, _
                     BASE_DIR & "\" & SETTINGS_FILE)
    MsgBox msg, IIf(Left$(msg, 7) = "STOPPED", vbExclamation, vbInformation), "Hash inputs"
End Sub

Public Sub MakeHashKey()
    Dim path As String, b(0 To 31) As Byte, f As Integer
    path = BASE_DIR & "\" & KEY_FILE
    If Len(Dir$(path)) > 0 Then
        MsgBox "A key already exists and was not changed:" & vbLf & path & vbLf & vbLf & _
               "Replacing it would change every hash.", vbExclamation, "Hash key"
        Exit Sub
    End If
    If BCryptGenRandom(0, VarPtr(b(0)), 32, BCRYPT_USE_SYSTEM_PREFERRED_RNG) <> 0 Then
        MsgBox "Could not generate a random key.", vbCritical, "Hash key"
        Exit Sub
    End If
    f = FreeFile
    Open path For Output As #f
    Print #f, BytesToHex(b, 32)
    Close #f
    MsgBox "Created " & path & vbLf & vbLf & _
           "Keep it private: never share it, commit it, or copy it into inputs-dev.", vbInformation, "Hash key"
End Sub

' ---- Core (public so it can be tested on other folders) ----

' Returns a summary; starts with "STOPPED" if nothing (or not everything) was written.
' testAnswer is for automated tests only: "hash", "keep" or "stop" answers every
' new-column question without showing a dialog.
Public Function HashFolder(ByVal inDir As String, ByVal outDir As String, ByVal keyPath As String, _
                           ByVal settingsPath As String, Optional ByVal testAnswer As String = "") As String
    Dim fso As Object, files As New Collection, newCols As Object, blockers As Object
    Dim f As Variant, k As Variant, it As Variant, ans As String
    Dim skipped As String, msg As String, n As Long
    Dim addNames As New Collection, addActions As New Collection, addFiles As New Collection

    On Error GoTo Fail
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(inDir) Then HashFolder = "STOPPED: input folder not found: " & inDir: Exit Function
    If Not fso.FolderExists(outDir) Then HashFolder = "STOPPED: output folder not found: " & outDir: Exit Function
    inDir = fso.GetAbsolutePathName(inDir)
    outDir = fso.GetAbsolutePathName(outDir)
    If StrComp(inDir, outDir, vbTextCompare) = 0 Then HashFolder = "STOPPED: input and output folders are the same.": Exit Function
    If Not fso.FileExists(keyPath) Then HashFolder = "STOPPED: no key file at " & keyPath & ". Run MakeHashKey first.": Exit Function
    msg = LoadKey(keyPath)
    If msg <> "" Then HashFolder = "STOPPED: " & msg: Exit Function

    For Each f In fso.GetFolder(inDir).files
        If Left$(f.Name, 2) <> "~$" Then
            Select Case LCase$(fso.GetExtensionName(f.Name))
                Case "csv", "xlsx", "xlsm", "xls": files.Add f.Name
                Case Else: skipped = skipped & vbLf & "  " & f.Name
            End Select
        End If
    Next

    SetQuiet True
    msg = LoadSettings(settingsPath)
    If msg <> "" Then Cleanup: HashFolder = "STOPPED: nothing was written. " & msg: Exit Function

    ' Pass 1: read every header; collect names not in hash_columns.xlsx.
    Set newCols = CreateObject("Scripting.Dictionary")
    Set blockers = CreateObject("Scripting.Dictionary")
    For Each f In files
        ProcessFile inDir & "\" & f, "", False, newCols, blockers
    Next
    If blockers.Count > 0 Then
        Cleanup
        HashFolder = "STOPPED: nothing was written. These columns have data but no name:" & vbLf & "  " & _
                     Join(blockers.Keys, vbLf & "  ") & vbLf & vbLf & "Give each a header, or delete it, then run again."
        Exit Function
    End If

    ' Ask about each new name once, and save the answers.
    If newCols.Count > 0 Then
        SetQuiet False
        For Each k In newCols.Keys
            it = newCols(k)
            ans = AskColumn(it(0), it(1), it(2), testAnswer)
            If ans = "stop" Then Exit For
            addNames.Add it(0): addActions.Add ans: addFiles.Add it(1)
            mNames.Add it(0): mActions.Add ans
        Next
        SetQuiet True
        If addNames.Count > 0 Then SaveAnswers settingsPath, addNames, addActions, addFiles
        If ans = "stop" Then
            Cleanup
            HashFolder = "STOPPED: nothing was written." & vbLf & addNames.Count & _
                         " answer(s) given before stopping were saved in " & settingsPath & "."
            Exit Function
        End If
    End If

    ' Pass 2: write the hashed copies.
    If BCryptOpenAlgorithmProvider(mAlg, StrPtr("SHA256"), 0, BCRYPT_ALG_HANDLE_HMAC_FLAG) <> 0 Then
        Err.Raise vbObjectError + 1, , "Could not open the Windows SHA-256 provider."
    End If
    For Each f In files
        ProcessFile inDir & "\" & f, outDir & "\" & f, True, newCols, blockers
        n = n + 1
    Next
    Cleanup

    msg = "Hashed " & n & " file(s) from" & vbLf & "  " & inDir & vbLf & "into" & vbLf & "  " & outDir
    If addNames.Count > 0 Then msg = msg & vbLf & vbLf & addNames.Count & " new column name(s) saved in " & settingsPath & "."
    If skipped <> "" Then msg = msg & vbLf & vbLf & "Skipped (not csv/xlsx/xls), not copied:" & skipped
    HashFolder = msg
    Exit Function

Fail:
    msg = Err.Description
    On Error Resume Next
    If Not mOpenWb Is Nothing Then mOpenWb.Close SaveChanges:=False
    Cleanup
    HashFolder = "STOPPED on error: " & msg & vbLf & _
                 "Files in the output folder may be incomplete. Fix the problem and run again."
End Function

Private Sub Cleanup()
    If mAlg <> 0 Then BCryptCloseAlgorithmProvider mAlg, 0
    mAlg = 0
    Set mOpenWb = Nothing
    Erase mKey
    SetQuiet False
End Sub

Private Sub SetQuiet(ByVal quiet As Boolean)
    Application.ScreenUpdating = Not quiet
    Application.DisplayAlerts = Not quiet
    Application.EnableEvents = Not quiet
End Sub

Private Sub ProcessFile(ByVal inPath As String, ByVal outPath As String, ByVal doWrite As Boolean, _
                        newCols As Object, blockers As Object)
    If LCase$(Right$(inPath, 4)) = ".csv" Then
        HashCsv inPath, outPath, doWrite, newCols, blockers
    Else
        HashWorkbook inPath, outPath, doWrite, newCols, blockers
    End If
End Sub

' ---- Column names (hash_columns.xlsx) ----

' Returns "hash", "keep", or "" if the name is not listed. First matching row wins.
Private Function ColumnAction(ByVal h As String) As String
    Dim i As Long
    h = LCase$(Trim$(h))
    If h = "" Then Exit Function
    For i = 1 To mNames.Count
        If h Like LikePattern(mNames(i)) Then ColumnAction = mActions(i): Exit Function
    Next
End Function

' Only * is a wildcard; [ # ? in names are matched literally.
Private Function LikePattern(ByVal nm As String) As String
    nm = LCase$(Trim$(nm))
    nm = Replace(nm, "[", "[[]")
    nm = Replace(nm, "#", "[#]")
    nm = Replace(nm, "?", "[?]")
    LikePattern = nm
End Function

' Returns "" on success, else a problem description.
Private Function LoadSettings(ByVal path As String) As String
    Dim wb As Workbook, ws As Worksheet, mine As Boolean
    Dim r As Long, lastRow As Long, nm As String, act As String, bad As String

    Set mNames = New Collection
    Set mActions = New Collection
    If Len(Dir$(path)) = 0 Then CreateSettings path

    Set wb = OpenWb(path, True, mine)
    Set ws = wb.Worksheets(1)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        nm = "": act = ""
        If Not IsError(ws.Cells(r, 1).Value) Then nm = Trim$(CStr(ws.Cells(r, 1).Value))
        If Not IsError(ws.Cells(r, 2).Value) Then act = LCase$(Trim$(CStr(ws.Cells(r, 2).Value)))
        If nm <> "" Then
            If act = "hash" Or act = "keep" Then
                mNames.Add nm
                mActions.Add act
            Else
                bad = bad & vbLf & "  row " & r & ": " & nm & " -> """ & act & """"
            End If
        End If
    Next
    If mine Then wb.Close SaveChanges:=False
    If bad <> "" Then LoadSettings = "In " & path & ", the Action column must be hash or keep:" & bad
End Function

Private Sub CreateSettings(ByVal path As String)
    Dim wb As Workbook, ws As Worksheet, nm As Variant, r As Long
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Name = "Columns"
    ws.Range("A1:C1").Value = Array("Column name", "Action", "First seen in")
    ws.Range("A1:C1").Font.Bold = True
    ws.Range("E1").Value = "Action: hash = replace values with a keyed hash (anything that identifies a student); " & _
                           "keep = copy unchanged. Names ignore case; * is a wildcard."
    r = 2
    For Each nm In DefaultHashNames()
        ws.Cells(r, 1).Value = nm: ws.Cells(r, 2).Value = "hash": ws.Cells(r, 3).Value = "(default)"
        r = r + 1
    Next
    For Each nm In DefaultKeepNames()
        ws.Cells(r, 1).Value = nm: ws.Cells(r, 2).Value = "keep": ws.Cells(r, 3).Value = "(default)"
        r = r + 1
    Next
    ws.Columns("A:C").AutoFit
    wb.SaveAs Filename:=path, FileFormat:=xlOpenXMLWorkbook
    wb.Close SaveChanges:=False
End Sub

Private Sub SaveAnswers(ByVal path As String, names As Collection, actions As Collection, files As Collection)
    Dim wb As Workbook, ws As Worksheet, mine As Boolean, r As Long, i As Long
    Set wb = OpenWb(path, False, mine)
    Set ws = wb.Worksheets(1)
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    For i = 1 To names.Count
        ws.Cells(r, 1).NumberFormat = "@"
        ws.Cells(r, 1).Value = names(i)
        ws.Cells(r, 2).Value = actions(i)
        ws.Cells(r, 3).Value = files(i)
        r = r + 1
    Next
    wb.Save
    If mine Then wb.Close SaveChanges:=False
End Sub

' Uses the workbook if it is already open in Excel (mine = False), else opens it.
Private Function OpenWb(ByVal path As String, ByVal readOnly As Boolean, mine As Boolean) As Workbook
    Dim w As Workbook, nm As String
    nm = Mid$(path, InStrRev(path, "\") + 1)
    For Each w In Workbooks
        If StrComp(w.Name, nm, vbTextCompare) = 0 Then Set OpenWb = w: mine = False: Exit Function
    Next
    Set OpenWb = Workbooks.Open(Filename:=path, UpdateLinks:=0, ReadOnly:=readOnly, AddToMru:=False)
    mine = True
End Function

' Returns "hash", "keep" or "stop".
Private Function AskColumn(ByVal h As String, ByVal fname As String, ByVal samples As String, _
                           ByVal testAnswer As String) As String
    Dim prompt As String
    If testAnswer <> "" Then AskColumn = LCase$(testAnswer): Exit Function
    If samples = "" Then samples = "  (all blank)" & vbLf
    prompt = "New column name:  " & h & vbLf & "In file:  " & fname & vbLf & vbLf & _
             "Example values:" & vbLf & samples & vbLf & _
             "Could this column identify a student (name, ID, email, username, free-text answer)?" & vbLf & vbLf & _
             "Yes  = HASH it" & vbLf & "No   = KEEP it unchanged" & vbLf & "Cancel = stop (nothing is written)" & vbLf & vbLf & _
             "Your answer is saved in hash_columns.xlsx, so you won't be asked about this name again."
    Select Case MsgBox(prompt, vbYesNoCancel + vbQuestion + vbDefaultButton1, "New column")
        Case vbYes: AskColumn = "hash"
        Case vbNo: AskColumn = "keep"
        Case Else: AskColumn = "stop"
    End Select
End Function

' Records a header that is not in hash_columns.xlsx (once per name, ignoring case).
Private Sub NoteNew(newCols As Object, ByVal h As String, ByVal fname As String)
    If Not newCols.Exists(LCase$(h)) Then newCols.Add LCase$(h), Array(h, fname, "", 0)
End Sub

Private Sub AddSample(newCols As Object, ByVal h As String, ByVal v As String)
    Dim it As Variant
    If v = "" Or Not newCols.Exists(LCase$(h)) Then Exit Sub
    it = newCols(LCase$(h))
    If it(3) >= MAX_SAMPLES Then Exit Sub
    If Len(v) > 40 Then v = Left$(v, 40) & "..."
    it(2) = it(2) & "  " & v & vbLf
    it(3) = it(3) + 1
    newCols(LCase$(h)) = it
End Sub

Private Sub AddBlocker(blockers As Object, ByVal s As String)
    If Not blockers.Exists(s) Then blockers.Add s, True
End Sub

' ---- Excel workbooks ----

Private Sub HashWorkbook(ByVal inPath As String, ByVal outPath As String, ByVal doWrite As Boolean, _
                         newCols As Object, blockers As Object)
    Dim ws As Worksheet, fname As String
    fname = Mid$(inPath, InStrRev(inPath, "\") + 1)
    Set mOpenWb = Workbooks.Open(Filename:=inPath, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
    For Each ws In mOpenWb.Worksheets
        HashSheet ws, fname, doWrite, newCols, blockers
    Next
    If doWrite Then mOpenWb.SaveAs Filename:=outPath, FileFormat:=mOpenWb.FileFormat
    mOpenWb.Close SaveChanges:=False
    Set mOpenWb = Nothing
End Sub

Private Sub HashSheet(ws As Worksheet, ByVal fname As String, ByVal doWrite As Boolean, _
                      newCols As Object, blockers As Object)
    Dim lastRow As Long, lastCol As Long, c As Long, r As Long
    Dim h As String, act As String, s As String, rng As Range, vals As Variant

    If Application.WorksheetFunction.CountA(ws.Cells) = 0 Then Exit Sub
    lastRow = ws.Cells.Find("*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    lastCol = ws.Cells.Find("*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

    For c = 1 To lastCol
        h = ""
        If Not IsError(ws.Cells(1, c).Value) Then h = Trim$(CStr(ws.Cells(1, c).Value))
        act = ColumnAction(h)
        vals = Empty
        If lastRow >= 2 Then
            Set rng = ws.Range(ws.Cells(2, c), ws.Cells(lastRow, c))
            If rng.Cells.Count = 1 Then
                ReDim vals(1 To 1, 1 To 1)
                vals(1, 1) = rng.Value2
            Else
                vals = rng.Value2
            End If
        End If

        If act = "" Then
            If h <> "" Then
                NoteNew newCols, h, fname
                If Not IsEmpty(vals) Then
                    For r = 1 To UBound(vals, 1)
                        If Not IsError(vals(r, 1)) Then AddSample newCols, h, Trim$(CStr(vals(r, 1)))
                    Next
                End If
            ElseIf Not IsEmpty(vals) Then
                If Application.WorksheetFunction.CountA(rng) > 0 Then
                    AddBlocker blockers, fname & " [" & ws.Name & "]: column " & c
                End If
            End If
        ElseIf act = "hash" And doWrite And Not IsEmpty(vals) Then
            For r = 1 To UBound(vals, 1)
                If Not IsEmpty(vals(r, 1)) And Not IsError(vals(r, 1)) Then
                    s = NormalizeValue(vals(r, 1))
                    If s <> "" Then vals(r, 1) = HmacHex(s)
                End If
            Next
            rng.NumberFormat = "@"
            rng.Value2 = vals
        End If
    Next
End Sub

' ---- CSV files (as text, so untouched fields keep their exact characters) ----

Private Sub HashCsv(ByVal inPath As String, ByVal outPath As String, ByVal doWrite As Boolean, _
                    newCols As Object, blockers As Object)
    Dim b() As Byte, ob() As Byte, hasBom As Boolean, start As Long, cp As Long, ok As Boolean
    Dim text As String, L As Long, pos As Long, raw As String, term As String, outRaw As String
    Dim fname As String, rowNum As Long, col As Long, nHead As Long
    Dim heads() As String, acts() As String, act As String, hname As String, v As String
    Dim parts() As String, np As Long, i As Long

    fname = Mid$(inPath, InStrRev(inPath, "\") + 1)
    b = ReadFileBytes(inPath)
    If UBound(b) >= 2 Then hasBom = (b(0) = &HEF And b(1) = &HBB And b(2) = &HBF)
    If hasBom Then start = 3
    cp = CP_UTF8
    text = DecodeText(b, start, CP_UTF8, MB_ERR_INVALID_CHARS, ok)
    If Not ok Then
        cp = CP_ANSI
        text = DecodeText(b, start, CP_ANSI, 0, ok)
    End If

    L = Len(text)
    pos = 1
    rowNum = 1
    ReDim heads(0 To 0): ReDim acts(0 To 0)
    ReDim parts(0 To 1023)

    If L > 0 Then
        Do
            NextField text, L, pos, raw, term
            col = col + 1
            outRaw = raw
            If rowNum = 1 Then
                ReDim Preserve heads(0 To col): ReDim Preserve acts(0 To col)
                heads(col) = Trim$(Unquote(raw))
                acts(col) = ColumnAction(heads(col))
                If acts(col) = "" And heads(col) <> "" Then NoteNew newCols, heads(col), fname
            Else
                v = Trim$(Unquote(raw))
                act = "": hname = ""
                If col <= nHead Then act = acts(col): hname = heads(col)
                If act = "" Then
                    If hname <> "" Then
                        AddSample newCols, hname, v
                    ElseIf v <> "" Then
                        AddBlocker blockers, fname & ": column " & col
                    End If
                ElseIf act = "hash" And v <> "" And doWrite Then
                    outRaw = HmacHex(NormalizeValue(Unquote(raw)))
                End If
            End If

            If doWrite Then
                If np > UBound(parts) Then ReDim Preserve parts(0 To 2 * UBound(parts) + 1)
                parts(np) = outRaw & term
                np = np + 1
            End If

            If term = "" Then Exit Do
            If term <> "," Then
                If rowNum = 1 Then nHead = col
                rowNum = rowNum + 1
                col = 0
                If pos > L Then Exit Do
            End If
        Loop
    End If

    If Not doWrite Then Exit Sub
    If np > 0 Then
        ReDim Preserve parts(0 To np - 1)
        ob = EncodeText(Join(parts, ""), cp)
    Else
        ob = ""
    End If
    If hasBom Then
        Dim withBom() As Byte
        ReDim withBom(0 To UBound(ob) + 3)
        withBom(0) = &HEF: withBom(1) = &HBB: withBom(2) = &HBF
        For i = 0 To UBound(ob)
            withBom(i + 3) = ob(i)
        Next
        ob = withBom
    End If
    WriteFileBytes outPath, ob
End Sub

' Reads one CSV field starting at pos. raw is the field exactly as written
' (quotes included); term is the "," or line ending after it, or "" at end of file.
Private Sub NextField(text As String, ByVal L As Long, pos As Long, raw As String, term As String)
    Dim p As Long, q As Long, ch As String, startPos As Long
    startPos = pos
    If pos > L Then raw = "": term = "": Exit Sub

    p = pos
    If Mid$(text, pos, 1) = """" Then
        p = pos + 1
        Do
            q = InStr(p, text, """")
            If q = 0 Then p = L + 1: Exit Do
            If Mid$(text, q + 1, 1) = """" Then
                p = q + 2
            Else
                p = q + 1
                Exit Do
            End If
        Loop
    End If
    Do While p <= L
        ch = Mid$(text, p, 1)
        If ch = "," Or ch = vbCr Or ch = vbLf Then Exit Do
        p = p + 1
    Loop

    raw = Mid$(text, startPos, p - startPos)
    If p > L Then
        term = ""
        pos = p
    ElseIf Mid$(text, p, 2) = vbCrLf Then
        term = vbCrLf
        pos = p + 2
    Else
        term = Mid$(text, p, 1)
        pos = p + 1
    End If
End Sub

Private Function Unquote(ByVal raw As String) As String
    Dim q As Long
    If Left$(raw, 1) = """" Then
        q = InStrRev(raw, """")
        If q > 1 Then raw = Mid$(raw, 2, q - 2) Else raw = Mid$(raw, 2)
        raw = Replace(raw, """""", """")
    End If
    Unquote = raw
End Function

' ---- Hashing ----

' Same text form R uses: numbers stored in Excel lose any ".0"; values that
' look like email addresses are lowercased (merge_rosters.R joins on
' lowercased email).
Private Function NormalizeValue(ByVal v As Variant) As String
    Dim s As String
    If VarType(v) = vbDouble Or VarType(v) = vbCurrency Or VarType(v) = vbLong Or VarType(v) = vbInteger Then
        If v = Fix(v) And Abs(v) < 1E+15 Then s = Format$(v, "0") Else s = CStr(v)
    Else
        s = CStr(v)
    End If
    s = Trim$(s)
    If s Like "*?@?*.?*" And InStr(s, " ") = 0 Then s = LCase$(s)
    NormalizeValue = s
End Function

Private Function HmacHex(ByVal s As String) As String
    Dim hHash As LongPtr, data() As Byte, d(0 To 31) As Byte
    data = EncodeText(s, CP_UTF8)
    If BCryptCreateHash(mAlg, hHash, 0, 0, VarPtr(mKey(0)), UBound(mKey) + 1, 0) <> 0 Then
        Err.Raise vbObjectError + 2, , "Could not create a hash."
    End If
    BCryptHashData hHash, VarPtr(data(0)), UBound(data) + 1, 0
    BCryptFinishHash hHash, VarPtr(d(0)), 32, 0
    BCryptDestroyHash hHash
    HmacHex = BytesToHex(d, HASH_HEX_CHARS \ 2)
End Function

Private Function BytesToHex(b() As Byte, ByVal n As Long) As String
    Dim i As Long, s As String
    For i = 0 To n - 1
        s = s & Right$("0" & LCase$(Hex$(b(i))), 2)
    Next
    BytesToHex = s
End Function

' Returns "" on success, else a problem description.
Private Function LoadKey(ByVal keyPath As String) As String
    Dim b() As Byte, s As String, ok As Boolean
    b = ReadFileBytes(keyPath)
    s = DecodeText(b, 0, CP_UTF8, 0, ok)
    s = Replace(Replace(Replace(Replace(s, vbCr, ""), vbLf, ""), vbTab, ""), ChrW$(&HFEFF), "")
    s = Trim$(s)
    If Len(s) < 16 Then LoadKey = "the key in " & keyPath & " is missing or too short.": Exit Function
    mKey = EncodeText(s, CP_UTF8)
End Function

' ---- Bytes and text ----

Private Function ReadFileBytes(ByVal path As String) As Byte()
    Dim f As Integer, b() As Byte
    f = FreeFile
    Open path For Binary Access Read As #f
    If LOF(f) > 0 Then
        ReDim b(0 To LOF(f) - 1)
        Get #f, , b
    Else
        b = ""
    End If
    Close #f
    ReadFileBytes = b
End Function

Private Sub WriteFileBytes(ByVal path As String, b() As Byte)
    Dim f As Integer
    If Len(Dir$(path)) > 0 Then Kill path
    f = FreeFile
    Open path For Binary Access Write As #f
    If UBound(b) >= 0 Then Put #f, , b
    Close #f
End Sub

Private Function DecodeText(b() As Byte, ByVal start As Long, ByVal cp As Long, ByVal flags As Long, ok As Boolean) As String
    Dim cnt As Long, n As Long, s As String
    ok = True
    cnt = UBound(b) - start + 1
    If cnt <= 0 Then Exit Function
    n = MultiByteToWideChar(cp, flags, VarPtr(b(start)), cnt, 0, 0)
    If n = 0 Then ok = False: Exit Function
    s = String$(n, vbNullChar)
    MultiByteToWideChar cp, flags, VarPtr(b(start)), cnt, StrPtr(s), n
    DecodeText = s
End Function

Private Function EncodeText(ByVal s As String, ByVal cp As Long) As Byte()
    Dim n As Long, b() As Byte
    If Len(s) = 0 Then
        b = ""
    Else
        n = WideCharToMultiByte(cp, 0, StrPtr(s), Len(s), 0, 0, 0, 0)
        ReDim b(0 To n - 1)
        WideCharToMultiByte cp, 0, StrPtr(s), Len(s), VarPtr(b(0)), n, 0, 0
    End If
    EncodeText = b
End Function
