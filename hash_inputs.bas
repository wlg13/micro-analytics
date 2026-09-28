Attribute VB_Name = "HashInputs"
' Make a de-identified copy of every file in inputs\ in inputs-dev\.
'
' Identifying columns (listed in IdHeaders, EmailHeaders, NameHeaders) are
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
' Safety:
'   - Files in inputs\ are opened read-only and never saved.
'   - If any column header is in none of the four lists below, the macro
'     stops before writing anything and lists those headers. Add each to a
'     hash list or to SafeHeaders, then run again. So a new column (e.g. a
'     free-text survey question) can't reach inputs-dev unreviewed.
'   - CSV files are processed as text, not opened in Excel, so leading
'     zeros, long numbers and dates in the other columns are not changed.

Option Explicit

Private Const BASE_DIR As String = "C:\Users\billg\OneDrive - The Pennsylvania State University\104\104 Database -- Micro-analytics"
Private Const KEY_FILE As String = "hash_key.txt"
Private Const HASH_HEX_CHARS As Long = 16

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

' ---- Header lists: edit these ----
' Matching ignores case; * and ? work as wildcards.
' IDs: trimmed, whole numbers written without decimals, then hashed.
Private Function IdHeaders() As Variant
    IdHeaders = Array("ID", "PSU ID", "Campus ID")
End Function

' Emails: trimmed and lowercased, then hashed (merge_rosters.R joins on
' lowercased email).
Private Function EmailHeaders() As Variant
    EmailHeaders = Array("Email", "sis_id")
End Function

' Names and other identifying text: trimmed, then hashed.
Private Function NameHeaders() As Variant
    NameHeaders = Array("Name", "First Name", "Last Name", "Pronouns")
End Function

' Not identifying: copied unchanged.
Private Function SafeHeaders() As Variant
    SafeHeaders = Array("Notify", "Units", "Program and Plan", "Level", "Status Note", "Section", _
                        "100874763*")
End Function

' ---- Macros to run ----

Public Sub HashInputsToDev()
    Dim msg As String
    msg = HashFolder(BASE_DIR & "\inputs", BASE_DIR & "\inputs-dev", BASE_DIR & "\" & KEY_FILE)
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
Public Function HashFolder(ByVal inDir As String, ByVal outDir As String, ByVal keyPath As String) As String
    Dim fso As Object, files As New Collection, unknown As Object
    Dim f As Variant, skipped As String, msg As String, n As Long

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

    Set unknown = CreateObject("Scripting.Dictionary")
    SetQuiet True

    ' Pass 1: check every header before writing anything.
    For Each f In files
        ProcessFile inDir & "\" & f, "", False, unknown
    Next
    If unknown.Count > 0 Then
        SetQuiet False
        Erase mKey
        HashFolder = "STOPPED: nothing was written. These columns are not in any header list:" & vbLf & "  " & _
                     Join(unknown.Keys, vbLf & "  ") & vbLf & vbLf & _
                     "Add each to IdHeaders, EmailHeaders or NameHeaders (to hash it) or to SafeHeaders " & _
                     "(to copy it unchanged), then run again."
        Exit Function
    End If

    ' Pass 2: write the hashed copies.
    If BCryptOpenAlgorithmProvider(mAlg, StrPtr("SHA256"), 0, BCRYPT_ALG_HANDLE_HMAC_FLAG) <> 0 Then
        Err.Raise vbObjectError + 1, , "Could not open the Windows SHA-256 provider."
    End If
    For Each f In files
        ProcessFile inDir & "\" & f, outDir & "\" & f, True, unknown
        n = n + 1
    Next
    Cleanup

    msg = "Hashed " & n & " file(s) from" & vbLf & "  " & inDir & vbLf & "into" & vbLf & "  " & outDir
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

Private Sub ProcessFile(ByVal inPath As String, ByVal outPath As String, ByVal doWrite As Boolean, unknown As Object)
    If LCase$(Right$(inPath, 4)) = ".csv" Then
        HashCsv inPath, outPath, doWrite, unknown
    Else
        HashWorkbook inPath, outPath, doWrite, unknown
    End If
End Sub

' ---- Header classification ----

' Returns "id", "email", "name", "safe", or "" if the header is in no list.
Private Function ClassifyHeader(ByVal h As String) As String
    h = Trim$(h)
    If h = "" Then Exit Function
    If InList(h, IdHeaders()) Then
        ClassifyHeader = "id"
    ElseIf InList(h, EmailHeaders()) Then
        ClassifyHeader = "email"
    ElseIf InList(h, NameHeaders()) Then
        ClassifyHeader = "name"
    ElseIf InList(h, SafeHeaders()) Then
        ClassifyHeader = "safe"
    End If
End Function

Private Function InList(ByVal h As String, ByVal patterns As Variant) As Boolean
    Dim p As Variant
    For Each p In patterns
        If LCase$(h) Like LCase$(p) Then InList = True: Exit Function
    Next
End Function

Private Sub AddUnknown(unknown As Object, ByVal s As String)
    If Not unknown.Exists(s) Then unknown.Add s, True
End Sub

' ---- Excel workbooks ----

Private Sub HashWorkbook(ByVal inPath As String, ByVal outPath As String, ByVal doWrite As Boolean, unknown As Object)
    Dim ws As Worksheet, fname As String
    fname = Mid$(inPath, InStrRev(inPath, "\") + 1)
    Set mOpenWb = Workbooks.Open(Filename:=inPath, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
    For Each ws In mOpenWb.Worksheets
        HashSheet ws, fname, doWrite, unknown
    Next
    If doWrite Then mOpenWb.SaveAs Filename:=outPath, FileFormat:=mOpenWb.FileFormat
    mOpenWb.Close SaveChanges:=False
    Set mOpenWb = Nothing
End Sub

Private Sub HashSheet(ws As Worksheet, ByVal fname As String, ByVal doWrite As Boolean, unknown As Object)
    Dim lastRow As Long, lastCol As Long, c As Long, r As Long
    Dim h As String, kind As String, s As String, rng As Range, vals As Variant

    If Application.WorksheetFunction.CountA(ws.Cells) = 0 Then Exit Sub
    lastRow = ws.Cells.Find("*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    lastCol = ws.Cells.Find("*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

    For c = 1 To lastCol
        h = ""
        If Not IsError(ws.Cells(1, c).Value) Then h = Trim$(CStr(ws.Cells(1, c).Value))
        kind = ClassifyHeader(h)
        If lastRow >= 2 Then Set rng = ws.Range(ws.Cells(2, c), ws.Cells(lastRow, c))

        If kind = "" Then
            If h <> "" Then
                AddUnknown unknown, fname & " [" & ws.Name & "]: " & h
            ElseIf lastRow >= 2 Then
                If Application.WorksheetFunction.CountA(rng) > 0 Then
                    AddUnknown unknown, fname & " [" & ws.Name & "]: unnamed column " & c
                End If
            End If
        ElseIf kind <> "safe" And doWrite And lastRow >= 2 Then
            If rng.Cells.Count = 1 Then
                ReDim vals(1 To 1, 1 To 1)
                vals(1, 1) = rng.Value2
            Else
                vals = rng.Value2
            End If
            For r = 1 To UBound(vals, 1)
                If Not IsEmpty(vals(r, 1)) And Not IsError(vals(r, 1)) Then
                    s = NormalizeValue(vals(r, 1), kind)
                    If s <> "" Then vals(r, 1) = HmacHex(s)
                End If
            Next
            rng.NumberFormat = "@"
            rng.Value2 = vals
        End If
    Next
End Sub

' ---- CSV files (as text, so untouched fields keep their exact characters) ----

Private Sub HashCsv(ByVal inPath As String, ByVal outPath As String, ByVal doWrite As Boolean, unknown As Object)
    Dim b() As Byte, ob() As Byte, hasBom As Boolean, start As Long, cp As Long, ok As Boolean
    Dim text As String, L As Long, pos As Long, raw As String, term As String, outRaw As String
    Dim fname As String, rowNum As Long, col As Long, nHead As Long
    Dim heads() As String, kinds() As String, kind As String, hname As String, v As String
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
    ReDim heads(0 To 0): ReDim kinds(0 To 0)
    ReDim parts(0 To 1023)

    If L > 0 Then
        Do
            NextField text, L, pos, raw, term
            col = col + 1
            outRaw = raw
            If rowNum = 1 Then
                ReDim Preserve heads(0 To col): ReDim Preserve kinds(0 To col)
                heads(col) = Trim$(Unquote(raw))
                kinds(col) = ClassifyHeader(heads(col))
                If kinds(col) = "" And heads(col) <> "" Then AddUnknown unknown, fname & ": " & heads(col)
            Else
                v = Trim$(Unquote(raw))
                kind = "": hname = ""
                If col <= nHead Then kind = kinds(col): hname = heads(col)
                If kind = "" Then
                    If hname = "" And v <> "" Then AddUnknown unknown, fname & ": unnamed column " & col
                ElseIf kind <> "safe" And v <> "" And doWrite Then
                    outRaw = HmacHex(NormalizeValue(Unquote(raw), kind))
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

' Same text form R uses: IDs stored as numbers lose any ".0", emails are lowercased.
Private Function NormalizeValue(ByVal v As Variant, ByVal kind As String) As String
    Dim s As String
    If VarType(v) = vbDouble Or VarType(v) = vbCurrency Or VarType(v) = vbLong Or VarType(v) = vbInteger Then
        If v = Fix(v) And Abs(v) < 1E+15 Then s = Format$(v, "0") Else s = CStr(v)
    Else
        s = CStr(v)
    End If
    s = Trim$(s)
    If kind = "email" Then s = LCase$(s)
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
