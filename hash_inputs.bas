Attribute VB_Name = "HashInputs"
' Make a de-identified copy of files in inputs\ in inputs-dev\.
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
' Each time you add or update a file in inputs\:
'   Run HashFileToDev and pick the file(s). It asks before writing each one.
' To redo every file (e.g. after changing an answer in hash_columns.xlsx):
'   Run HashInputsToDev. It lists all the files and asks once before writing.
' HashMacroVersion shows which version of this file Excel has.
'
' Column names:
'   hash_columns.xlsx (in BASE_DIR) lists every column name the macro knows,
'   each with Action "hash" or "keep". When the picked file(s) have column
'   names that are not in the list, the macro (after asking) adds them all
'   to hash_columns.xlsx with a suggested Action and "NEW - ..." in the
'   Check column, saying why (counts only, never student answers), opens
'   the file and stops. Check each NEW row, change Action if needed, delete
'   the NEW text, save, and run again. Nothing is written to inputs-dev
'   while any column of the picked files is still marked NEW.
'   In names, case is ignored, * matches any characters and # one digit
'   (e.g. "Quiz *", or "#" and "##" for Canvas points columns "1", "12").
'
' Safety:
'   - Files in inputs\ are opened read-only and never saved.
'   - Nothing is written to inputs-dev until every column is reviewed and
'     you click OK.
'   - CSV files are processed as text, not opened in Excel, so leading
'     zeros, long numbers and dates in the other columns are not changed.
'   - Hashed values that look like email addresses are lowercased first,
'     so Jane@PSU.edu and jane@psu.edu get the same hash.

Option Explicit

Private Const MACRO_VERSION As String = "2026-10-07"
Private Const BASE_DIR As String = "C:\Users\billg\OneDrive - The Pennsylvania State University\104\104 Database -- Micro-analytics"
Private Const KEY_FILE As String = "hash_key.txt"
Private Const SETTINGS_FILE As String = "hash_columns.xlsx"
Private Const HASH_HEX_CHARS As Long = 16

Private Const CP_UTF8 As Long = 65001
Private Const CP_ANSI As Long = 1252
Private Const MB_ERR_INVALID_CHARS As Long = 8
Private Const BCRYPT_ALG_HANDLE_HMAC_FLAG As Long = 8
Private Const BCRYPT_USE_SYSTEM_PREFERRED_RNG As Long = 2

' Positions in a new-column record (see NoteNew).
Private Const NC_NAME As Long = 0
Private Const NC_FILE As Long = 1
Private Const NC_DISTINCT As Long = 2
Private Const NC_COUNT As Long = 3
Private Const NC_EMAIL As Long = 4
Private Const NC_ID As Long = 5
Private Const NC_NUM As Long = 6
Private Const NC_LEN As Long = 7
Private Const NC_MAXLEN As Long = 8

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
Private mNames As Collection     ' column names (may contain * and #) from hash_columns.xlsx
Private mActions As Collection   ' "hash", "keep" or "pending" (Check says NEW), same order

' Rows written to hash_columns.xlsx when it is first created.
Private Function DefaultHashNames() As Variant
    DefaultHashNames = Array("ID", "PSU ID", "Campus ID", "Email", "sis_id", _
                             "Name", "First Name", "Last Name", "Pronouns")
End Function

Private Function DefaultKeepNames() As Variant
    DefaultKeepNames = Array("Notify", "Units", "Program and Plan", "Level", "Status Note", "Section", _
                             "section_id", "section_sis_id", "submitted", "attempt", _
                             "n correct", "n incorrect", "score", "#", "##")
End Function

' ---- Macros to run ----

' Hash one or more files you pick in inputs\, asking before each is written.
Public Sub HashFileToDev()
    Dim fso As Object, inDir As String, p As Variant, list As String, msg As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    inDir = fso.GetAbsolutePathName(BASE_DIR & "\inputs")
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Pick the file(s) in inputs to hash into inputs-dev"
        .InitialFileName = inDir & "\"
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "Data files", "*.csv; *.xlsx; *.xlsm; *.xls"
        If .Show <> -1 Then Exit Sub
        For Each p In .SelectedItems
            If StrComp(fso.GetParentFolderName(p), inDir, vbTextCompare) <> 0 Then
                MsgBox "Only files directly in" & vbLf & "  " & inDir & vbLf & "can be hashed. Nothing was written.", _
                       vbExclamation, "Hash file"
                Exit Sub
            End If
            list = list & "|" & fso.GetFileName(p)
        Next
    End With
    msg = HashFileList(inDir, BASE_DIR & "\inputs-dev", BASE_DIR & "\" & KEY_FILE, _
                       BASE_DIR & "\" & SETTINGS_FILE, Mid$(list, 2))
    MsgBox msg, IIf(Left$(msg, 7) = "STOPPED", vbExclamation, vbInformation), "Hash file"
End Sub

' Hash every file in inputs\, after one confirmation listing what will be written.
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

Public Sub HashMacroVersion()
    MsgBox "hash_inputs.bas version " & MACRO_VERSION, vbInformation, "Hash macros"
End Sub

' ---- Core (public so it can be tested on other folders) ----

' Both return a summary that starts with "STOPPED" if nothing was written to
' the output folder. testConfirm is for automated tests only: "yes" or "no"
' answers every confirmation without showing a dialog, and the column list
' is not opened for review.

' Every csv/xlsx/xls file in inDir, with one confirmation for all of them.
Public Function HashFolder(ByVal inDir As String, ByVal outDir As String, ByVal keyPath As String, _
                           ByVal settingsPath As String, Optional ByVal testConfirm As String = "") As String
    Dim fso As Object, files As New Collection, f As Variant, skipped As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(inDir) Then HashFolder = "STOPPED: input folder not found: " & inDir: Exit Function
    For Each f In fso.GetFolder(inDir).files
        If Left$(f.Name, 2) <> "~$" Then
            If IsDataFile(f.Name) Then files.Add f.Name Else skipped = skipped & vbLf & "  " & f.Name
        End If
    Next
    HashFolder = RunHash(inDir, outDir, keyPath, settingsPath, files, skipped, False, testConfirm)
End Function

' The files named in fileList (names in inDir separated by "|"), with one
' confirmation per file.
Public Function HashFileList(ByVal inDir As String, ByVal outDir As String, ByVal keyPath As String, _
                             ByVal settingsPath As String, ByVal fileList As String, _
                             Optional ByVal testConfirm As String = "") As String
    Dim fso As Object, files As New Collection, f As Variant
    Set fso = CreateObject("Scripting.FileSystemObject")
    For Each f In Split(fileList, "|")
        If Not fso.FileExists(inDir & "\" & f) Then
            HashFileList = "STOPPED: nothing was written. File not found: " & inDir & "\" & f: Exit Function
        End If
        If Not IsDataFile(f) Then
            HashFileList = "STOPPED: nothing was written. Not a csv/xlsx/xls file: " & f: Exit Function
        End If
        files.Add CStr(f)
    Next
    If files.Count = 0 Then HashFileList = "STOPPED: no file was picked.": Exit Function
    HashFileList = RunHash(inDir, outDir, keyPath, settingsPath, files, "", True, testConfirm)
End Function

Private Function IsDataFile(ByVal nm As String) As Boolean
    Select Case LCase$(Mid$(nm, InStrRev(nm, ".") + 1))
        Case "csv", "xlsx", "xlsm", "xls": IsDataFile = True
    End Select
End Function

Private Function RunHash(ByVal inDir As String, ByVal outDir As String, ByVal keyPath As String, _
                         ByVal settingsPath As String, files As Collection, ByVal skipped As String, _
                         ByVal perFile As Boolean, ByVal testConfirm As String) As String
    Dim fso As Object, newCols As Object, pending As Object, blockers As Object
    Dim f As Variant, k As Variant, it As Variant, prompt As String, settingsName As String, review As String
    Dim msg As String, written As String, declined As String, nWritten As Long, i As Long

    On Error GoTo Fail
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(inDir) Then RunHash = "STOPPED: input folder not found: " & inDir: Exit Function
    If Not fso.FolderExists(outDir) Then RunHash = "STOPPED: output folder not found: " & outDir: Exit Function
    inDir = fso.GetAbsolutePathName(inDir)
    outDir = fso.GetAbsolutePathName(outDir)
    If StrComp(inDir, outDir, vbTextCompare) = 0 Then RunHash = "STOPPED: input and output folders are the same.": Exit Function
    If Not fso.FileExists(keyPath) Then RunHash = "STOPPED: no key file at " & keyPath & ". Run MakeHashKey first.": Exit Function
    msg = LoadKey(keyPath)
    If msg <> "" Then RunHash = "STOPPED: " & msg: Exit Function
    settingsName = fso.GetFileName(settingsPath)

    SetQuiet True
    msg = LoadSettings(settingsPath)
    If msg <> "" Then Cleanup: RunHash = "STOPPED: nothing was written. " & msg: Exit Function

    ' Pass 1: read every header; collect names that are new or still marked NEW.
    Set newCols = CreateObject("Scripting.Dictionary")
    Set pending = CreateObject("Scripting.Dictionary")
    Set blockers = CreateObject("Scripting.Dictionary")
    For Each f In files
        ProcessFile inDir & "\" & f, "", False, newCols, pending, blockers
    Next
    If blockers.Count > 0 Then
        Cleanup
        RunHash = "STOPPED: nothing was written. These columns have data but no name:" & vbLf & "  " & _
                  Join(blockers.Keys, vbLf & "  ") & vbLf & vbLf & "Give each a header, or delete it, then run again."
        Exit Function
    End If

    ' New or unreviewed columns: add new ones to the list for review, then stop.
    If newCols.Count > 0 Or pending.Count > 0 Then
        If newCols.Count > 0 Then
            prompt = "Found " & newCols.Count & " column name(s) that are not in " & settingsName & "." & vbLf & vbLf & _
                     "Add them to " & settingsName & " with a suggested hash or keep for each, " & _
                     "so you can review them there?" & vbLf & vbLf & _
                     "Nothing is written to inputs-dev this time." & vbLf & vbLf & _
                     "OK = add them" & vbLf & "Cancel = stop (nothing is changed)"
            If Not Confirm(prompt, testConfirm) Then
                Cleanup
                RunHash = "STOPPED: nothing was changed."
                Exit Function
            End If
            AppendNewRows settingsPath, newCols
        End If
        i = 0
        For Each k In newCols.Keys
            i = i + 1
            it = newCols(k)
            If i <= 15 Then review = review & vbLf & "  " & ShortName(it(NC_NAME))
        Next
        For Each k In pending.Keys
            i = i + 1
            If i <= 15 Then review = review & vbLf & "  " & ShortName(pending(k))
        Next
        If i > 15 Then review = review & vbLf & "  ... and " & (i - 15) & " more"
        Cleanup
        If testConfirm = "" Then OpenForReview settingsPath
        RunHash = "STOPPED: nothing was written to inputs-dev." & vbLf & vbLf & _
                  i & " column(s) are marked NEW in the Check column of " & settingsName & ":" & review & vbLf & vbLf & _
                  "For each: check the suggested Action (hash or keep), change it if needed, " & _
                  "then delete the NEW text in Check. Save the file and run again."
        Exit Function
    End If

    ' Pass 2: confirm, then write the hashed copies.
    If BCryptOpenAlgorithmProvider(mAlg, StrPtr("SHA256"), 0, BCRYPT_ALG_HANDLE_HMAC_FLAG) <> 0 Then
        Err.Raise vbObjectError + 1, , "Could not open the Windows SHA-256 provider."
    End If

    If perFile Then
        For Each f In files
            prompt = "Write the hashed copy of" & vbLf & "  " & f & vbLf & "to" & vbLf & "  " & outDir & " ?"
            If fso.FileExists(outDir & "\" & f) Then
                prompt = prompt & vbLf & vbLf & "A file with this name is already there and will be REPLACED."
            End If
            prompt = prompt & vbLf & vbLf & "OK = write it" & vbLf & "Cancel = skip this file"
            If Confirm(prompt, testConfirm) Then
                ProcessFile inDir & "\" & f, outDir & "\" & f, True, newCols, pending, blockers
                written = written & vbLf & "  " & f
                nWritten = nWritten + 1
            Else
                declined = declined & vbLf & "  " & f
            End If
        Next
    Else
        prompt = "Ready to write " & files.Count & " hashed file(s) to" & vbLf & "  " & outDir & vbLf
        i = 0
        For Each f In files
            i = i + 1
            If i <= 15 Then
                prompt = prompt & vbLf & "  " & f & IIf(fso.FileExists(outDir & "\" & f), "   (REPLACES existing copy)", "   (new)")
            End If
        Next
        If files.Count > 15 Then prompt = prompt & vbLf & "  ... and " & (files.Count - 15) & " more"
        prompt = prompt & vbLf & vbLf & "OK = write them" & vbLf & "Cancel = stop (nothing is written)"
        If Confirm(prompt, testConfirm) Then
            For Each f In files
                ProcessFile inDir & "\" & f, outDir & "\" & f, True, newCols, pending, blockers
                written = written & vbLf & "  " & f
                nWritten = nWritten + 1
            Next
        End If
    End If
    Cleanup

    If nWritten = 0 Then
        msg = "STOPPED: nothing was written."
    Else
        msg = "Wrote hashed copies to" & vbLf & "  " & outDir & ":" & written
        If declined <> "" Then msg = msg & vbLf & vbLf & "Not written (you chose Cancel):" & declined
    End If
    If skipped <> "" Then msg = msg & vbLf & vbLf & "Skipped (not csv/xlsx/xls), not copied:" & skipped
    RunHash = msg
    Exit Function

Fail:
    msg = Err.Description
    On Error Resume Next
    If Not mOpenWb Is Nothing Then mOpenWb.Close SaveChanges:=False
    Cleanup
    RunHash = "STOPPED on error: " & msg & vbLf & _
              "Files in the output folder may be incomplete. Fix the problem and run again."
End Function

' OK/Cancel dialog; Cancel is the default button.
Private Function Confirm(ByVal prompt As String, ByVal testConfirm As String) As Boolean
    If testConfirm <> "" Then Confirm = (LCase$(testConfirm) = "yes"): Exit Function
    SetQuiet False
    Confirm = (MsgBox(prompt, vbOKCancel + vbQuestion + vbDefaultButton2, "Confirm") = vbOK)
    SetQuiet True
End Function

Private Function ShortName(ByVal nm As String) As String
    If Len(nm) > 70 Then nm = Left$(nm, 70) & "..."
    ShortName = nm
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
                        newCols As Object, pending As Object, blockers As Object)
    If LCase$(Right$(inPath, 4)) = ".csv" Then
        HashCsv inPath, outPath, doWrite, newCols, pending, blockers
    Else
        HashWorkbook inPath, outPath, doWrite, newCols, pending, blockers
    End If
End Sub

' ---- Column names (hash_columns.xlsx) ----

' Returns "hash", "keep", "pending" (row still marked NEW), or "" if the name
' is not listed. A row matches a header exactly, or as a pattern with * and #.
' First matching row wins.
Private Function ColumnAction(ByVal h As String) As String
    Dim i As Long
    h = LCase$(Trim$(h))
    If h = "" Then Exit Function
    For i = 1 To mNames.Count
        If h = LCase$(mNames(i)) Then ColumnAction = mActions(i): Exit Function
    Next
    For i = 1 To mNames.Count
        If h Like LikePattern(mNames(i)) Then ColumnAction = mActions(i): Exit Function
    Next
End Function

' * and # are wildcards; [ and ? in names are matched literally.
Private Function LikePattern(ByVal nm As String) As String
    nm = LCase$(Trim$(nm))
    nm = Replace(nm, "[", "[[]")
    nm = Replace(nm, "?", "[?]")
    LikePattern = nm
End Function

' Returns "" on success, else a problem description.
Private Function LoadSettings(ByVal path As String) As String
    Dim wb As Workbook, ws As Worksheet, mine As Boolean
    Dim r As Long, lastRow As Long, nm As String, act As String, chk As String, bad As String

    Set mNames = New Collection
    Set mActions = New Collection
    If Len(Dir$(path)) = 0 Then CreateSettings path

    Set wb = OpenWb(path, True, mine)
    Set ws = wb.Worksheets(1)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        nm = "": act = "": chk = ""
        If Not IsError(ws.Cells(r, 1).Value) Then nm = Trim$(CStr(ws.Cells(r, 1).Value))
        If Not IsError(ws.Cells(r, 2).Value) Then act = LCase$(Trim$(CStr(ws.Cells(r, 2).Value)))
        If Not IsError(ws.Cells(r, 4).Value) Then chk = UCase$(Trim$(CStr(ws.Cells(r, 4).Value)))
        If nm <> "" Then
            If Left$(chk, 3) = "NEW" Then
                mNames.Add nm
                mActions.Add "pending"
            ElseIf act = "hash" Or act = "keep" Then
                mNames.Add nm
                mActions.Add act
            Else
                bad = bad & vbLf & "  row " & r & ": " & ShortName(nm) & " -> """ & act & """"
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
    ws.Range("A1:D1").Value = Array("Column name", "Action", "First seen in", "Check")
    ws.Range("A1:D1").Font.Bold = True
    ws.Range("F1").Value = "Action: hash = replace values with a keyed hash (anything that identifies a student); " & _
                           "keep = copy unchanged. Names ignore case; * matches any characters, # one digit. " & _
                           "Rows with NEW in Check are waiting for your review: fix Action, then delete NEW."
    r = 2
    For Each nm In DefaultHashNames()
        ws.Cells(r, 1).Value = nm: ws.Cells(r, 2).Value = "hash": ws.Cells(r, 3).Value = "(default)"
        r = r + 1
    Next
    For Each nm In DefaultKeepNames()
        ws.Cells(r, 1).NumberFormat = "@"
        ws.Cells(r, 1).Value = nm: ws.Cells(r, 2).Value = "keep": ws.Cells(r, 3).Value = "(default)"
        r = r + 1
    Next
    ws.Columns("A:C").AutoFit
    wb.SaveAs Filename:=path, FileFormat:=xlOpenXMLWorkbook
    wb.Close SaveChanges:=False
End Sub

' Adds one row per new column: name, suggested Action, file, and "NEW - why".
Private Sub AppendNewRows(ByVal path As String, newCols As Object)
    Dim wb As Workbook, ws As Worksheet, mine As Boolean, r As Long, k As Variant, it As Variant, sug As Variant
    Set wb = OpenWb(path, False, mine)
    Set ws = wb.Worksheets(1)
    If Trim$(CStr(ws.Range("D1").Value)) = "" Then ws.Range("D1").Value = "Check": ws.Range("D1").Font.Bold = True
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    For Each k In newCols.Keys
        it = newCols(k)
        sug = Suggest(it)
        ws.Cells(r, 1).NumberFormat = "@"
        ws.Cells(r, 1).Value = it(NC_NAME)
        ws.Cells(r, 2).Value = sug(0)
        ws.Cells(r, 3).Value = it(NC_FILE)
        ws.Cells(r, 4).Value = "NEW - " & sug(1)
        r = r + 1
    Next
    wb.Save
    If mine Then wb.Close SaveChanges:=False
End Sub

' Opens the column list for the user and selects the first row marked NEW.
Private Sub OpenForReview(ByVal path As String)
    Dim wb As Workbook, ws As Worksheet, mine As Boolean, r As Long, lastRow As Long
    Set wb = OpenWb(path, False, mine)
    Set ws = wb.Worksheets(1)
    wb.Activate
    ws.Activate
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        If UCase$(Left$(Trim$(CStr(ws.Cells(r, 4).Value)), 3)) = "NEW" Then
            Application.Goto ws.Cells(r, 1), True
            Exit Sub
        End If
    Next
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

' Suggested Action and the reason, from counts only (no values are kept).
Private Function Suggest(it As Variant) As Variant
    Dim n As Long, d As Long, avg As Long, stats As String
    n = it(NC_COUNT)
    d = it(NC_DISTINCT).Count
    If n > 0 Then avg = it(NC_LEN) \ n
    stats = " (" & n & " answers, " & d & " different, average " & avg & " characters)"
    If n = 0 Then
        Suggest = Array("keep", "suggested keep: column is empty")
    ElseIf it(NC_EMAIL) > 0 Then
        Suggest = Array("hash", "suggested hash: contains email addresses" & stats)
    ElseIf it(NC_ID) * 2 >= n Then
        Suggest = Array("hash", "suggested hash: looks like ID numbers" & stats)
    ElseIf it(NC_ID) + it(NC_NUM) = n Then
        Suggest = Array("keep", "suggested keep: numbers only" & stats)
    ElseIf avg > 40 Or it(NC_MAXLEN) > 100 Then
        Suggest = Array("hash", "suggested hash: written answers, may identify students" & stats)
    ElseIf d <= 10 Then
        Suggest = Array("keep", "suggested keep: few different answers, like multiple choice" & stats)
    ElseIf d * 2 <= n And avg <= 30 Then
        Suggest = Array("keep", "suggested keep: answers repeat often" & stats)
    ElseIf d * 10 >= n * 8 Then
        Suggest = Array("hash", "suggested hash: mostly different answers, may identify students" & stats)
    Else
        Suggest = Array("hash", "suggested hash: unsure, please check" & stats)
    End If
End Function

' Records a header that is not in hash_columns.xlsx (once per name, ignoring case).
Private Sub NoteNew(newCols As Object, ByVal h As String, ByVal fname As String)
    If Not newCols.Exists(LCase$(h)) Then
        newCols.Add LCase$(h), Array(h, fname, CreateObject("Scripting.Dictionary"), 0, 0, 0, 0, 0, 0)
    End If
End Sub

' Updates a new column's counts with one value. Only counts are kept, plus
' the distinct values in memory for counting; nothing is written anywhere.
Private Sub AddStat(newCols As Object, ByVal h As String, ByVal v As String)
    Dim it As Variant, d As Object
    v = Trim$(v)
    If v = "" Or Not newCols.Exists(LCase$(h)) Then Exit Sub
    it = newCols(LCase$(h))
    Set d = it(NC_DISTINCT)
    If Not d.Exists(v) And d.Count < 10000 Then d.Add v, True
    it(NC_COUNT) = it(NC_COUNT) + 1
    If v Like "*?@?*.?*" And InStr(v, " ") = 0 Then
        it(NC_EMAIL) = it(NC_EMAIL) + 1
    ElseIf Len(v) >= 6 And Not v Like "*[!0-9]*" Then
        it(NC_ID) = it(NC_ID) + 1
    ElseIf IsNumeric(v) Then
        it(NC_NUM) = it(NC_NUM) + 1
    End If
    it(NC_LEN) = it(NC_LEN) + Len(v)
    If Len(v) > it(NC_MAXLEN) Then it(NC_MAXLEN) = Len(v)
    newCols(LCase$(h)) = it
End Sub

Private Sub AddPending(pending As Object, ByVal h As String)
    If Not pending.Exists(LCase$(h)) Then pending.Add LCase$(h), h
End Sub

Private Sub AddBlocker(blockers As Object, ByVal s As String)
    If Not blockers.Exists(s) Then blockers.Add s, True
End Sub

' ---- Excel workbooks ----

Private Sub HashWorkbook(ByVal inPath As String, ByVal outPath As String, ByVal doWrite As Boolean, _
                         newCols As Object, pending As Object, blockers As Object)
    Dim ws As Worksheet, fname As String
    fname = Mid$(inPath, InStrRev(inPath, "\") + 1)
    Set mOpenWb = Workbooks.Open(Filename:=inPath, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
    For Each ws In mOpenWb.Worksheets
        HashSheet ws, fname, doWrite, newCols, pending, blockers
    Next
    If doWrite Then mOpenWb.SaveAs Filename:=outPath, FileFormat:=mOpenWb.FileFormat
    mOpenWb.Close SaveChanges:=False
    Set mOpenWb = Nothing
End Sub

Private Sub HashSheet(ws As Worksheet, ByVal fname As String, ByVal doWrite As Boolean, _
                      newCols As Object, pending As Object, blockers As Object)
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

        If act = "pending" Then
            AddPending pending, h
        ElseIf act = "" Then
            If h <> "" Then
                NoteNew newCols, h, fname
                If Not IsEmpty(vals) Then
                    For r = 1 To UBound(vals, 1)
                        If Not IsError(vals(r, 1)) Then AddStat newCols, h, CStr(vals(r, 1))
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
                    newCols As Object, pending As Object, blockers As Object)
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
                If acts(col) = "pending" Then
                    AddPending pending, heads(col)
                ElseIf acts(col) = "" And heads(col) <> "" Then
                    NoteNew newCols, heads(col), fname
                End If
            Else
                v = Trim$(Unquote(raw))
                act = "": hname = ""
                If col <= nHead Then act = acts(col): hname = heads(col)
                If act = "" Then
                    If hname <> "" Then
                        AddStat newCols, hname, v
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
