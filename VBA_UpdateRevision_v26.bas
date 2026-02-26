' ============================================================
' MACRO: Full revision update pipeline (v26)
' ============================================================
' Alt+F8 -> UpdateRevision -> Run
'
' Stage 1: Rename all folders & files (C02->C03), delete pdf/bak
' Stage 2: Process docx/dwg content in each subfolder
' Stage 3: Export PDF to PDF_ASSEMBLY
'
' v26 fixes (from v25):
'   - B-020 FIX: PERM stamp date (10.06.24) != doc date -> wildcard DD.MM.YY replace
'   - B-021 FIX: UL/ALL DOCX stamp dates not replaced -> wildcard for all story ranges
' v25 fixes (from v24):
'   - B-014 FIX: PERM auth date in "ot DD.MM.YYYY" header not replaced
'   - B-016 FIX: GIP surname-only search via LastWord helper
'   - B-017 FIX: Logo image1.jpg -> SVG with updated year (ZIP manipulation)
'   - B-018 FIX: PDF export from DWG via PlotToFile
'   - B-019 FIX: Skip Model space when CLD layouts exist (10x speedup)
' v24 (from v23):
'   - B-012 FIX: Cyrillic C in DWG Rev field (Chr(1057) vs Chr(67))
'   - B-009 improved: prevInv also scans ACDBTEXT/ACDBMTEXT entities
'   - B-013 FIX: Nested block refs — process inner block definition entities
' v23 (from v22):
'   - DWG hang fix: acadApp.Visible=True + suppress PROXYNOTICE/FILEDIA
'   - GIP fix: surname-only search + footer table iteration in ProcessDocx
'   - PERM body text: non-table paragraphs also processed for GIP/dates
' v22 (from v17 base):
'   - FindDocxFolders: also detects .dwg files (T-010)
'   - InputBox removed: OLD_FIO="" when header empty (T-006/v18)
'   - Word restart on Err 80010108 (T-007/v19)
'   - ProcessDwg: AutoCAD COM for DWG stamps (T-011..T-015)
'   - B-009 FIX: prevInv auto-detect scans ALL spaces incl Model (T-017)
' ============================================================

Const PDF_FOLDER_NAME As String = "PDF_ASSEMBLY"

' Column numbers in xlsx
Const COL_SHIFR As Long = 4      ' D
Const COL_OLD_REV As Long = 6    ' F
Const COL_OLD_INV As Long = 7    ' G
Const COL_OLD_DATE As Long = 8   ' H
Const COL_GIP As Long = 10       ' J
Const COL_OLD_PERM As Long = 11  ' K
Const COL_NEW_INV As Long = 12   ' L
Const COL_DATE As Long = 13      ' M
Const COL_NEW_PERM As Long = 14  ' N
Const COL_OLD_GIP As Long = 10   ' J - old GIP is in same column, row 3 (header row)
Const DATA_START_ROW As Long = 4
Const HEADER_ROW As Long = 3     ' Row with column headers (contains old GIP name)

Public TRACE_LOG_PATH As String

Sub AppendTrace(msg As String)
    If TRACE_LOG_PATH = "" Then Exit Sub
    On Error Resume Next
    Dim ff As Long
    ff = FreeFile
    Open TRACE_LOG_PATH For Append As #ff
    Print #ff, Format(Now, "yyyy-mm-dd hh:nn:ss") & " | " & msg
    Close #ff
    On Error GoTo 0
End Sub


' ============================================================
' UTILITY FUNCTIONS
' ============================================================

Function Transliterate(rusText As String) As String
    Dim result As String
    Dim i As Long
    Dim code As Long
    result = ""
    For i = 1 To Len(rusText)
        code = AscW(Mid(rusText, i, 1))
        Select Case code
            Case 1040: result = result & "A"
            Case 1041: result = result & "B"
            Case 1042: result = result & "V"
            Case 1043: result = result & "G"
            Case 1044: result = result & "D"
            Case 1045: result = result & "E"
            Case 1025: result = result & "E"
            Case 1046: result = result & "Zh"
            Case 1047: result = result & "Z"
            Case 1048: result = result & "I"
            Case 1049: result = result & "Y"
            Case 1050: result = result & "K"
            Case 1051: result = result & "L"
            Case 1052: result = result & "M"
            Case 1053: result = result & "N"
            Case 1054: result = result & "O"
            Case 1055: result = result & "P"
            Case 1056: result = result & "R"
            Case 1057: result = result & "S"
            Case 1058: result = result & "T"
            Case 1059: result = result & "U"
            Case 1060: result = result & "F"
            Case 1061: result = result & "Kh"
            Case 1062: result = result & "Ts"
            Case 1063: result = result & "Ch"
            Case 1064: result = result & "Sh"
            Case 1065: result = result & "Shch"
            Case 1066
            Case 1067: result = result & "Y"
            Case 1068
            Case 1069: result = result & "E"
            Case 1070: result = result & "Yu"
            Case 1071: result = result & "Ya"
            Case 1072: result = result & "a"
            Case 1073: result = result & "b"
            Case 1074: result = result & "v"
            Case 1075: result = result & "g"
            Case 1076: result = result & "d"
            Case 1077: result = result & "e"
            Case 1105: result = result & "e"
            Case 1078: result = result & "zh"
            Case 1079: result = result & "z"
            Case 1080: result = result & "i"
            Case 1081: result = result & "y"
            Case 1082: result = result & "k"
            Case 1083: result = result & "l"
            Case 1084: result = result & "m"
            Case 1085: result = result & "n"
            Case 1086: result = result & "o"
            Case 1087: result = result & "p"
            Case 1088: result = result & "r"
            Case 1089: result = result & "s"
            Case 1090: result = result & "t"
            Case 1091: result = result & "u"
            Case 1092: result = result & "f"
            Case 1093: result = result & "kh"
            Case 1094: result = result & "ts"
            Case 1095: result = result & "ch"
            Case 1096: result = result & "sh"
            Case 1097: result = result & "shch"
            Case 1098
            Case 1099: result = result & "y"
            Case 1100
            Case 1101: result = result & "e"
            Case 1102: result = result & "yu"
            Case 1103: result = result & "ya"
            Case Else: result = result & Mid(rusText, i, 1)
        End Select
    Next i
    Transliterate = result
End Function

Function GetParentFolder(folderPath As String) As String
    Dim p As String
    p = folderPath
    If Right(p, 1) = "\" Then p = Left(p, Len(p) - 1)
    Dim pos As Long
    pos = InStrRev(p, "\")
    If pos > 0 Then
        GetParentFolder = Left(p, pos)
    Else
        GetParentFolder = p & "\"
    End If
End Function

Function FormatDateDMY(dt As Date) As String
    FormatDateDMY = Format(dt, "DD.MM.YYYY")
End Function

Function FormatDateShort(dt As Date) As String
    FormatDateShort = Format(dt, "DD.MM.YY")
End Function

Function FormatDateFolder(dt As Date) As String
    FormatDateFolder = Format(dt, "DD.MM.YYYY")
End Function

Function NormalizeRev(rev As String) As String
    Dim s As String
    s = Trim(rev)
    s = Replace(s, ChrW(1057), "C")
    s = Replace(s, ChrW(1089), "C")
    NormalizeRev = UCase(s)
End Function

' v24: Check if string contains only digit characters (0-9)
Function IsAllDigits(s As String) As Boolean
    Dim i As Long
    If Len(s) = 0 Then IsAllDigits = False: Exit Function
    For i = 1 To Len(s)
        If Mid(s, i, 1) < "0" Or Mid(s, i, 1) > "9" Then
            IsAllDigits = False: Exit Function
        End If
    Next i
    IsAllDigits = True
End Function

' v25 B-016: Extract last word (surname) from a full name string
Function LastWord(s As String) As String
    Dim parts() As String
    Dim t As String
    t = Trim(s)
    If t = "" Then
        LastWord = ""
        Exit Function
    End If
    parts = Split(t, " ")
    LastWord = parts(UBound(parts))
End Function

' v25 B-014: Find first DD.MM.YYYY date pattern in text string
' Returns the 10-char date string, or "" if not found
Function FindDateDDMMYYYY(text As String) As String
    Dim i As Long
    If Len(text) < 10 Then
        FindDateDDMMYYYY = ""
        Exit Function
    End If
    For i = 1 To Len(text) - 9
        If Mid(text, i + 2, 1) = "." And Mid(text, i + 5, 1) = "." Then
            Dim dd As String, mm As String, yyyy As String
            dd = Mid(text, i, 2)
            mm = Mid(text, i + 3, 2)
            yyyy = Mid(text, i + 6, 4)
            If IsNumeric(dd) And IsNumeric(mm) And IsNumeric(yyyy) Then
                If CLng(dd) >= 1 And CLng(dd) <= 31 And CLng(mm) >= 1 And CLng(mm) <= 12 Then
                    FindDateDDMMYYYY = Mid(text, i, 10)
                    Exit Function
                End If
            End If
        End If
    Next i
    FindDateDDMMYYYY = ""
End Function

Function DetectNewRevFromHeader(ws As Object) As String
    Dim hdr As String
    hdr = Trim(CStr(ws.Cells(2, COL_NEW_PERM).Value & ""))
    hdr = NormalizeRev(hdr)
    Dim i As Long
    For i = 1 To Len(hdr) - 2
        If Mid(hdr, i, 1) = "C" Then
            Dim num As String
            num = Mid(hdr, i + 1, 2)
            If IsNumeric(num) Then
                DetectNewRevFromHeader = "C" & num
                Exit Function
            End If
        End If
    Next i
    DetectNewRevFromHeader = ""
End Function


' Validate strict date token format DD.MM.YY or DD.MM.YYYY
Function IsValidDateToken(s As String) As Boolean
    On Error GoTo Invalid
    Dim dd As Long
    Dim mm As Long
    Dim yy As Long

    s = Trim(s)
    If Len(s) <> 8 And Len(s) <> 10 Then GoTo Invalid
    If Mid(s, 3, 1) <> "." Or Mid(s, 6, 1) <> "." Then GoTo Invalid

    dd = CLng(Mid(s, 1, 2))
    mm = CLng(Mid(s, 4, 2))
    yy = CLng(Mid(s, 7))

    If dd < 1 Or dd > 31 Then GoTo Invalid
    If mm < 1 Or mm > 12 Then GoTo Invalid

    If Len(s) = 8 Then
        If yy < 0 Or yy > 99 Then GoTo Invalid
    Else
        If yy < 1900 Or yy > 2099 Then GoTo Invalid
    End If

    IsValidDateToken = True
    Exit Function
Invalid:
    IsValidDateToken = False
End Function

' Replace dates by wildcard, but only if found token is a valid date
Function ReplaceValidatedDateWildcard(targetRange As Object, patternText As String, replacementText As String) As Long
    Dim cnt As Long
    cnt = 0

    Dim findRange As Object
    Set findRange = targetRange.Duplicate

    With findRange.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = patternText
        .Forward = True
        .Wrap = 0
        .Format = False
        .MatchCase = True
        .MatchWildcards = True

        Do While .Execute
            If IsValidDateToken(findRange.Text) Then
                findRange.Text = replacementText
                cnt = cnt + 1
            End If
            findRange.Collapse 0
        Loop

        .MatchWildcards = False
    End With

    ReplaceValidatedDateWildcard = cnt
End Function


' ============================================================
' STAGE 1: Recursive rename & cleanup
' ============================================================

' Delete generated/temporary files recursively (except ARCHIVE/PDF_ASSEMBLY trees)
Sub DeleteFilesRecursive(fso As Object, folderPath As String, _
                          oldRev As String, ByRef delCount As Long, ByRef logText As String)
    Dim folder As Object
    Set folder = fso.GetFolder(folderPath)
    
    Dim folderNameUpper As String
    folderNameUpper = UCase(folder.Name)
    If folderNameUpper = "ARCHIVE" Or folderNameUpper = PDF_FOLDER_NAME Then Exit Sub
    
    Dim f As Object
    For Each f In folder.Files
        Dim fName As String
        Dim ext As String
        Dim shouldDelete As Boolean
        
        fName = f.Name
        ext = LCase(fso.GetExtensionName(fName))
        shouldDelete = False
        
        If Left(fName, 2) = "~$" Then
            shouldDelete = True
        ElseIf fName = "hardcopy.log" Or fName = "acad.err" Then
            shouldDelete = True
        ElseIf ext = "pdf" Or ext = "bak" Or ext = "dwl" Or ext = "dwl2" Then
            shouldDelete = True
        End If
        
        If shouldDelete Then
            On Error Resume Next
            f.Delete True
            If Err.Number = 0 Then
                logText = logText & "  DEL: " & fName & vbCrLf
                delCount = delCount + 1
            Else
                logText = logText & "  DEL ERR: " & fName & " Err" & Err.Number & " " & Err.Description & vbCrLf
                Err.Clear
            End If
            On Error GoTo 0
        End If
    Next f
    
    Dim sub_ As Object
    For Each sub_ In folder.SubFolders
        If UCase(sub_.Name) <> "ARCHIVE" And UCase(sub_.Name) <> PDF_FOLDER_NAME Then
            DeleteFilesRecursive fso, sub_.Path, oldRev, delCount, logText
        End If
    Next sub_
End Sub


' Rename files in a single folder (not recursive - called per folder)
Sub RenameFilesInFolder(fso As Object, folderPath As String, _
                         oldRev As String, newRev As String, _
                         ByRef renCount As Long, ByRef logText As String)
    
    If Not fso.FolderExists(folderPath) Then Exit Sub
    
    Dim folder As Object
    Set folder = fso.GetFolder(folderPath)
    
    ' Collect ALL file names first, then filter
    Dim allNames() As String
    Dim allCnt As Long
    allCnt = 0
    Dim f As Object
    For Each f In folder.Files
        allCnt = allCnt + 1
        ReDim Preserve allNames(1 To allCnt)
        allNames(allCnt) = f.Name
    Next f
    
    ' Filter for files containing oldRev (Latin C or Cyrillic C)
    Dim names() As String
    Dim cnt As Long
    cnt = 0
    
    ' Also build Cyrillic version: replace Latin C with Cyrillic S (U+0421)
    Dim cyrOldRev As String
    cyrOldRev = Replace(oldRev, "C", ChrW(1057))
    
    Dim k As Long
    For k = 1 To allCnt
        If InStr(1, allNames(k), oldRev, vbBinaryCompare) > 0 Or _
           InStr(1, allNames(k), cyrOldRev, vbBinaryCompare) > 0 Then
            cnt = cnt + 1
            ReDim Preserve names(1 To cnt)
            names(cnt) = allNames(k)
        End If
    Next k
    
    logText = logText & "    [FILES in " & folder.Name & "] total=" & allCnt & " match='" & oldRev & "'=" & cnt & vbCrLf
    
    ' Build Cyrillic new revision too
    Dim cyrNewRev As String
    cyrNewRev = Replace(newRev, "C", ChrW(1057))
    
    ' Rename from collected array using VBA Name statement
    Dim i As Long
    For i = 1 To cnt
        Dim oldName As String
        Dim newName As String
        oldName = names(i)
        ' Replace both Latin and Cyrillic versions
        newName = Replace(oldName, oldRev, newRev, , , vbBinaryCompare)
        newName = Replace(newName, cyrOldRev, newRev, , , vbBinaryCompare)
        
        If oldName <> newName Then
            Dim oldFull As String
            Dim newFull As String
            oldFull = folderPath & "\" & oldName
            newFull = folderPath & "\" & newName
            
            If Not fso.FileExists(newFull) Then
                On Error Resume Next
                ' Use VBA Name statement instead of FSO (more reliable)
                Name oldFull As newFull
                If Err.Number = 0 Then
                    logText = logText & "  FILE: " & oldName & " -> " & newName & vbCrLf
                    renCount = renCount + 1
                Else
                    logText = logText & "  FILE ERR: " & oldName & " Err" & Err.Number & " " & Err.Description & vbCrLf
                    Err.Clear
                End If
                On Error GoTo 0
            Else
                logText = logText & "  FILE SKIP(exists): " & newName & vbCrLf
            End If
        End If
    Next i
End Sub


' Rename subfolders recursively (bottom-up: deepest first)
Sub RenameFoldersRecursive(fso As Object, folderPath As String, _
                            oldRev As String, newRev As String, _
                            ByRef renCount As Long, ByRef logText As String)
    
    If Not fso.FolderExists(folderPath) Then Exit Sub
    
    Dim folder As Object
    Set folder = fso.GetFolder(folderPath)
    
    logText = logText & "  [SCAN] " & folder.Name & " (" & folder.SubFolders.Count & " subdirs, " & folder.Files.Count & " files)" & vbCrLf
    
    ' First: recurse into subfolders (process children first)
    Dim subPaths() As String
    Dim subCnt As Long
    subCnt = 0
    Dim sub_ As Object
    For Each sub_ In folder.SubFolders
        subCnt = subCnt + 1
        ReDim Preserve subPaths(1 To subCnt)
        subPaths(subCnt) = sub_.Path
    Next sub_
    
    Dim i As Long
    For i = 1 To subCnt
        RenameFoldersRecursive fso, subPaths(i), oldRev, newRev, renCount, logText
    Next i
    
    ' Rename files in this folder
    RenameFilesInFolder fso, folderPath, oldRev, newRev, renCount, logText
    
    ' Then: rename subfolders of current folder (re-read after recursion)
    Set folder = fso.GetFolder(folderPath)
    Dim subNames() As String
    
    ' Build Cyrillic version of oldRev
    Dim cyrOldRev2 As String
    cyrOldRev2 = Replace(oldRev, "C", ChrW(1057))
    
    subCnt = 0
    For Each sub_ In folder.SubFolders
        If UCase(sub_.Name) <> "ARCHIVE" And UCase(sub_.Name) <> "PDF_ASSEMBLY" Then
            If InStr(1, sub_.Name, "_" & oldRev, vbBinaryCompare) > 0 Or _
               InStr(1, sub_.Name, "_" & cyrOldRev2, vbBinaryCompare) > 0 Then
                subCnt = subCnt + 1
                ReDim Preserve subNames(1 To subCnt)
                subNames(subCnt) = sub_.Name
            End If
        End If
    Next sub_
    
    Dim cyrNewRev2 As String
    cyrNewRev2 = Replace(newRev, "C", ChrW(1057))
    
    For i = 1 To subCnt
        Dim oldSub As String
        Dim newSub As String
        oldSub = subNames(i)
        newSub = Replace(oldSub, "_" & oldRev, "_" & newRev, , , vbBinaryCompare)
        newSub = Replace(newSub, "_" & cyrOldRev2, "_" & newRev, , , vbBinaryCompare)
        
        Dim oldSubPath As String
        Dim newSubPath As String
        oldSubPath = folderPath & "\" & oldSub
        newSubPath = folderPath & "\" & newSub
        
        If oldSub <> newSub Then
            If Not fso.FolderExists(newSubPath) Then
                On Error Resume Next
                fso.MoveFolder oldSubPath, newSubPath
                If Err.Number = 0 Then
                    logText = logText & "  DIR: " & oldSub & " -> " & newSub & vbCrLf
                    renCount = renCount + 1
                Else
                    logText = logText & "  DIR ERR: " & oldSub & " Err" & Err.Number & " " & Err.Description & vbCrLf
                    Err.Clear
                End If
                On Error GoTo 0
            End If
        End If
    Next i
End Sub


' Find all folders containing docx files (for Stage 2)
Sub FindDocxFolders(fso As Object, folderPath As String, _
                     ByRef folders As Collection)
    Dim folder As Object
    Set folder = fso.GetFolder(folderPath)
    
    ' Check if this folder has docx or dwg files
    Dim f As Object
    For Each f In folder.Files
        Dim extLower As String
        extLower = LCase(fso.GetExtensionName(f.Name))
        If extLower = "docx" Or extLower = "doc" Or extLower = "dwg" Then
            folders.Add folderPath
            Exit For
        End If
    Next f
    
    ' Recurse
    Dim sub_ As Object
    For Each sub_ In folder.SubFolders
        ' Skip ARCHIVE and PDF_ASSEMBLY
        If UCase(sub_.Name) <> "ARCHIVE" And UCase(sub_.Name) <> PDF_FOLDER_NAME Then
            FindDocxFolders fso, sub_.Path, folders
        End If
    Next sub_
End Sub


' ============================================================
' v25 B-017: Replace logo image in DOCX (JPG -> SVG with new year)
' Uses Shell.Application for ZIP operations (Shell() is forbidden)
' ============================================================

' Read file content as raw bytes into a string (for ASCII/Latin XML files)
Private Function ReadFileBytes(filePath As String) As String
    Dim ff As Long
    ff = FreeFile
    Dim fLen As Long
    Open filePath For Binary Access Read As #ff
    fLen = LOF(ff)
    If fLen > 0 Then
        Dim buf As String
        buf = Space(fLen)
        Get #ff, , buf
        ReadFileBytes = buf
    Else
        ReadFileBytes = ""
    End If
    Close #ff
End Function

' Write string to file as raw bytes
Private Sub WriteFileBytes(filePath As String, content As String)
    Dim ff As Long
    ff = FreeFile
    ' Delete existing file first
    On Error Resume Next
    Kill filePath
    Err.Clear
    On Error GoTo 0
    Open filePath For Binary Access Write As #ff
    Put #ff, , content
    Close #ff
End Sub

' Write UTF-8 file using ADODB.Stream (for SVG with Cyrillic)
Private Sub WriteFileUtf8(filePath As String, content As String)
    Dim utfStream As Object
    Set utfStream = CreateObject("ADODB.Stream")
    utfStream.Type = 2  ' adTypeText
    utfStream.Charset = "UTF-8"
    utfStream.Open
    utfStream.WriteText content
    utfStream.SaveToFile filePath, 2  ' adSaveCreateOverWrite
    utfStream.Close
    Set utfStream = Nothing
End Sub

' Replace logo image1.jpg with SVG containing updated year in DOCX file
' filePath = already-saved DOCX file (must NOT be open in Word)
' newYear = year string like "2026"
' Returns True on success
Function ReplaceLogoInDocx(filePath As String, newYear As String) As Boolean
    On Error GoTo LogoErr
    ReplaceLogoInDocx = False

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FileExists(filePath) Then Exit Function

    ' Create temp working directory
    Dim tempBase As String
    tempBase = fso.GetSpecialFolder(2).Path & "\" & fso.GetTempName()
    fso.CreateFolder tempBase

    ' Copy docx -> temp zip
    Dim tempZip As String
    tempZip = tempBase & "\work.zip"
    fso.CopyFile filePath, tempZip

    ' Create extraction folder
    Dim extractDir As String
    extractDir = tempBase & "\ex\"
    fso.CreateFolder extractDir

    ' Extract ZIP using Shell.Application COM
    Dim oShell As Object
    Set oShell = CreateObject("Shell.Application")
    Dim zipNs As Object
    Set zipNs = oShell.Namespace(tempZip)
    Dim extNs As Object
    Set extNs = oShell.Namespace(extractDir)

    If zipNs Is Nothing Or extNs Is Nothing Then GoTo LogoCleanup

    ' CopyHere: 4=no progress, 16=yes to all, 256=no UI, 1024=no confirm new dir
    extNs.CopyHere zipNs.Items, 4 + 16 + 256 + 1024

    ' Wait for async extraction to complete
    Dim waitCnt As Long
    waitCnt = 0
    Do
        Application.Wait Now + TimeValue("00:00:01")
        DoEvents
        waitCnt = waitCnt + 1
        If fso.FolderExists(extractDir & "word") Then Exit Do
    Loop While waitCnt < 30

    ' Extra settle time
    Application.Wait Now + TimeValue("00:00:02")

    ' --- Delete old image ---
    Dim mediaDir As String
    mediaDir = extractDir & "word\media\"
    If fso.FileExists(mediaDir & "image1.jpg") Then
        fso.DeleteFile mediaDir & "image1.jpg", True
    End If
    If fso.FileExists(mediaDir & "image1.jpeg") Then
        fso.DeleteFile mediaDir & "image1.jpeg", True
    End If

    ' --- Create SVG logo ---
    Dim q As String
    q = Chr(34)
    Dim svgText As String
    ' НЭПТ = ChrW(1053) & ChrW(1069) & ChrW(1055) & ChrW(1058)
    Dim neptCyr As String
    neptCyr = ChrW(1053) & ChrW(1069) & ChrW(1055) & ChrW(1058)

    svgText = "<?xml version=" & q & "1.0" & q & " encoding=" & q & "UTF-8" & q & "?>" & vbLf
    svgText = svgText & "<svg xmlns=" & q & "http://www.w3.org/2000/svg" & q & " viewBox=" & q & "0 0 490 188" & q & ">" & vbLf
    svgText = svgText & "  <g transform=" & q & "translate(90,94)" & q & " fill=" & q & "none" & q & " stroke=" & q & "#1a6b2e" & q & " stroke-width=" & q & "4" & q & ">" & vbLf
    svgText = svgText & "    <circle cx=" & q & "0" & q & " cy=" & q & "0" & q & " r=" & q & "85" & q & "/>" & vbLf
    svgText = svgText & "    <ellipse rx=" & q & "85" & q & " ry=" & q & "28" & q & " transform=" & q & "rotate(0)" & q & "/>" & vbLf
    svgText = svgText & "    <ellipse rx=" & q & "85" & q & " ry=" & q & "28" & q & " transform=" & q & "rotate(60)" & q & "/>" & vbLf
    svgText = svgText & "    <ellipse rx=" & q & "85" & q & " ry=" & q & "28" & q & " transform=" & q & "rotate(120)" & q & "/>" & vbLf
    svgText = svgText & "    <circle cx=" & q & "0" & q & " cy=" & q & "0" & q & " r=" & q & "26" & q & "/>" & vbLf
    svgText = svgText & "  </g>" & vbLf
    svgText = svgText & "  <text x=" & q & "90" & q & " y=" & q & "100" & q & " font-family=" & q & "Arial" & q
    svgText = svgText & " font-size=" & q & "17" & q & " font-weight=" & q & "bold" & q
    svgText = svgText & " fill=" & q & "#1a6b2e" & q & " text-anchor=" & q & "middle" & q & ">"
    svgText = svgText & neptCyr & "</text>" & vbLf
    svgText = svgText & "  <text x=" & q & "295" & q & " y=" & q & "75" & q & " font-family=" & q & "Arial" & q
    svgText = svgText & " font-size=" & q & "34" & q & " font-weight=" & q & "bold" & q
    svgText = svgText & " fill=" & q & "#000000" & q & " text-anchor=" & q & "middle" & q & ">"
    svgText = svgText & "JSC &#x22;NEPT&#x22;</text>" & vbLf
    svgText = svgText & "  <text x=" & q & "295" & q & " y=" & q & "125" & q & " font-family=" & q & "Arial" & q
    svgText = svgText & " font-size=" & q & "34" & q & " font-weight=" & q & "bold" & q
    svgText = svgText & " fill=" & q & "#000000" & q & " text-anchor=" & q & "middle" & q & ">"
    svgText = svgText & "Moscow " & newYear & "</text>" & vbLf
    svgText = svgText & "</svg>"

    WriteFileUtf8 mediaDir & "image1.svg", svgText

    ' --- Update footer rels: image1.jpg -> image1.svg ---
    Dim relsNames As Variant
    relsNames = Array("word\_rels\footer1.xml.rels", "word\_rels\footer2.xml.rels")
    Dim ri As Long
    For ri = 0 To UBound(relsNames)
        Dim relsPath As String
        relsPath = extractDir & relsNames(ri)
        If fso.FileExists(relsPath) Then
            Dim relsContent As String
            relsContent = ReadFileBytes(relsPath)
            relsContent = Replace(relsContent, "media/image1.jpg", "media/image1.svg")
            relsContent = Replace(relsContent, "media/image1.jpeg", "media/image1.svg")
            WriteFileBytes relsPath, relsContent
        End If
    Next ri

    ' --- Update [Content_Types].xml: add SVG extension ---
    Dim ctPath As String
    ctPath = extractDir & "[Content_Types].xml"
    If fso.FileExists(ctPath) Then
        Dim ctContent As String
        ctContent = ReadFileBytes(ctPath)
        ' Add SVG content type if not present
        If InStr(1, ctContent, "Extension=" & q & "svg" & q, vbTextCompare) = 0 Then
            ctContent = Replace(ctContent, "</Types>", _
                "<Default Extension=" & q & "svg" & q & " ContentType=" & q & "image/svg+xml" & q & "/></Types>")
        End If
        WriteFileBytes ctPath, ctContent
    End If

    ' --- Repack into new ZIP ---
    Dim newZip As String
    newZip = tempBase & "\result.zip"

    ' Create empty ZIP (End of Central Directory record = 22 bytes)
    Dim ffz As Long
    ffz = FreeFile
    Open newZip For Binary Access Write As #ffz
    Dim zipHdr(0 To 21) As Byte
    zipHdr(0) = 80: zipHdr(1) = 75: zipHdr(2) = 5: zipHdr(3) = 6
    ' bytes 4-21 are already 0
    Put #ffz, 1, zipHdr
    Close #ffz

    ' Copy extracted (modified) files into new ZIP
    Dim newZipNs As Object
    Set newZipNs = oShell.Namespace(newZip)
    Dim srcNs As Object
    Set srcNs = oShell.Namespace(extractDir)

    If newZipNs Is Nothing Or srcNs Is Nothing Then GoTo LogoCleanup

    newZipNs.CopyHere srcNs.Items, 4 + 16 + 256 + 1024

    ' Wait for async packing
    waitCnt = 0
    Do
        Application.Wait Now + TimeValue("00:00:01")
        DoEvents
        waitCnt = waitCnt + 1
        If fso.GetFile(newZip).Size > 1000 Then Exit Do
    Loop While waitCnt < 60

    ' Extra settle time
    Application.Wait Now + TimeValue("00:00:03")

    ' Replace original DOCX with new ZIP
    fso.CopyFile newZip, filePath, True

    ReplaceLogoInDocx = True

LogoCleanup:
    On Error Resume Next
    If fso.FolderExists(tempBase) Then fso.DeleteFolder tempBase, True
    Set fso = Nothing
    Set oShell = Nothing
    Exit Function

LogoErr:
    On Error Resume Next
    If Not fso Is Nothing Then
        If tempBase <> "" Then
            If fso.FolderExists(tempBase) Then fso.DeleteFolder tempBase, True
        End If
    End If
    Set fso = Nothing
    Set oShell = Nothing
    ReplaceLogoInDocx = False
End Function


' ============================================================
' PERM FILE PROCESSING via direct XML manipulation (no Word Find.Replace)
' This avoids table formatting corruption caused by wdReplaceAll
' on multi-run text nodes (dates split across multiple w:t elements)
' ============================================================

' Helper: replace text that may be SPLIT ACROSS MULTIPLE RUNS in a paragraph.
' Iterates all paragraphs in a given range (cell/story), finds split text,
' puts new text in first run, clears contribution from other runs.
' PRESERVES all run formatting (rPr) - only changes w:t text content.
' Returns count of replacements made.
' Replace text within a cell range using Word Find on the cell range only.
' Using cell-scoped Find does NOT merge runs (unlike doc.Content.Find).
' Returns count of replacements made.
Function ReplaceInCellRange(cellRng As Object, oldText As String, newText As String) As Long
    If oldText = "" Or oldText = newText Then
        ReplaceInCellRange = 0
        Exit Function
    End If
    
    Dim count As Long
    count = 0
    
    On Error Resume Next
    With cellRng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = oldText
        .Replacement.Text = newText
        .Forward = True
        .Wrap = 0           ' wdFindStop = 0: stop at end, don't wrap outside cell
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .Execute Replace:=2  ' wdReplaceAll
    End With
    If Err.Number = 0 Then count = 1  ' wdReplaceAll doesn't return count; use 1 as flag
    Err.Clear
    On Error GoTo 0
    
    ReplaceInCellRange = count
End Function


' Build combined text of all runs in a paragraph via Characters collection.
' Word OM: para.Range.Characters iterates individual characters.
' We use this to detect if a paragraph contains the target text.
' Then use Find on the cell to do the actual replace (preserves run formatting).
' Returns count of replacements across all cells.
Function ReplaceInRuns(rng As Object, oldText As String, newText As String) As Long
    ' This wrapper is called with a cell range (rng already trimmed by -1 char)
    ' Delegate to cell-scoped Find which is safe and doesn't merge runs
    ReplaceInRuns = ReplaceInCellRange(rng, oldText, newText)
End Function


' Process PERM file via pure Word object model (run-by-run replacement)
' NO ZIP, NO PowerShell, NO Shell.Application - works in ALL environments
' Preserves table formatting by replacing only run text, not formatting
' Returns: number of changes made, or -1 on error
Function ProcessPermXml(filePath As String, _
                        oldRev As String, newRev As String, _
                        oldYear As String, newYear As String, _
                        oldFioRus As String, newFioRus As String, _
                        oldFioEng As String, newFioEng As String, _
                        oldInv As String, newInv As String, _
                        oldPerm As String, newPerm As String, _
                        oldDateShort As String, newDateShort As String, _
                        oldDateFull_DDMMYYYY As String, newDateFull As String, _
                        wordApp As Object, pdfFolder As String, _
                        ByRef pdfOk As Boolean, ByRef pdfErrMsg As String, _
                        ByRef newPermPath As String) As Long
    
    On Error GoTo ErrHandler
    
    pdfOk = False
    pdfErrMsg = ""
    newPermPath = ""
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim docFolder As String
    docFolder = Left(filePath, InStrRev(filePath, "\"))
    
    ' Open the PERM document
    Dim doc As Object
    Set doc = wordApp.Documents.Open(filePath, ReadOnly:=False)
    
    Dim totalChanges As Long
    totalChanges = 0
    
    ' ---- Build replacement strings ----
    Dim cyrOldRev As String
    Dim cyrNewRev As String
    cyrOldRev = ChrW(1057) & Mid(oldRev, 2)   ' Cyrillic С + number
    cyrNewRev = ChrW(1057) & Mid(newRev, 2)
    
    ' v25 B-016: Use LastWord for surname extraction
    Dim oldSurRus As String: oldSurRus = ""
    Dim newSurRus As String: newSurRus = ""
    If oldFioRus <> "" And newFioRus <> "" Then
        oldSurRus = LastWord(oldFioRus)
        newSurRus = LastWord(newFioRus)
    End If

    Dim oldSurEng As String: oldSurEng = ""
    Dim newSurEng As String: newSurEng = ""
    If oldFioEng <> "" And newFioEng <> "" Then
        oldSurEng = LastWord(oldFioEng)
        newSurEng = LastWord(newFioEng)
    End If
    
    ' ---- Process all tables via ReplaceInRuns ----
    ' This works on ALL tables in the document
    Dim tbl As Object
    For Each tbl In doc.Tables
        Dim cel As Object
        For Each cel In tbl.Range.Cells
            Dim cellRng As Object
            Set cellRng = cel.Range
            
            ' Remove the trailing cell marker char from range (wdCharacter=1)
            cellRng.MoveEnd 1, -1
            
            ' 1. Permission number (005-24 -> 99-26)
            If oldPerm <> "" And newPerm <> "" And oldPerm <> newPerm Then
                totalChanges = totalChanges + ReplaceInRuns(cellRng, oldPerm, newPerm)
            End If
            
            ' 2. Inv number (18736 -> 26085)
            If oldInv <> "" And newInv <> "" And oldInv <> newInv Then
                totalChanges = totalChanges + ReplaceInRuns(cellRng, oldInv, newInv)
            End If
            
            ' 3. Cyrillic revision (С02 -> С03)
            If cyrOldRev <> cyrNewRev Then
                totalChanges = totalChanges + ReplaceInRuns(cellRng, cyrOldRev, cyrNewRev)
            End If
            
            ' 4. Latin revision (C02 -> C03) - for (C02) in row 8
            If oldRev <> newRev Then
                totalChanges = totalChanges + ReplaceInRuns(cellRng, oldRev, newRev)
            End If
            
            ' 5. Short date DD.MM.YY (15.12.24 -> 19.02.26)
            ' This is split across runs - ReplaceInRuns handles it!
            If oldDateShort <> "" And newDateShort <> "" And oldDateShort <> newDateShort Then
                totalChanges = totalChanges + ReplaceInRuns(cellRng, oldDateShort, newDateShort)
            End If
            
            ' 6. Full date DD.MM.YYYY (10.01.2024 -> 19.02.2026)
            If oldDateFull_DDMMYYYY <> "" And newDateFull <> "" And oldDateFull_DDMMYYYY <> newDateFull Then
                totalChanges = totalChanges + ReplaceInRuns(cellRng, oldDateFull_DDMMYYYY, newDateFull)
            End If
            
            ' 7. GIP surname Rus
            If oldSurRus <> "" And newSurRus <> "" And oldSurRus <> newSurRus Then
                totalChanges = totalChanges + ReplaceInRuns(cellRng, oldSurRus, newSurRus)
            End If
            
            ' 8. GIP surname Eng
            If oldSurEng <> "" And newSurEng <> "" And oldSurEng <> newSurEng Then
                totalChanges = totalChanges + ReplaceInRuns(cellRng, oldSurEng, newSurEng)
            End If
        Next cel
    Next tbl

    ' ---- v23: Process body text paragraphs NOT in tables ----
    ' PERM header area ("Разрешение №005-24 от 15.12.2024 Утвердил: Хохлачева")
    ' is body text outside tables — missed by table-only iteration above
    Dim para As Object
    For Each para In doc.Paragraphs
        Dim paraRng As Object
        Set paraRng = para.Range
        ' Check if paragraph is inside a table (wdWithInTable = 12)
        Dim pInTable As Boolean
        pInTable = False
        On Error Resume Next
        pInTable = paraRng.Information(12)
        If Err.Number <> 0 Then pInTable = False
        Err.Clear
        On Error GoTo ErrHandler
        If Not pInTable Then
            ' Permission number
            If oldPerm <> "" And newPerm <> "" And oldPerm <> newPerm Then
                totalChanges = totalChanges + ReplaceInCellRange(paraRng, oldPerm, newPerm)
            End If
            ' Inv number
            If oldInv <> "" And newInv <> "" And oldInv <> newInv Then
                totalChanges = totalChanges + ReplaceInCellRange(paraRng, oldInv, newInv)
            End If
            ' Cyrillic revision
            If cyrOldRev <> cyrNewRev Then
                totalChanges = totalChanges + ReplaceInCellRange(paraRng, cyrOldRev, cyrNewRev)
            End If
            ' Latin revision
            If oldRev <> newRev Then
                totalChanges = totalChanges + ReplaceInCellRange(paraRng, oldRev, newRev)
            End If
            ' Short date DD.MM.YY
            If oldDateShort <> "" And newDateShort <> "" And oldDateShort <> newDateShort Then
                totalChanges = totalChanges + ReplaceInCellRange(paraRng, oldDateShort, newDateShort)
            End If
            ' Full date DD.MM.YYYY
            If oldDateFull_DDMMYYYY <> "" And newDateFull <> "" And oldDateFull_DDMMYYYY <> newDateFull Then
                totalChanges = totalChanges + ReplaceInCellRange(paraRng, oldDateFull_DDMMYYYY, newDateFull)
            End If
            ' GIP full FIO
            If oldFioRus <> "" And newFioRus <> "" And oldFioRus <> newFioRus Then
                totalChanges = totalChanges + ReplaceInCellRange(paraRng, oldFioRus, newFioRus)
            End If
            ' GIP surname Rus
            If oldSurRus <> "" And newSurRus <> "" And oldSurRus <> newSurRus Then
                totalChanges = totalChanges + ReplaceInCellRange(paraRng, oldSurRus, newSurRus)
            End If
            ' GIP full FIO Eng
            If oldFioEng <> "" And newFioEng <> "" And oldFioEng <> newFioEng Then
                totalChanges = totalChanges + ReplaceInCellRange(paraRng, oldFioEng, newFioEng)
            End If
            ' GIP surname Eng
            If oldSurEng <> "" And newSurEng <> "" And oldSurEng <> newSurEng Then
                totalChanges = totalChanges + ReplaceInCellRange(paraRng, oldSurEng, newSurEng)
            End If
            ' Year
            If oldYear <> "" And newYear <> "" And oldYear <> newYear Then
                totalChanges = totalChanges + ReplaceInCellRange(paraRng, oldYear, newYear)
            End If
        End If
    Next para

    ' ---- v25 B-014: Replace authorization date in PERM header ----
    ' The PERM header cell "005-24 ot 10.01.2024" has a date that differs from
    ' the document date in Excel. After PERM number replacement, find any cell/para
    ' containing "ot " + DD.MM.YYYY and replace that date with newDateFull.
    ' ChrW(1086)=o, ChrW(1090)=t -> "ot" in Russian
    If newDateFull <> "" Then
        Dim otPattern As String
        otPattern = " " & ChrW(1086) & ChrW(1090) & " "  ' " ot "

        ' Scan table cells for "ot" + date
        Dim tblOt As Object
        For Each tblOt In doc.Tables
            Dim celOt As Object
            For Each celOt In tblOt.Range.Cells
                Dim celOtRng As Object
                Set celOtRng = celOt.Range
                celOtRng.MoveEnd 1, -1
                Dim celOtTxt As String
                celOtTxt = celOtRng.Text
                If InStr(celOtTxt, otPattern) > 0 Then
                    Dim otDate As String
                    otDate = FindDateDDMMYYYY(celOtTxt)
                    If otDate <> "" And otDate <> newDateFull Then
                        totalChanges = totalChanges + ReplaceInCellRange(celOtRng, otDate, newDateFull)
                    End If
                End If
            Next celOt
        Next tblOt

        ' Scan body text paragraphs for "ot" + date
        Dim paraOt As Object
        For Each paraOt In doc.Paragraphs
            Dim paraOtRng As Object
            Set paraOtRng = paraOt.Range
            Dim pOtInTable As Boolean
            pOtInTable = False
            On Error Resume Next
            pOtInTable = paraOtRng.Information(12)
            If Err.Number <> 0 Then pOtInTable = False
            Err.Clear
            On Error GoTo ErrHandler
            If Not pOtInTable Then
                Dim paraOtTxt As String
                paraOtTxt = paraOtRng.Text
                If InStr(paraOtTxt, otPattern) > 0 Then
                    Dim otDateP As String
                    otDateP = FindDateDDMMYYYY(paraOtTxt)
                    If otDateP <> "" And otDateP <> newDateFull Then
                        totalChanges = totalChanges + ReplaceInCellRange(paraOtRng, otDateP, newDateFull)
                    End If
                End If
            End If
        Next paraOt
    End If

    ' ---- SaveAs with new PERM name ----
    Dim newPermFileName As String
    newPermFileName = newPerm & ".docx"
    newPermPath = docFolder & newPermFileName
    
    On Error Resume Next
    doc.SaveAs2 fileName:=newPermPath, FileFormat:=12  ' wdFormatDocumentDefault
    Dim saveErr As Long
    saveErr = Err.Number
    On Error GoTo ErrHandler
    
    If saveErr <> 0 Then
        ' SaveAs failed - just Save in place
        doc.Save
        newPermPath = filePath
    End If
    
    ' ---- Export PDF ----
    pdfOk = False
    Dim pdfName As String
    pdfName = newPerm & ".pdf"
    Dim pdfPath As String
    pdfPath = pdfFolder & pdfName
    
    On Error Resume Next
    doc.SaveAs2 fileName:=pdfPath, FileFormat:=17  ' wdFormatPDF
    If Err.Number = 0 Then
        pdfOk = True
    Else
        pdfErrMsg = "PDF Err " & Err.Number
        Err.Clear
    End If
    On Error GoTo ErrHandler
    
    doc.Close SaveChanges:=False
    Set doc = Nothing
    
    ' ---- Delete original if renamed ----
    If saveErr = 0 And newPermPath <> filePath Then
        On Error Resume Next
        If fso.FileExists(filePath) Then fso.DeleteFile filePath, True
        On Error GoTo ErrHandler
    End If
    
    ProcessPermXml = totalChanges
    Exit Function
    
ErrHandler:
    pdfErrMsg = "PERM Err " & Err.Number & ": " & Err.Description
    ProcessPermXml = -1
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close SaveChanges:=False
    Set doc = Nothing
    On Error GoTo 0
End Function


' Count occurrences of searchStr in text (case-sensitive)
Function CountOccurrences(text As String, searchStr As String) As Long
    Dim count As Long
    count = 0
    Dim pos As Long
    pos = 1
    Do
        pos = InStr(pos, text, searchStr)
        If pos = 0 Then Exit Do
        count = count + 1
        pos = pos + Len(searchStr)
    Loop
    CountOccurrences = count
End Function


' Replace all occurrences of oldStr with newStr in XML (case-sensitive)
Function ReplaceAllInXml(xmlStr As String, oldStr As String, newStr As String) As String
    ReplaceAllInXml = Replace(xmlStr, oldStr, newStr)
End Function


' Replace date string that may be SPLIT ACROSS MULTIPLE RUNS within a paragraph.
' Approach: scan each <w:p>...</w:p> block, find all <w:t> text nodes,
' concatenate them, detect the date, then put new date in first affected run
' and clear others.
Function ReplaceDateInXmlParagraphs(xmlContent As String, _
                                     oldDate As String, newDate As String, _
                                     ByRef repCount As Long) As String
    repCount = 0
    Dim result As String
    result = xmlContent
    
    ' First try simple single-run replacement
    Dim simpleOld As String
    simpleOld = ">" & oldDate & "<"
    Dim simpleNew As String
    simpleNew = ">" & newDate & "<"
    
    Dim pos As Long
    pos = 1
    Dim newResult As String
    newResult = result
    
    Do
        pos = InStr(pos, newResult, simpleOld)
        If pos = 0 Then Exit Do
        newResult = Left(newResult, pos - 1) & simpleNew & Mid(newResult, pos + Len(simpleOld))
        repCount = repCount + 1
        pos = pos + Len(simpleNew)
    Loop
    result = newResult
    
    ' Now handle multi-run case: scan paragraphs
    ' Find each <w:p ... > ... </w:p> block
    Dim pStart As Long
    Dim pEnd As Long
    pStart = 1
    
    Do
        ' Find next paragraph start
        Dim pTagStart As Long
        pTagStart = InStr(pStart, result, "<w:p ")
        Dim pTagStart2 As Long
        pTagStart2 = InStr(pStart, result, "<w:p>")
        
        If pTagStart = 0 And pTagStart2 = 0 Then Exit Do
        If pTagStart = 0 Then pTagStart = pTagStart2
        If pTagStart2 = 0 Then
            ' pTagStart stays
        ElseIf pTagStart2 < pTagStart Then
            pTagStart = pTagStart2
        End If
        
        ' Find matching </w:p>
        pEnd = InStr(pTagStart, result, "</w:p>")
        If pEnd = 0 Then Exit Do
        pEnd = pEnd + 5  ' include </w:p>
        
        Dim paraXml As String
        paraXml = Mid(result, pTagStart, pEnd - pTagStart + 1)
        
        ' Extract all text content from w:t elements in this paragraph
        Dim combinedText As String
        combinedText = ExtractAllWtText(paraXml)
        
        ' Check if old date exists in combined text (even split across runs)
        If InStr(combinedText, oldDate) > 0 Then
            ' Replace: put newDate in first run containing part of oldDate, clear others
            Dim newParaXml As String
            newParaXml = ReplaceMultiRunDate(paraXml, oldDate, newDate)
            If newParaXml <> paraXml Then
                result = Left(result, pTagStart - 1) & newParaXml & Mid(result, pEnd + 1)
                repCount = repCount + 1
                pStart = pTagStart + Len(newParaXml)
            Else
                pStart = pEnd + 1
            End If
        Else
            pStart = pEnd + 1
        End If
    Loop
    
    ReplaceDateInXmlParagraphs = result
End Function


' Extract all text from w:t elements in a paragraph XML string
Function ExtractAllWtText(paraXml As String) As String
    Dim result As String
    result = ""
    Dim pos As Long
    pos = 1
    
    Do
        Dim wtStart As Long
        wtStart = InStr(pos, paraXml, "<w:t")
        If wtStart = 0 Then Exit Do
        
        ' Find end of opening tag
        Dim tagEnd As Long
        tagEnd = InStr(wtStart, paraXml, ">")
        If tagEnd = 0 Then Exit Do
        
        ' Check if self-closing
        If Mid(paraXml, tagEnd - 1, 1) = "/" Then
            pos = tagEnd + 1
        Else
            ' Find </w:t>
            Dim wtEnd As Long
            wtEnd = InStr(tagEnd, paraXml, "</w:t>")
            If wtEnd = 0 Then Exit Do
            result = result & Mid(paraXml, tagEnd + 1, wtEnd - tagEnd - 1)
            pos = wtEnd + 6
        End If
    Loop
    
    ExtractAllWtText = result
End Function


' Replace date that's split across multiple runs in a paragraph
' Strategy: find runs contributing to the date, put full newDate in first run, 
' clear chars from other runs that were part of oldDate
Function ReplaceMultiRunDate(paraXml As String, oldDate As String, newDate As String) As String
    ' Build list of w:t positions and their text
    Dim wtStarts() As Long
    Dim wtEnds() As Long   ' position of > in opening tag
    Dim wtTextStarts() As Long
    Dim wtTextEnds() As Long
    Dim wtTexts() As String
    Dim wtCount As Long
    wtCount = 0
    
    Dim pos As Long
    pos = 1
    
    Do
        Dim wtStart As Long
        wtStart = InStr(pos, paraXml, "<w:t")
        If wtStart = 0 Then Exit Do
        
        Dim tagEnd As Long
        tagEnd = InStr(wtStart, paraXml, ">")
        If tagEnd = 0 Then Exit Do
        
        ' Skip self-closing
        If Mid(paraXml, tagEnd - 1, 1) = "/" Then
            pos = tagEnd + 1
        Else
            Dim wtEnd As Long
            wtEnd = InStr(tagEnd, paraXml, "</w:t>")
            If wtEnd = 0 Then Exit Do
            
            wtCount = wtCount + 1
            ReDim Preserve wtStarts(1 To wtCount)
            ReDim Preserve wtEnds(1 To wtCount)
            ReDim Preserve wtTextStarts(1 To wtCount)
            ReDim Preserve wtTextEnds(1 To wtCount)
            ReDim Preserve wtTexts(1 To wtCount)
            
            wtStarts(wtCount) = wtStart
            wtEnds(wtCount) = tagEnd
            wtTextStarts(wtCount) = tagEnd + 1
            wtTextEnds(wtCount) = wtEnd - 1
            wtTexts(wtCount) = Mid(paraXml, tagEnd + 1, wtEnd - tagEnd - 1)
            
            pos = wtEnd + 6
        End If
    Loop
    
    If wtCount = 0 Then
        ReplaceMultiRunDate = paraXml
        Exit Function
    End If
    
    ' Build combined text with position tracking
    Dim combined As String
    combined = ""
    Dim runForChar() As Long
    Dim charInRun() As Long
    ReDim runForChar(1 To 1)
    ReDim charInRun(1 To 1)
    
    Dim ri As Long
    For ri = 1 To wtCount
        Dim ci As Long
        For ci = 1 To Len(wtTexts(ri))
            Dim combinedLen As Long
            combinedLen = Len(combined) + ci
            ReDim Preserve runForChar(1 To combinedLen)
            ReDim Preserve charInRun(1 To combinedLen)
            runForChar(combinedLen) = ri
            charInRun(combinedLen) = ci
        Next ci
        combined = combined & wtTexts(ri)
    Next ri
    
    ' Find oldDate in combined
    Dim datePos As Long
    datePos = InStr(combined, oldDate)
    If datePos = 0 Then
        ReplaceMultiRunDate = paraXml
        Exit Function
    End If
    
    ' Determine which runs are involved
    Dim firstRunIdx As Long
    Dim lastRunIdx As Long
    firstRunIdx = runForChar(datePos)
    lastRunIdx = runForChar(datePos + Len(oldDate) - 1)
    
    ' Strategy: 
    ' - In first run: replace the portion that is part of oldDate with newDate
    ' - In subsequent runs (up to lastRun): clear the portion that is part of oldDate
    
    ' How many chars of first run are before the date?
    Dim charsBeforeDate As Long
    charsBeforeDate = charInRun(datePos) - 1  ' chars in first run before date starts
    
    ' New text for first run: keep chars before date, add newDate
    ' (the rest of the date chars in first run are consumed)
    Dim firstRunText As String
    firstRunText = Left(wtTexts(firstRunIdx), charsBeforeDate) & newDate
    
    ' For runs firstRunIdx+1 to lastRunIdx: remove chars that belong to oldDate
    ' How many chars of oldDate were in first run?
    Dim dateCharsInFirst As Long
    dateCharsInFirst = Len(wtTexts(firstRunIdx)) - charsBeforeDate
    
    ' Build modified para XML
    Dim newPara As String
    newPara = paraXml
    Dim offset As Long
    offset = 0  ' track changes in string positions due to replacements
    
    ' Replace text in first run
    Dim origFirstStart As Long
    Dim origFirstEnd As Long
    origFirstStart = wtTextStarts(firstRunIdx) + offset
    origFirstEnd = wtTextEnds(firstRunIdx) + offset
    newPara = Left(newPara, origFirstStart - 1) & firstRunText & Mid(newPara, origFirstEnd + 1)
    offset = offset + (Len(firstRunText) - Len(wtTexts(firstRunIdx)))
    
    ' Clear date portions from subsequent runs
    Dim dateCharsConsumed As Long
    dateCharsConsumed = dateCharsInFirst
    
    Dim rii As Long
    For rii = firstRunIdx + 1 To lastRunIdx
        Dim charsToRemove As Long
        charsToRemove = Len(oldDate) - dateCharsConsumed
        If charsToRemove > Len(wtTexts(rii)) Then
            charsToRemove = Len(wtTexts(rii))
        End If
        
        Dim newRunText As String
        newRunText = Mid(wtTexts(rii), charsToRemove + 1)  ' keep chars after date portion
        
        Dim runTextStart As Long
        Dim runTextEnd As Long
        runTextStart = wtTextStarts(rii) + offset
        runTextEnd = wtTextEnds(rii) + offset
        newPara = Left(newPara, runTextStart - 1) & newRunText & Mid(newPara, runTextEnd + 1)
        offset = offset + (Len(newRunText) - Len(wtTexts(rii)))
        
        dateCharsConsumed = dateCharsConsumed + charsToRemove
    Next rii
    
    ReplaceMultiRunDate = newPara
End Function


' Replace any DD.MM.YYYY date containing oldYear in XML
Function ReplaceDateYearInXml(xmlContent As String, oldYear As String, _
                               newDateFull As String, ByRef repCount As Long) As String
    repCount = 0
    Dim result As String
    result = xmlContent
    
    ' Scan for pattern: any 10-char sequence "DD.MM.YYYY" ending with oldYear
    ' This is a simple scan since we know dates are in text nodes
    Dim pos As Long
    pos = 1
    
    Do
        Dim yearPos As Long
        yearPos = InStr(pos, result, oldYear)
        If yearPos = 0 Then Exit Do
        
        ' Check if preceded by DD.MM. (6 chars before year)
        If yearPos >= 7 Then
            Dim candidate As String
            candidate = Mid(result, yearPos - 6, 10)
            ' Validate DD.MM.YYYY format
            If Mid(candidate, 3, 1) = "." And Mid(candidate, 6, 1) = "." Then
                If IsNumeric(Left(candidate, 2)) And IsNumeric(Mid(candidate, 4, 2)) Then
                    ' Check it's in a text node context (between > and <)
                    ' Check char before candidate is > or middle of text
                    Dim before As String
                    If yearPos > 7 Then before = Mid(result, yearPos - 7, 1) Else before = ""
                    Dim after As String
                    after = Mid(result, yearPos + 4, 1)
                    
                    If after = "<" Or after = " " Or after = Chr(13) Or after = Chr(10) Then
                        result = Left(result, yearPos - 7) & newDateFull & Mid(result, yearPos + 4)
                        repCount = repCount + 1
                        pos = yearPos - 6 + Len(newDateFull)
                    Else
                        pos = yearPos + 4
                    End If
                Else
                    pos = yearPos + 4
                End If
            Else
                pos = yearPos + 4
            End If
        Else
            pos = yearPos + 4
        End If
    Loop
    
    ReplaceDateYearInXml = result
End Function





' Try to start/get AutoCAD across multiple versions (Civil/AutoCAD)
Function CreateAutoCadApp() As Object
    Dim app As Object
    Set app = Nothing

    On Error Resume Next
    Set app = GetObject(, "AutoCAD.Application")
    If app Is Nothing Then Set app = CreateObject("AutoCAD.Application.26")
    If app Is Nothing Then Set app = CreateObject("AutoCAD.Application.25")
    If app Is Nothing Then Set app = CreateObject("AutoCAD.Application.24")
    If app Is Nothing Then Set app = CreateObject("AutoCAD.Application.23")
    If app Is Nothing Then Set app = CreateObject("AutoCAD.Application")
    On Error GoTo 0

    Set CreateAutoCadApp = app
End Function

Sub PrepareAutoCadSession(acadApp As Object)
    If acadApp Is Nothing Then Exit Sub
    On Error Resume Next
    acadApp.Visible = True
    acadApp.SetSystemVariable "PROXYNOTICE", 0
    acadApp.SetSystemVariable "PROXYSHOW", 0
    acadApp.SetSystemVariable "FILEDIA", 0
    acadApp.SetSystemVariable "CMDDIA", 0
    acadApp.SetSystemVariable "RECOVERYMODE", 0
    Err.Clear
    On Error GoTo 0
End Sub

Sub UpdateRevision()
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' ========================================
    ' 1. Auto-detect revisions
    ' ========================================
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_SHIFR).End(xlUp).Row
    
    If lastRow < DATA_START_ROW Then
        MsgBox "No data in table!", vbExclamation
        Exit Sub
    End If
    
    Dim OLD_REVISION As String
    OLD_REVISION = ""
    Dim r As Long
    For r = DATA_START_ROW To lastRow
        Dim revVal As String
        revVal = NormalizeRev(CStr(ws.Cells(r, COL_OLD_REV).Value & ""))
        If revVal <> "" Then
            OLD_REVISION = revVal
            Exit For
        End If
    Next r
    
    Dim NEW_REVISION As String
    NEW_REVISION = DetectNewRevFromHeader(ws)
    
    If NEW_REVISION = "" And OLD_REVISION <> "" Then
        Dim revNum As Long
        revNum = Val(Mid(OLD_REVISION, 2))
        NEW_REVISION = "C" & Format(revNum + 1, "00")
    End If
    
    If OLD_REVISION = "" Or NEW_REVISION = "" Then
        MsgBox "Cannot detect revisions!", vbCritical
        Exit Sub
    End If
    
    ' ========================================
    ' 1b. Read OLD_YEAR and OLD_FIO from Excel
    ' ========================================
    ' OLD_YEAR: from column H (old date), first non-empty row
    Dim OLD_YEAR As String
    OLD_YEAR = ""
    For r = DATA_START_ROW To lastRow
        Dim oldDtV As Variant
        oldDtV = ws.Cells(r, COL_OLD_DATE).Value
        If IsDate(oldDtV) Then
            OLD_YEAR = CStr(Year(CDate(oldDtV)))
            Exit For
        End If
    Next r
    ' Fallback: extract year from any date-like string in column H
    If OLD_YEAR = "" Then
        For r = DATA_START_ROW To lastRow
            Dim cellStr As String
            cellStr = Trim(CStr(ws.Cells(r, COL_OLD_DATE).Value & ""))
            If Len(cellStr) >= 4 Then
                ' Try last 4 chars as year
                Dim yr As String
                yr = Right(cellStr, 4)
                If IsNumeric(yr) And CLng(yr) >= 2000 And CLng(yr) <= 2099 Then
                    OLD_YEAR = yr
                    Exit For
                End If
            End If
        Next r
    End If
    If OLD_YEAR = "" Then OLD_YEAR = CStr(Year(Now) - 1)  ' fallback
    
    ' OLD_FIO_RUS: from header row COL_OLD_GIP (column J, row HEADER_ROW)
    ' This is the old GIP name that was current in the previous revision.
    ' Convention: row 3 (HEADER_ROW) in column J has the old GIP as column header label
    ' If not found there, read from first data row that has a value different from new GIP.
    Dim OLD_FIO_RUS As String
    OLD_FIO_RUS = ""
    
    ' First try: read from header row (row 3), col J
    Dim hdrGip As String
    hdrGip = Trim(CStr(ws.Cells(HEADER_ROW, COL_OLD_GIP).Value & ""))
    If hdrGip <> "" And Not IsNumeric(hdrGip) And hdrGip <> "GIP" And hdrGip <> "ГИП" Then
        OLD_FIO_RUS = hdrGip
    End If
    
    ' Second try: if header doesn't have it, read new GIP and infer old from a separate cell
    ' Fallback: read first data row GIP (which will be new GIP) and OLD is unknown
    Dim NEW_GIP_RUS As String
    NEW_GIP_RUS = ""
    For r = DATA_START_ROW To lastRow
        Dim gipVal As String
        gipVal = Trim(CStr(ws.Cells(r, COL_GIP).Value & ""))
        If gipVal <> "" Then
            ' Detect if Russian
            Dim isRussian As Boolean
            isRussian = False
            Dim c As Long
            For c = 1 To Len(gipVal)
                If AscW(Mid(gipVal, c, 1)) >= 1040 Then
                    isRussian = True
                    Exit For
                End If
            Next c
            If isRussian And NEW_GIP_RUS = "" Then NEW_GIP_RUS = gipVal
        End If
        If NEW_GIP_RUS <> "" Then Exit For
    Next r
    
    ' v18 fix: if OLD_FIO_RUS is empty or same as new GIP, skip FIO replacement silently
    If OLD_FIO_RUS = NEW_GIP_RUS Then OLD_FIO_RUS = ""
    
    ' OLD_FIO_ENG: transliterate from OLD_FIO_RUS
    Dim OLD_FIO_ENG As String
    OLD_FIO_ENG = Transliterate(OLD_FIO_RUS)
    
    ' ========================================
    ' 2. Read data from xlsx
    ' ========================================
    Dim newFioRus As String
    Dim newFioEng As String
    newFioRus = NEW_GIP_RUS  ' already found above
    newFioEng = ""
    
    ' Also check for English GIP name
    For r = DATA_START_ROW To lastRow
        Dim gipVal2 As String
        gipVal2 = Trim(CStr(ws.Cells(r, COL_GIP).Value & ""))
        If gipVal2 <> "" Then
            Dim isRus2 As Boolean
            isRus2 = False
            Dim c2 As Long
            For c2 = 1 To Len(gipVal2)
                If AscW(Mid(gipVal2, c2, 1)) >= 1040 Then
                    isRus2 = True
                    Exit For
                End If
            Next c2
            If Not isRus2 And newFioEng = "" Then newFioEng = gipVal2
        End If
        If newFioEng <> "" Then Exit For
    Next r
    
    If newFioEng = "" And newFioRus <> "" Then
        newFioEng = Transliterate(newFioRus)
    End If
    
    ' Get new date for folder name (from col M)
    Dim newDateForFolder As String
    newDateForFolder = ""
    For r = DATA_START_ROW To lastRow
        Dim dtVal As Variant
        dtVal = ws.Cells(r, COL_DATE).Value
        If IsDate(dtVal) Then
            newDateForFolder = FormatDateFolder(CDate(dtVal))
            Exit For
        End If
    Next r
    
    ' Collect per-shifr data
    Dim shifrs As New Collection
    Dim oldInvs As New Collection
    Dim newInvs As New Collection
    Dim newDates As New Collection
    Dim newDateShorts As New Collection
    Dim oldPerms As New Collection
    Dim newPerms As New Collection
    Dim newYears As New Collection
    Dim oldDateDDMMs As New Collection
    Dim newDateDDMMs As New Collection
    
    For r = DATA_START_ROW To lastRow
        Dim shifrVal As String
        shifrVal = Trim(CStr(ws.Cells(r, COL_SHIFR).Value & ""))
        If shifrVal = "" Then GoTo NextRow
        
        Dim oldInv As String
        oldInv = Trim(CStr(ws.Cells(r, COL_OLD_INV).Value & ""))
        Dim newInv As String
        newInv = Trim(CStr(ws.Cells(r, COL_NEW_INV).Value & ""))
        Dim oldPerm As String
        oldPerm = Trim(CStr(ws.Cells(r, COL_OLD_PERM).Value & ""))
        Dim newPerm As String
        newPerm = Trim(CStr(ws.Cells(r, COL_NEW_PERM).Value & ""))
        
        Dim oldDateVal As Variant
        oldDateVal = ws.Cells(r, COL_OLD_DATE).Value
        Dim oldDateDDMM As String
        oldDateDDMM = ""
        If IsDate(oldDateVal) Then
            oldDateDDMM = Format(CDate(oldDateVal), "DD.MM.")
        End If
        
        Dim dateVal As Variant
        dateVal = ws.Cells(r, COL_DATE).Value
        Dim newDateStr As String
        Dim newDateShort As String
        Dim newYear As String
        newDateStr = ""
        newDateShort = ""
        newYear = ""
        If IsDate(dateVal) Then
            newDateStr = FormatDateDMY(CDate(dateVal))
            newDateShort = FormatDateShort(CDate(dateVal))
            newYear = CStr(Year(CDate(dateVal)))
        ElseIf IsNumeric(dateVal) And dateVal > 2000 And dateVal < 2100 Then
            newYear = CStr(CLng(dateVal))
        End If
        
        Dim newDateDDMM As String
        newDateDDMM = ""
        If IsDate(dateVal) Then
            newDateDDMM = Format(CDate(dateVal), "DD.MM.")
        End If
        
        shifrs.Add shifrVal
        oldInvs.Add oldInv
        newInvs.Add newInv
        newDates.Add newDateStr
        newDateShorts.Add newDateShort
        oldPerms.Add oldPerm
        newPerms.Add newPerm
        newYears.Add newYear
        oldDateDDMMs.Add oldDateDDMM
        newDateDDMMs.Add newDateDDMM
NextRow:
    Next r
    
    If shifrs.Count = 0 Then
        MsgBox "No data found!", vbExclamation
        Exit Sub
    End If
    
    ' Fill empty oldDateDDMMs from first non-empty row
    Dim firstOldDateDDMM As String
    Dim i As Long
    firstOldDateDDMM = ""
    For i = 1 To oldDateDDMMs.Count
        If oldDateDDMMs(i) <> "" Then
            firstOldDateDDMM = oldDateDDMMs(i)
            Exit For
        End If
    Next i
    If firstOldDateDDMM <> "" Then
        For i = 1 To oldDateDDMMs.Count
            If oldDateDDMMs(i) = "" Then
                ' Replace empty with first found
                Dim tempColl As New Collection
                Dim j2 As Long
                For j2 = 1 To oldDateDDMMs.Count
                    If j2 = i Then
                        tempColl.Add firstOldDateDDMM
                    Else
                        tempColl.Add oldDateDDMMs(j2)
                    End If
                Next j2
                Set oldDateDDMMs = tempColl
            End If
        Next i
    End If
    
    ' ========================================
    ' 3. CONFIRMATION DIALOG
    ' ========================================
    Dim revMsg As String
    revMsg = "==================================" & vbCrLf
    revMsg = revMsg & "   REVISION UPDATE" & vbCrLf
    revMsg = revMsg & "==================================" & vbCrLf & vbCrLf
    revMsg = revMsg & "   " & OLD_REVISION & "  >>>  " & NEW_REVISION & vbCrLf & vbCrLf
    revMsg = revMsg & "Sets: " & shifrs.Count & vbCrLf
    For i = 1 To shifrs.Count
        revMsg = revMsg & "  " & i & ". " & shifrs(i) & vbCrLf
    Next i
    revMsg = revMsg & vbCrLf
    revMsg = revMsg & "FIO: " & OLD_FIO_RUS & " -> " & newFioRus & vbCrLf
    revMsg = revMsg & "FIO eng: " & OLD_FIO_ENG & " -> " & newFioEng & vbCrLf
    revMsg = revMsg & "New date: " & newDateForFolder & vbCrLf & vbCrLf
    revMsg = revMsg & "Pipeline:" & vbCrLf
    revMsg = revMsg & "  1. Delete .pdf/.bak/.dwl/.dwl2/~$ + logs" & vbCrLf
    revMsg = revMsg & "  2. Rename all folders & files" & vbCrLf
    revMsg = revMsg & "  3. Update docx content & stamps" & vbCrLf
    revMsg = revMsg & "  4. Export PDF to " & PDF_FOLDER_NAME & vbCrLf & vbCrLf
    revMsg = revMsg & "Continue?"
    
    If MsgBox(revMsg, vbYesNo + vbQuestion + vbDefaultButton2, _
              "Rev " & OLD_REVISION & " -> " & NEW_REVISION) = vbNo Then
        Exit Sub
    End If
    
    ' ========================================
    ' 4. Select MAIN folder
    ' ========================================
    Dim mainFolder As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select MAIN folder (..._" & OLD_REVISION & " ot ...)"
        If .Show = -1 Then
            mainFolder = .SelectedItems(1)
            If Right(mainFolder, 1) = "\" Then
                mainFolder = Left(mainFolder, Len(mainFolder) - 1)
            End If
        Else
            Exit Sub
        End If
    End With
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    TRACE_LOG_PATH = mainFolder & "\UPDATE_TRACE.log"
    On Error Resume Next
    Dim ffTrace As Long
    ffTrace = FreeFile
    Open TRACE_LOG_PATH For Output As #ffTrace
    Print #ffTrace, "TRACE START"
    Close #ffTrace
    On Error GoTo 0
    AppendTrace "Run start OLD=" & OLD_REVISION & " NEW=" & NEW_REVISION

    Dim logText As String
    logText = ""
    
    ' ========================================
    ' STAGE 1A: Delete .pdf and .bak
    ' ========================================
    Application.StatusBar = "Stage 1: Deleting generated/temp files..."
    logText = logText & "=== STAGE 1A: DELETE PDF/BAK/TEMP ===" & vbCrLf
    
    Dim delCount As Long
    delCount = 0
    DeleteFilesRecursive fso, mainFolder, OLD_REVISION, delCount, logText
    logText = logText & "Deleted: " & delCount & vbCrLf & vbCrLf
    
    ' ========================================
    ' STAGE 1B: Rename folders & files (bottom-up)
    ' ========================================
    Application.StatusBar = "Stage 1: Renaming..."
    logText = logText & "=== STAGE 1B: RENAME ===" & vbCrLf
    
    Dim renCount As Long
    renCount = 0
    RenameFoldersRecursive fso, mainFolder, OLD_REVISION, NEW_REVISION, renCount, logText
    
    ' Rename main folder itself
    Dim mainFolderName As String
    mainFolderName = fso.GetFolder(mainFolder).Name
    If InStr(1, mainFolderName, OLD_REVISION, vbTextCompare) > 0 Or _
       InStr(1, mainFolderName, Replace(OLD_REVISION, "C", ChrW(1057)), vbBinaryCompare) > 0 Then
        Dim newMainName As String
        newMainName = Replace(mainFolderName, OLD_REVISION, NEW_REVISION, , , vbBinaryCompare)
        newMainName = Replace(newMainName, Replace(OLD_REVISION, "C", ChrW(1057)), NEW_REVISION, , , vbBinaryCompare)
        
        ' Also update date in folder name if present
        ' Find DD.MM.YYYY pattern and replace with new date
        If newDateForFolder <> "" Then
            Dim j As Long
            For j = 1 To Len(newMainName) - 9
                Dim chunk As String
                chunk = Mid(newMainName, j, 10)
                If Mid(chunk, 3, 1) = "." And Mid(chunk, 6, 1) = "." And IsNumeric(Mid(chunk, 1, 2)) Then
                    newMainName = Left(newMainName, j - 1) & newDateForFolder & Mid(newMainName, j + 10)
                    Exit For
                End If
            Next j
        End If
        
        Dim parentPath As String
        parentPath = GetParentFolder(mainFolder)
        Dim newMainPath As String
        newMainPath = parentPath & newMainName
        
        If Not fso.FolderExists(newMainPath) Then
            On Error Resume Next
            fso.MoveFolder mainFolder, newMainPath
            If Err.Number = 0 Then
                logText = logText & "  MAIN: " & mainFolderName & " -> " & newMainName & vbCrLf
                renCount = renCount + 1
                mainFolder = newMainPath
            Else
                logText = logText & "  MAIN ERR: " & mainFolderName & " Err" & Err.Number & " " & Err.Description & vbCrLf
                Err.Clear
            End If
            On Error GoTo 0
        End If
    End If
    
    logText = logText & "Renamed: " & renCount & vbCrLf & vbCrLf
    
    ' ========================================
    ' STAGE 2: Find docx folders and process
    ' ========================================
    Application.StatusBar = "Stage 2: Processing docx..."
    logText = logText & "=== STAGE 2: DOCX PROCESSING ===" & vbCrLf
    
    ' Find all folders with docx files
    Dim docxFolders As New Collection
    FindDocxFolders fso, mainFolder, docxFolders
    
    logText = logText & "Found " & docxFolders.Count & " folder(s) with docx/dwg" & vbCrLf

    ' v18: Diagnostic log line
    logText = logText & "[Excel] Shifrs=" & shifrs.Count & " OLD_FIO='" & OLD_FIO_RUS & "' -> '" & newFioRus & "' OLD_YEAR=" & OLD_YEAR & vbCrLf

    ' Open Word
    Dim wordApp As Object
    On Error Resume Next
    Set wordApp = GetObject(, "Word.Application")
    If wordApp Is Nothing Then
        Set wordApp = CreateObject("Word.Application")
    End If
    On Error GoTo 0
    
    If wordApp Is Nothing Then
        MsgBox "Cannot start Word!" & vbCrLf & vbCrLf & logText, vbCritical
        Exit Sub
    End If
    
    wordApp.Visible = False
    wordApp.DisplayAlerts = 0

    ' v20+: Open AutoCAD (optional - DWG skipped if not available)
    Dim acadApp As Object
    Dim exportDwgPdf As Boolean
    Dim forceModelScan As Boolean
    exportDwgPdf = False
    forceModelScan = False
    Set acadApp = CreateAutoCadApp()

    If acadApp Is Nothing Then
        logText = logText & "[AutoCAD] Not available - DWG files will be SKIPPED" & vbCrLf
    Else
        PrepareAutoCadSession acadApp

        ' Home/laptop workaround: allow running DOCX-only when Civil causes DWG hangs
        Dim processDwgChoice As VbMsgBoxResult
        processDwgChoice = MsgBox("Process DWG files too?" & vbCrLf & _
                            "Yes = DOCX + DWG" & vbCrLf & _
                            "No = DOCX only (recommended if Civil 3D hangs)", _
                            vbYesNo + vbQuestion + vbDefaultButton2, "DWG processing")
        If processDwgChoice = vbNo Then
            Set acadApp = Nothing
            logText = logText & "[AutoCAD] DWG processing disabled by user" & vbCrLf
        Else
            Dim exportDwgPdfChoice As VbMsgBoxResult
            exportDwgPdfChoice = MsgBox("Export DWG to PDF too?" & vbCrLf & _
                                        "Yes = run PlotToFile for layouts" & vbCrLf & _
                                        "No = skip DWG PDF export (recommended if plotting hangs)", _
                                        vbYesNo + vbQuestion + vbDefaultButton2, "DWG PDF export")
            exportDwgPdf = (exportDwgPdfChoice = vbYes)
            If Not exportDwgPdf Then
                logText = logText & "[AutoCAD] DWG PDF export disabled by user" & vbCrLf
            End If

            Dim forceModelScanChoice As VbMsgBoxResult
            forceModelScanChoice = MsgBox("Scan Model space even if CLD layouts exist?" & vbCrLf & _
                                          "Yes = process Model too (needed when stamp is in Model)" & vbCrLf & _
                                          "No = skip Model for speed", _
                                          vbYesNo + vbQuestion + vbDefaultButton1, "Model space scan")
            forceModelScan = (forceModelScanChoice = vbYes)
            If forceModelScan Then
                logText = logText & "[AutoCAD] Model scan forced ON" & vbCrLf
            Else
                logText = logText & "[AutoCAD] Model scan forced OFF" & vbCrLf
            End If
        End If
    End If

    Dim totalFiles As Long
    Dim totalChanges As Long
    Dim totalPdf As Long
    totalFiles = 0
    totalChanges = 0
    totalPdf = 0

    ' Process each docx/dwg folder
    Dim fi As Long
    For fi = 1 To docxFolders.Count
        Dim docxFolder As String
        docxFolder = docxFolders(fi)
        
        logText = logText & vbCrLf & ">> Folder: " & docxFolder & vbCrLf
        
        ' Create PDF_ASSEMBLY in parent of docx folder
        Dim pdfFolder As String
        pdfFolder = GetParentFolder(docxFolder) & PDF_FOLDER_NAME & "\"
        If Not fso.FolderExists(pdfFolder) Then
            On Error Resume Next
            fso.CreateFolder pdfFolder
            On Error GoTo 0
        End If
        
        ' Process each shifr
        For i = 1 To shifrs.Count
            Dim currentShifr As String
            currentShifr = shifrs(i)
            
            Dim shifrUnderscore As String
            shifrUnderscore = Replace(currentShifr, ".", "_")
            
            ' Collect Word files in this folder (.docx + .doc)
            Dim docFiles() As String
            Dim dCnt As Long
            dCnt = 0
            
            Dim dFile As String
            dFile = Dir(docxFolder & "\*.docx")
            Do While dFile <> ""
                dCnt = dCnt + 1
                ReDim Preserve docFiles(1 To dCnt)
                docFiles(dCnt) = dFile
                dFile = Dir()
            Loop
            dFile = Dir(docxFolder & "\*.doc")
            Do While dFile <> ""
                dCnt = dCnt + 1
                ReDim Preserve docFiles(1 To dCnt)
                docFiles(dCnt) = dFile
                dFile = Dir()
            Loop
            
            Dim di As Long
            For di = 1 To dCnt
                Dim fileName As String
                fileName = docFiles(di)
                
                Dim fnameNorm As String
                fnameNorm = Replace(fileName, "_", ".")
                
                ' Check if file matches current shifr OR is a permission file
                Dim matchesShifr As Boolean
                Dim matchesPerm As Boolean
                matchesShifr = (InStr(1, fnameNorm, currentShifr, vbTextCompare) > 0) Or _
                               (InStr(1, fileName, shifrUnderscore, vbTextCompare) > 0)
                matchesPerm = False
                If oldPerms(i) <> "" Then
                    Dim permBase As String
                    permBase = Left(fileName, InStrRev(fileName, ".") - 1)
                    matchesPerm = (permBase = oldPerms(i))
                End If
                
                If matchesShifr Or matchesPerm Then

                    If Left(fileName, 1) <> "~" Then
                        Dim fullPath As String
                        fullPath = docxFolder & "\" & fileName

                        ' v19: Check Word is alive; restart if crashed
                        On Error Resume Next
                        Dim wdAlive As Long
                        wdAlive = wordApp.Documents.Count
                        If Err.Number <> 0 Then
                            Err.Clear
                            Set wordApp = Nothing
                            Set wordApp = CreateObject("Word.Application")
                            wordApp.Visible = False
                            wordApp.DisplayAlerts = 0
                            logText = logText & "  [Word restarted]" & vbCrLf
                        End If
                        On Error GoTo 0

                        Application.StatusBar = "Processing: " & fileName
                        
                        Dim isCAB As Boolean
                        Dim isCAZ As Boolean
                        Dim isPERM As Boolean
                        isCAB = (InStr(1, UCase(fileName), "CAB") > 0)
                        isCAZ = (InStr(1, UCase(fileName), "CAZ") > 0)
                        ' Permission file: name matches old permission number (e.g. "005-24.docx")
                        Dim fnameNoExt As String
                        fnameNoExt = Left(fileName, InStrRev(fileName, ".") - 1)
                        isPERM = (fnameNoExt = oldPerms(i))
                        
                        Dim changes As Long
                        Dim pdfOk As Boolean
                        Dim pdfErrMsg As String
                        
                        If isPERM Then
                            ' === PERM: Use XML-based processing to preserve table formatting ===
                            Dim permNewPath As String
                            Dim oldDateShortPerm As String
                            Dim oldDateFullPerm As String
                            ' Build old short date DD.MM.YY from oldDateDDMMs(i) + oldYear short
                            oldDateShortPerm = ""
                            If oldDateDDMMs(i) <> "" Then
                                oldDateShortPerm = oldDateDDMMs(i) & Right(OLD_YEAR, 2)
                            End If
                            ' Build old full date DD.MM.YYYY (first row's date for PERM)
                            oldDateFullPerm = ""
                            If oldDateDDMMs(i) <> "" Then
                                oldDateFullPerm = oldDateDDMMs(i) & OLD_YEAR
                            End If
                            
                            changes = ProcessPermXml(fullPath, _
                                                     OLD_REVISION, NEW_REVISION, _
                                                     OLD_YEAR, newYears(i), _
                                                     OLD_FIO_RUS, newFioRus, _
                                                     OLD_FIO_ENG, newFioEng, _
                                                     oldInvs(i), newInvs(i), _
                                                     oldPerms(i), newPerms(i), _
                                                     oldDateShortPerm, newDateShorts(i), _
                                                     oldDateFullPerm, newDates(i), _
                                                     wordApp, pdfFolder, _
                                                     pdfOk, pdfErrMsg, permNewPath)
                        Else
                            changes = ProcessDocx(wordApp, fullPath, pdfFolder, _
                                                 OLD_REVISION, NEW_REVISION, _
                                                 OLD_YEAR, newYears(i), _
                                                 OLD_FIO_RUS, newFioRus, _
                                                 OLD_FIO_ENG, newFioEng, _
                                                 isCAB, isCAZ, isPERM, _
                                                 oldInvs(i), newInvs(i), _
                                                 newDates(i), newDateShorts(i), _
                                                 oldPerms(i), newPerms(i), _
                                                 oldDateDDMMs(i), newDateDDMMs(i), _
                                                 pdfOk, pdfErrMsg)
                        End If
                        
                        If changes >= 0 Then
                            logText = logText & "  OK: " & fileName & " (" & changes & " chg)"
                            totalChanges = totalChanges + changes
                            If pdfOk Then
                                logText = logText & " +PDF"
                                totalPdf = totalPdf + 1
                            End If
                            ' v25: show logo status and/or PDF error
                            If pdfErrMsg <> "" Then
                                If InStr(pdfErrMsg, "[logo updated]") > 0 Then
                                    logText = logText & " [logo updated]"
                                End If
                                If Not pdfOk And InStr(pdfErrMsg, "PDF Err") > 0 Then
                                    logText = logText & " PDF_ERR:" & pdfErrMsg
                                End If
                            End If
                            logText = logText & vbCrLf
                            
                            ' PERM file was renamed via SaveAs/XML in ProcessPerm/ProcessDocx
                            If isPERM And oldPerms(i) <> "" And newPerms(i) <> "" Then
                                logText = logText & "  PERM SAVED AS: " & newPerms(i) & ".docx [Word run method]" & vbCrLf
                            End If
                        Else
                            logText = logText & "  ERR: " & fileName & " " & pdfErrMsg & vbCrLf
                        End If
                        totalFiles = totalFiles + 1
                    End If
                End If
            Next di
        Next i

        ' ---- DWG processing in this folder (v20+) ----
        If Not acadApp Is Nothing Then
            For i = 1 To shifrs.Count
                Dim currentShifrDwg As String
                currentShifrDwg = shifrs(i)
                Dim shifrUnderscoreDwg As String
                shifrUnderscoreDwg = Replace(currentShifrDwg, ".", "_")

                Dim dwgFiles() As String
                Dim dwgCnt As Long
                dwgCnt = 0
                Dim dwgFile As String
                dwgFile = Dir(docxFolder & "\*.dwg")
                Do While dwgFile <> ""
                    dwgCnt = dwgCnt + 1
                    ReDim Preserve dwgFiles(1 To dwgCnt)
                    dwgFiles(dwgCnt) = dwgFile
                    dwgFile = Dir()
                Loop

                Dim dwi As Long
                For dwi = 1 To dwgCnt
                    Dim dwgName As String
                    dwgName = dwgFiles(dwi)
                    If Left(dwgName, 1) = "~" Then GoTo NextDwg

                    Dim dwgNorm As String
                    dwgNorm = Replace(dwgName, "_", ".")
                    Dim dwgMatchShifr As Boolean
                    dwgMatchShifr = (InStr(1, dwgNorm, currentShifrDwg, vbTextCompare) > 0) Or _
                                    (InStr(1, dwgName, shifrUnderscoreDwg, vbTextCompare) > 0)

                    If dwgMatchShifr Then
                        Dim dwgFullPath As String
                        dwgFullPath = docxFolder & "\" & dwgName
                        Application.StatusBar = "DWG: " & dwgName

                        ' Build old dates for DWG
                        Dim oldDateFullDwg As String
                        oldDateFullDwg = ""
                        If oldDateDDMMs(i) <> "" Then oldDateFullDwg = oldDateDDMMs(i) & OLD_YEAR
                        Dim oldDateShortDwg As String
                        oldDateShortDwg = ""
                        If oldDateDDMMs(i) <> "" Then oldDateShortDwg = oldDateDDMMs(i) & Right(OLD_YEAR, 2)

                        Dim dwgChanges As Long
                        Dim dwgDiagStr As String
                        dwgDiagStr = ""

                        ' P2: check AutoCAD COM liveliness and restart if needed (T-019)
                        On Error Resume Next
                        Dim acadAlive As Long
                        acadAlive = acadApp.Documents.Count
                        If Err.Number <> 0 Then
                            Err.Clear
                            Set acadApp = Nothing
                            Set acadApp = CreateAutoCadApp()
                            If Not acadApp Is Nothing Then
                                PrepareAutoCadSession acadApp
                                logText = logText & "  [AutoCAD restarted]" & vbCrLf
                            Else
                                logText = logText & "  [AutoCAD restart failed]" & vbCrLf
                                On Error GoTo 0
                                GoTo NextDwg
                            End If
                        End If
                        On Error GoTo 0

                        ' v25: pass pdfFolder for B-018 DWG PDF export
                        Dim dwgPdfFolder As String
                        If exportDwgPdf Then
                            dwgPdfFolder = pdfFolder
                        Else
                            dwgPdfFolder = ""
                        End If

                        AppendTrace "DWG BEGIN: " & dwgName
                        dwgChanges = ProcessDwg(acadApp, dwgFullPath, _
                                                OLD_REVISION, NEW_REVISION, _
                                                oldInvs(i), newInvs(i), _
                                                oldPerms(i), newPerms(i), _
                                                oldDateFullDwg, newDates(i), _
                                                oldDateShortDwg, newDateShorts(i), _
                                                OLD_YEAR, newYears(i), dwgDiagStr, _
                                                dwgPdfFolder, forceModelScan)

                        If dwgChanges >= 0 Then
                            AppendTrace "DWG OK: " & dwgName & " chg=" & dwgChanges & " " & dwgDiagStr
                            logText = logText & "  OK DWG: " & dwgName & " (" & dwgChanges & " chg) " & dwgDiagStr & vbCrLf
                            totalChanges = totalChanges + dwgChanges
                            ' v25: count DWG PDFs in total
                            If InStr(dwgDiagStr, "+PDF(") > 0 Then totalPdf = totalPdf + 1
                        Else
                            AppendTrace "DWG ERR: " & dwgName & " " & dwgDiagStr
                            logText = logText & "  ERR DWG: " & dwgName & " (" & dwgChanges & ") " & dwgDiagStr & vbCrLf
                        End If
                        totalFiles = totalFiles + 1
                    End If
NextDwg:
                Next dwi
            Next i
        End If
    Next fi

    Application.StatusBar = False
    
    ' ========================================
    ' FINAL REPORT
    ' ========================================
    Dim summary As String
    summary = "DONE! " & OLD_REVISION & " -> " & NEW_REVISION & vbCrLf & vbCrLf
    summary = summary & "Stage 1:" & vbCrLf
    summary = summary & "  Deleted (pdf/bak/temp): " & delCount & vbCrLf
    summary = summary & "  Renamed (dirs+files): " & renCount & vbCrLf & vbCrLf
    summary = summary & "Stage 2:" & vbCrLf
    summary = summary & "  Docx folders: " & docxFolders.Count & vbCrLf
    summary = summary & "  Files processed: " & totalFiles & vbCrLf
    summary = summary & "  Content changes: " & totalChanges & vbCrLf
    summary = summary & "  PDF exported: " & totalPdf & vbCrLf & vbCrLf
    summary = summary & "Log:" & vbCrLf & logText
    
    MsgBox summary, vbInformation, "Result"
    
    ' Write full log to file
    Dim logFile As String
    logFile = mainFolder & "\UPDATE_LOG.txt"
    On Error Resume Next
    Dim ff As Long
    ff = FreeFile
    Open logFile For Output As #ff
    Print #ff, summary
    Close #ff
    On Error GoTo 0
    
    Set fso = Nothing
End Sub


' ============================================================
' STAGE 2: Process single docx file
' ============================================================

Function ProcessDocx(wordApp As Object, filePath As String, _
                     pdfFolder As String, _
                     oldRev As String, newRev As String, _
                     oldYear As String, newYear As String, _
                     oldFioRus As String, newFioRus As String, _
                     oldFioEng As String, newFioEng As String, _
                     isCAB As Boolean, isCAZ As Boolean, isPERM As Boolean, _
                     oldInv As String, newInv As String, _
                     newDateFull As String, newDateShort As String, _
                     oldPerm As String, newPerm As String, _
                     oldDateDDMM As String, newDateDDMM As String, _
                     ByRef pdfOk As Boolean, _
                     ByRef pdfErrMsg As String) As Long
    
    On Error GoTo ErrHandler
    pdfOk = False
    pdfErrMsg = ""
    
    Dim doc As Object
    Set doc = wordApp.Documents.Open(filePath, ReadOnly:=False)
    
    Dim changeCount As Long
    changeCount = 0
    
    ' 1. Revision in main text
    With doc.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = oldRev
        .Replacement.Text = newRev
        .Forward = True
        .Wrap = 1
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        Do While .Execute(Replace:=1)
            changeCount = changeCount + 1
        Loop
    End With
    
    ' 2. Headers/Footers
    Dim sec As Object
    For Each sec In doc.Sections
        Dim hf As Object
        
        Set hf = sec.Footers(1)
        If newYear <> "" And oldYear <> "" Then
            With hf.Range.Find
                .ClearFormatting
                .Replacement.ClearFormatting
                .Text = oldYear
                .Replacement.Text = newYear
                .Forward = True
                .Wrap = 1
                .Format = False
                .MatchCase = True
                Do While .Execute(Replace:=1)
                    changeCount = changeCount + 1
                Loop
            End With
        End If
        
        With hf.Range.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = oldRev
            .Replacement.Text = newRev
            .Forward = True
            .Wrap = 1
            .Format = False
            .MatchCase = True
            Do While .Execute(Replace:=1)
                changeCount = changeCount + 1
            Loop
        End With
        
        Set hf = sec.Headers(1)
        With hf.Range.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = oldRev
            .Replacement.Text = newRev
            .Forward = True
            .Wrap = 1
            .Format = False
            .MatchCase = True
            Do While .Execute(Replace:=1)
                changeCount = changeCount + 1
            Loop
        End With
    Next sec
    
    ' 2b. B-021: Replace dates in ALL story ranges (main doc + headers + footers)
    Dim storyRange As Object
    ' UL/CAA/CAB stamps have dates like 10.06.2024 and 10.06.24 that differ from
    ' the document date in Excel — use wildcard to catch any date with old year.
    If newDateFull <> "" And oldYear <> "" Then
        ' Wildcard: [0-9]{2}[.][0-9]{2}[.]YYYY -> newDateFull (validated)
        For Each storyRange In doc.StoryRanges
            changeCount = changeCount + ReplaceValidatedDateWildcard(storyRange, "[0-9]{2}[.][0-9]{2}[.]" & oldYear, newDateFull)
        Next storyRange
    End If
    If newDateShort <> "" And oldYear <> "" Then
        Dim oldYY2 As String
        oldYY2 = Right(oldYear, 2)
        ' Wildcard: [0-9]{2}[.][0-9]{2}[.]YY -> newDateShort (validated)
        For Each storyRange In doc.StoryRanges
            changeCount = changeCount + ReplaceValidatedDateWildcard(storyRange, "[0-9]{2}[.][0-9]{2}[.]" & oldYY2, newDateShort)
        Next storyRange
    End If

    ' 3. FIO via StoryRanges
    ' (storyRange declared above in section 2b)
    
    If newFioRus <> "" And oldFioRus <> "" Then
        For Each storyRange In doc.StoryRanges
            With storyRange.Find
                .ClearFormatting
                .Replacement.ClearFormatting
                .Text = oldFioRus
                .Replacement.Text = newFioRus
                .Forward = True
                .Wrap = 0
                .Format = False
                .MatchCase = False
                Do While .Execute(Replace:=1)
                    changeCount = changeCount + 1
                Loop
            End With
        Next storyRange
    End If
    
    If newFioEng <> "" And oldFioEng <> "" Then
        For Each storyRange In doc.StoryRanges
            With storyRange.Find
                .ClearFormatting
                .Replacement.ClearFormatting
                .Text = oldFioEng
                .Replacement.Text = newFioEng
                .Forward = True
                .Wrap = 0
                .Format = False
                .MatchCase = False
                Do While .Execute(Replace:=1)
                    changeCount = changeCount + 1
                Loop
            End With
        Next storyRange
    End If

    ' 3b. v23/v25: Surname-only search (B-016: catches split-run FIO)
    If newFioRus <> "" And oldFioRus <> "" Then
        Dim oldSurDocx As String
        Dim newSurDocx As String
        oldSurDocx = LastWord(oldFioRus)
        newSurDocx = LastWord(newFioRus)
        If oldSurDocx <> newSurDocx And oldSurDocx <> "" Then
            For Each storyRange In doc.StoryRanges
                With storyRange.Find
                    .ClearFormatting
                    .Replacement.ClearFormatting
                    .Text = oldSurDocx
                    .Replacement.Text = newSurDocx
                    .Forward = True
                    .Wrap = 0
                    .Format = False
                    .MatchCase = False
                    Do While .Execute(Replace:=1)
                        changeCount = changeCount + 1
                    Loop
                End With
            Next storyRange
        End If
    End If

    ' 3c. v23: Explicit footer/header table cell search for GIP
    ' StoryRanges may miss footer table cells in some Word layouts
    If newFioRus <> "" And oldFioRus <> "" Then
        Dim secGip As Object
        For Each secGip In doc.Sections
            Dim hfIdx As Long
            For hfIdx = 1 To 3  ' wdHeaderFooterPrimary=1, wdHeaderFooterFirstPage=2, wdHeaderFooterEvenPages=3
                On Error Resume Next
                Dim hfRng As Object
                Set hfRng = Nothing
                If hfIdx = 1 Then Set hfRng = secGip.Footers(1).Range
                If hfIdx = 2 Then Set hfRng = secGip.Footers(2).Range
                If hfIdx = 3 Then Set hfRng = secGip.Headers(1).Range
                If Err.Number <> 0 Or hfRng Is Nothing Then
                    Err.Clear
                    GoTo NextHF
                End If
                Err.Clear
                On Error GoTo ErrHandler
                Dim hfTbl As Object
                For Each hfTbl In hfRng.Tables
                    Dim hfCel As Object
                    For Each hfCel In hfTbl.Range.Cells
                        Dim hfCelRng As Object
                        Set hfCelRng = hfCel.Range
                        hfCelRng.MoveEnd 1, -1
                        ' Full FIO
                        changeCount = changeCount + ReplaceInCellRange(hfCelRng, oldFioRus, newFioRus)
                        ' Surname only
                        If oldSurDocx <> newSurDocx And oldSurDocx <> "" Then
                            changeCount = changeCount + ReplaceInCellRange(hfCelRng, oldSurDocx, newSurDocx)
                        End If
                    Next hfCel
                Next hfTbl
NextHF:
            Next hfIdx
        Next secGip
    End If

    ' 4. CAB stamp
    If isCAB And newInv <> "" Then
        Dim secIdx As Long
        For secIdx = 1 To doc.Sections.Count
            Dim ftr As Object
            Set ftr = doc.Sections(secIdx).Footers(1)
            
            If ftr.Range.Tables.Count >= 3 Then
                Dim tblRev As Object
                Set tblRev = ftr.Range.Tables(2)
                
                If newPerm <> "" Then
                    On Error Resume Next
                    tblRev.cell(1, 4).Range.Text = newPerm
                    If Err.Number = 0 Then changeCount = changeCount + 1
                    Err.Clear
                    On Error GoTo ErrHandler
                End If
                
                If newDateShort <> "" Then
                    On Error Resume Next
                    tblRev.cell(1, 5).Range.Text = newDateShort
                    If Err.Number = 0 Then changeCount = changeCount + 1
                    Err.Clear
                    On Error GoTo ErrHandler
                End If
                
                Dim tblInv As Object
                Set tblInv = ftr.Range.Tables(3)
                
                On Error Resume Next
                Dim currentInv As String
                currentInv = Trim(tblInv.cell(3, 2).Range.Text)
                If Len(currentInv) > 1 Then
                    currentInv = Left(currentInv, Len(currentInv) - 2)
                End If
                
                tblInv.cell(1, 2).Range.Text = currentInv
                If Err.Number = 0 Then changeCount = changeCount + 1
                Err.Clear
                
                tblInv.cell(3, 2).Range.Text = newInv
                If Err.Number = 0 Then changeCount = changeCount + 1
                Err.Clear
                
                If newDateFull <> "" Then
                    tblInv.cell(2, 2).Range.Text = newDateFull
                    If Err.Number = 0 Then changeCount = changeCount + 1
                    Err.Clear
                End If
                On Error GoTo ErrHandler
            End If
        Next secIdx
    End If
    
    ' 5. CAZ: permission, Cyrillic rev, date
    If isCAZ And oldPerm <> "" And newPerm <> "" Then
        For Each storyRange In doc.StoryRanges
            With storyRange.Find
                .ClearFormatting
                .Replacement.ClearFormatting
                .Text = oldPerm
                .Replacement.Text = newPerm
                .Forward = True
                .Wrap = 0
                .Format = False
                .MatchCase = True
                Do While .Execute(Replace:=1)
                    changeCount = changeCount + 1
                Loop
            End With
        Next storyRange
    End If
    
    If isCAZ Then
        Dim cyrOldRev As String
        Dim cyrNewRev As String
        cyrOldRev = ChrW(1057) & Mid(oldRev, 2)
        cyrNewRev = ChrW(1057) & Mid(newRev, 2)
        For Each storyRange In doc.StoryRanges
            With storyRange.Find
                .ClearFormatting
                .Replacement.ClearFormatting
                .Text = cyrOldRev
                .Replacement.Text = cyrNewRev
                .Forward = True
                .Wrap = 0
                .Format = False
                .MatchCase = True
                Do While .Execute(Replace:=1)
                    changeCount = changeCount + 1
                Loop
            End With
        Next storyRange
    End If
    
    If isCAZ And oldDateDDMM <> "" And newDateDDMM <> "" And oldDateDDMM <> newDateDDMM Then
        For Each storyRange In doc.StoryRanges
            With storyRange.Find
                .ClearFormatting
                .Replacement.ClearFormatting
                .Text = oldDateDDMM
                .Replacement.Text = newDateDDMM
                .Forward = True
                .Wrap = 0
                .Format = False
                .MatchCase = True
                Do While .Execute(Replace:=1)
                    changeCount = changeCount + 1
                Loop
            End With
        Next storyRange
    End If
    
    If isCAZ And newYear <> "" And oldYear <> "" Then
        With doc.Content.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = oldYear
            .Replacement.Text = newYear
            .Forward = True
            .Wrap = 1
            .Format = False
            .MatchCase = True
            Do While .Execute(Replace:=1)
                changeCount = changeCount + 1
            Loop
        End With
    End If
    
    ' 5b. PERM: permission file (005-24.docx) - replace all data
    If isPERM Then
        ' Replace old permission number (005-24 -> 99-26)
        If oldPerm <> "" And newPerm <> "" Then
            With doc.Content.Find
                .ClearFormatting: .Replacement.ClearFormatting
                .Text = oldPerm: .Replacement.Text = newPerm
                .Forward = True: .Wrap = 1: .Format = False: .MatchCase = True
                Do While .Execute(Replace:=1): changeCount = changeCount + 1: Loop
            End With
        End If
        
        ' Replace old inv number (18736 -> 26085)
        If oldInv <> "" And newInv <> "" Then
            With doc.Content.Find
                .ClearFormatting: .Replacement.ClearFormatting
                .Text = oldInv: .Replacement.Text = newInv
                .Forward = True: .Wrap = 1: .Format = False: .MatchCase = True
                Do While .Execute(Replace:=1): changeCount = changeCount + 1: Loop
            End With
        End If
        
        ' Replace Cyrillic revision (С02 -> С03)
        Dim cyrOldPerm As String
        Dim cyrNewPerm As String
        cyrOldPerm = ChrW(1057) & Mid(oldRev, 2)
        cyrNewPerm = ChrW(1057) & Mid(newRev, 2)
        With doc.Content.Find
            .ClearFormatting: .Replacement.ClearFormatting
            .Text = cyrOldPerm: .Replacement.Text = cyrNewPerm
            .Forward = True: .Wrap = 1: .Format = False: .MatchCase = True
            Do While .Execute(Replace:=1): changeCount = changeCount + 1: Loop
        End With
        
        ' Replace date DD.MM.YY (short) - exact first, then wildcard for ANY date with old year
        ' B-020: stamp date (10.06.24) differs from doc date (07.05.24) - use wildcard
        If newDateShort <> "" And oldYear <> "" Then
            Dim oldShortDate As String
            Dim oldYY As String
            oldYY = Right(oldYear, 2)
            If oldDateDDMM <> "" Then
                oldShortDate = oldDateDDMM & oldYY
                If oldShortDate <> newDateShort Then
                    With doc.Content.Find
                        .ClearFormatting: .Replacement.ClearFormatting
                        .Text = oldShortDate: .Replacement.Text = newDateShort
                        .Forward = True: .Wrap = 1: .Format = False: .MatchCase = True
                        .Execute Replace:=2
                    End With
                    changeCount = changeCount + 1
                End If
            End If
            ' B-020: wildcard DD.MM.YY with validation (avoid false matches)
            changeCount = changeCount + ReplaceValidatedDateWildcard(doc.Content, "[0-9]{2}[.][0-9]{2}[.]" & oldYY, newDateShort)
        End If
        
        ' Replace date DD.MM.YYYY: any date with old year -> new date from col M
        ' Using wildcard to match XX.XX.2024 -> 19.02.2026
        If newDateFull <> "" And oldYear <> "" And newYear <> "" Then
            ' Use wildcard DD.MM.YYYY with validation
            changeCount = changeCount + ReplaceValidatedDateWildcard(doc.Content, "[0-9]{2}[.][0-9]{2}[.]" & oldYear, newDateFull)
        End If
        
        ' Replace GIP surname only (B-016: via LastWord)
        If oldFioRus <> "" And newFioRus <> "" Then
            Dim oldSurname As String
            Dim newSurname As String
            oldSurname = LastWord(oldFioRus)
            newSurname = LastWord(newFioRus)
            
            If oldSurname <> newSurname Then
                With doc.Content.Find
                    .ClearFormatting
                    .Replacement.ClearFormatting
                    .Text = oldSurname
                    .Replacement.Text = newSurname
                    .Forward = True
                    .Wrap = 1
                    .Format = False
                    .MatchCase = False
                    Do While .Execute(Replace:=1)
                        changeCount = changeCount + 1
                    Loop
                End With
            End If
        End If
        
        ' Replace GIP eng surname (B-016: via LastWord)
        If oldFioEng <> "" And newFioEng <> "" Then
            Dim oldSurnameEng As String
            Dim newSurnameEng As String
            oldSurnameEng = LastWord(oldFioEng)
            newSurnameEng = LastWord(newFioEng)
            If oldSurnameEng <> newSurnameEng Then
                With doc.Content.Find
                    .ClearFormatting
                    .Replacement.ClearFormatting
                    .Text = oldSurnameEng
                    .Replacement.Text = newSurnameEng
                    .Forward = True
                    .Wrap = 1
                    .Format = False
                    .MatchCase = False
                    Do While .Execute(Replace:=1)
                        changeCount = changeCount + 1
                    Loop
                End With
            End If
        End If
    End If
    
    ' 6. Save (for PERM: SaveAs with new name)
    Dim permRenamed As Boolean
    permRenamed = False
    Dim newPermPath As String
    newPermPath = ""

    Dim actualDocPath As String  ' path of the saved file (for logo replacement)
    actualDocPath = filePath

    If isPERM And oldPerm <> "" And newPerm <> "" Then
        ' SaveAs with new permission name
        Dim docFolder As String
        docFolder = Left(filePath, InStrRev(filePath, "\"))
        newPermPath = docFolder & newPerm & ".docx"
        On Error Resume Next
        doc.SaveAs2 fileName:=newPermPath, FileFormat:=12
        If Err.Number = 0 Then
            permRenamed = True
            actualDocPath = newPermPath
        Else
            Err.Clear
            doc.Save
        End If
        On Error GoTo ErrHandler
    Else
        doc.Save
    End If

    ' 6b. v25 B-017: Replace logo in CAB/CAA/CAZ files
    ' Check if this file type needs logo update
    Dim needsLogo As Boolean
    needsLogo = False
    If newYear <> "" Then
        Dim fnUpper As String
        fnUpper = UCase(filePath)
        If InStr(fnUpper, "-CAB") > 0 Or InStr(fnUpper, "-CAA") > 0 Or InStr(fnUpper, "-CAZ") > 0 Then
            needsLogo = True
        End If
    End If

    Dim logoOk As Boolean
    logoOk = False
    If needsLogo Then
        ' Must close doc before modifying file on disk
        doc.Close SaveChanges:=False
        Set doc = Nothing
        logoOk = ReplaceLogoInDocx(actualDocPath, newYear)
        ' Reopen doc for PDF export
        Set doc = wordApp.Documents.Open(actualDocPath, ReadOnly:=False)
    End If

    ' 7. Export PDF (use new name if PERM was renamed)
    Dim pdfName As String
    If permRenamed Then
        pdfName = newPerm & ".pdf"
    Else
        Dim docName As String
        docName = Mid(actualDocPath, InStrRev(actualDocPath, "\") + 1)
        pdfName = Left(docName, InStrRev(docName, ".") - 1) & ".pdf"
    End If
    Dim pdfPath As String
    pdfPath = pdfFolder & pdfName

    On Error Resume Next
    doc.SaveAs2 fileName:=pdfPath, FileFormat:=17
    If Err.Number = 0 Then
        pdfOk = True
    Else
        pdfErrMsg = "PDF Err " & Err.Number
        pdfOk = False
        Err.Clear
    End If
    On Error GoTo ErrHandler

    doc.Close SaveChanges:=False
    Set doc = Nothing

    ' Delete old PERM file after close (005-24.docx)
    If permRenamed And newPermPath <> filePath Then
        On Error Resume Next
        If Dir(filePath) <> "" Then Kill filePath
        Err.Clear
        On Error GoTo 0
    End If

    ' v25: append logo status to changeCount context (caller reads pdfErrMsg)
    If logoOk Then
        If pdfErrMsg = "" Then pdfErrMsg = "[logo updated]" Else pdfErrMsg = pdfErrMsg & " [logo updated]"
    End If

    ProcessDocx = changeCount
    Exit Function
    
ErrHandler:
    On Error Resume Next
    If Not doc Is Nothing Then
        doc.Close SaveChanges:=False
    End If
    pdfOk = False
    pdfErrMsg = "Err " & Err.Number
    ProcessDocx = -1
End Function


' ============================================================
' DWG PROCESSING (v20+, B-009 fix in v22)
' ============================================================

' Simple string replace via Split/Join (used for DWG text entities)
Function ReplaceAllStr(text As String, oldStr As String, newStr As String) As String
    If oldStr = "" Or oldStr = newStr Then
        ReplaceAllStr = text
        Exit Function
    End If
    ReplaceAllStr = Join(Split(text, oldStr), newStr)
End Function


' v24: Apply all DWG text replacements in one place (reduces code duplication)
' cyrOldRev = Cyrillic C variant of oldRev (e.g. Chr(1057)+"02" for "C02")
Private Function ApplyDwgReplacements(text As String, _
    oldRev As String, newRev As String, cyrOldRev As String, _
    oldPerm As String, newPerm As String, _
    oldInv As String, newInv As String, _
    oldDateFull As String, newDateFull As String, _
    oldDateShort As String, newDateShort As String, _
    oldYear As String, newYear As String, _
    prevInv As String) As String

    Dim r As String
    r = text

    ' Rev replacement: Latin C + Cyrillic C variants
    If oldRev <> newRev Then
        r = ReplaceAllStr(r, oldRev, newRev)
        If cyrOldRev <> oldRev Then r = ReplaceAllStr(r, cyrOldRev, newRev)
    End If

    ' PERM number
    If oldPerm <> "" And newPerm <> "" And oldPerm <> newPerm Then _
        r = ReplaceAllStr(r, oldPerm, newPerm)

    ' Inventory numbers
    If oldInv <> "" And newInv <> "" And oldInv <> newInv Then _
        r = ReplaceAllStr(r, oldInv, newInv)

    ' Dates
    If oldDateFull <> "" And newDateFull <> "" And oldDateFull <> newDateFull Then _
        r = ReplaceAllStr(r, oldDateFull, newDateFull)
    If oldDateShort <> "" And newDateShort <> "" And oldDateShort <> newDateShort Then _
        r = ReplaceAllStr(r, oldDateShort, newDateShort)

    ' Year
    If oldYear <> "" And newYear <> "" And oldYear <> newYear Then _
        r = ReplaceAllStr(r, oldYear, newYear)

    ' prevInv -> oldInv (Replace arch No)
    If prevInv <> "" And oldInv <> "" And prevInv <> oldInv Then _
        r = ReplaceAllStr(r, prevInv, oldInv)

    ApplyDwgReplacements = r
End Function


' Process a single DWG file via AutoCAD COM Automation
' Handles: ACDBTEXT, ACDBMTEXT, ACDBBLOCKREFERENCE (ATTRIB), ACDBTABLE
' v22: prevInv scans ALL spaces; v24: + Cyrillic C fix, nested blocks, ACDBTEXT prevInv
' v25: B-019 perf (skip Model when CLD layouts exist), B-018 PDF export
' Returns number of changes, or -1 on error
Function ProcessDwg(acadApp As Object, filePath As String, _
                    oldRev As String, newRev As String, _
                    oldInv As String, newInv As String, _
                    oldPerm As String, newPerm As String, _
                    oldDateFull As String, newDateFull As String, _
                    oldDateShort As String, newDateShort As String, _
                    oldYear As String, newYear As String, _
                    Optional ByRef dwgDiag As String = "", _
                    Optional pdfFolder As String = "", _
                    Optional forceModelScan As Boolean = False) As Long

    On Error GoTo DwgErrHandler
    dwgDiag = ""
    AppendTrace "ProcessDwg open: " & filePath

    ' v24: Build Cyrillic C variant of oldRev for B-012 fix
    Dim cyrOldRev As String
    cyrOldRev = Replace(oldRev, "C", ChrW(1057))

    ' v23: Suppress proxy/SHX dialogs before opening
    On Error Resume Next
    acadApp.SetSystemVariable "PROXYNOTICE", 0
    acadApp.SetSystemVariable "FILEDIA", 0
    Err.Clear
    On Error GoTo DwgErrHandler

    Dim dwgDoc As Object
    Set dwgDoc = Nothing
    On Error Resume Next
    Set dwgDoc = acadApp.Documents.Open(filePath, False)  ' ReadOnly=False
    If Err.Number <> 0 Or dwgDoc Is Nothing Then
        dwgDiag = "[OPEN ERR " & Err.Number & ": " & Err.Description & "]"
        AppendTrace "ProcessDwg open error: " & dwgDiag
        Err.Clear
        On Error GoTo 0
        ProcessDwg = -1
        Exit Function
    End If
    AppendTrace "ProcessDwg opened"
    On Error GoTo DwgErrHandler

    Dim changeCount As Long
    changeCount = 0

    ' ================================================
    ' v25 B-019: PERFORMANCE — detect if CLD layouts exist
    ' If yes, skip Model space (50k+ objects) in all loops
    ' Title block is always in Paper Space (Layouts), never in Model
    ' ================================================
    Dim hasCLDLayouts As Boolean
    hasCLDLayouts = False
    Dim chkLo As Object
    For Each chkLo In dwgDoc.Layouts
        If chkLo.Name <> "Model" And _
           UCase(Left(chkLo.Name, 3)) = "CLD" Then
            hasCLDLayouts = True
            Exit For
        End If
    Next chkLo

    ' ================================================
    ' STEP 1: Auto-detect prevInv (Replace arch.No)
    ' v22: scan ALL spaces (Model + all Layouts)
    ' v24: also scan ACDBTEXT/ACDBMTEXT + inner block defs
    ' ================================================
    Dim prevInv As String
    prevInv = ""

    Dim layout As Object
    Dim blk As Object
    Dim entIdx As Long
    Dim ent As Object
    Dim attribs As Variant
    Dim hasOldInv As Boolean
    Dim ai As Long
    Dim sv2 As String

    ' --- STEP 1A: scan top-level block reference attributes ---
    For Each layout In dwgDoc.Layouts
        ' v25 B-019: skip Model if CLD layouts exist
        If hasCLDLayouts And layout.Name = "Model" And (Not forceModelScan) Then GoTo NextLayout_1A
        Set blk = layout.Block
        For entIdx = 0 To blk.Count - 1
            Set ent = blk.Item(entIdx)
            If UCase(ent.ObjectName) = "ACDBBLOCKREFERENCE" Then
                On Error Resume Next
                attribs = ent.GetAttributes
                If Err.Number = 0 Then
                    hasOldInv = False
                    For ai = LBound(attribs) To UBound(attribs)
                        If Trim(attribs(ai).TextString) = oldInv Then
                            hasOldInv = True
                            Exit For
                        End If
                    Next ai
                    If hasOldInv And prevInv = "" Then
                        For ai = LBound(attribs) To UBound(attribs)
                            sv2 = Trim(attribs(ai).TextString)
                            If sv2 <> oldInv And sv2 <> newInv And sv2 <> "" Then
                                If IsAllDigits(sv2) And Len(sv2) >= 4 And Len(sv2) <= 8 Then
                                    prevInv = sv2
                                    Exit For
                                End If
                            End If
                        Next ai
                    End If
                End If
                Err.Clear
                On Error GoTo DwgErrHandler

                ' v24 B-013: also check INNER block definition entities
                If prevInv = "" Then
                    On Error Resume Next
                    Dim innerBlkName As String
                    innerBlkName = ent.Name
                    If Err.Number = 0 And innerBlkName <> "" Then
                        Dim innerBlk As Object
                        Set innerBlk = dwgDoc.Blocks.Item(innerBlkName)
                        If Err.Number = 0 And Not innerBlk Is Nothing Then
                            Dim innerIdx As Long
                            For innerIdx = 0 To innerBlk.Count - 1
                                Dim innerEnt As Object
                                Set innerEnt = innerBlk.Item(innerIdx)
                                If UCase(innerEnt.ObjectName) = "ACDBBLOCKREFERENCE" Then
                                    Dim innerAttribs As Variant
                                    innerAttribs = innerEnt.GetAttributes
                                    If Err.Number = 0 Then
                                        hasOldInv = False
                                        For ai = LBound(innerAttribs) To UBound(innerAttribs)
                                            If Trim(innerAttribs(ai).TextString) = oldInv Then
                                                hasOldInv = True
                                                Exit For
                                            End If
                                        Next ai
                                        If hasOldInv And prevInv = "" Then
                                            For ai = LBound(innerAttribs) To UBound(innerAttribs)
                                                sv2 = Trim(innerAttribs(ai).TextString)
                                                If sv2 <> oldInv And sv2 <> newInv And sv2 <> "" Then
                                                    If IsAllDigits(sv2) And Len(sv2) >= 4 And Len(sv2) <= 8 Then
                                                        prevInv = sv2
                                                        Exit For
                                                    End If
                                                End If
                                            Next ai
                                        End If
                                    End If
                                    Err.Clear
                                End If
                            Next innerIdx
                        End If
                    End If
                    Err.Clear
                    On Error GoTo DwgErrHandler
                End If
            End If
            If prevInv <> "" Then Exit For
        Next entIdx
        If prevInv <> "" Then Exit For
NextLayout_1A:
    Next layout

    ' --- STEP 1B: v24 fallback — scan ACDBTEXT/ACDBMTEXT for prevInv ---
    If prevInv = "" Then
        Dim foundOldInvText As Boolean
        foundOldInvText = False
        Dim candidateInv As String
        candidateInv = ""

        For Each layout In dwgDoc.Layouts
            ' v25 B-019: skip Model if CLD layouts exist
            If hasCLDLayouts And layout.Name = "Model" And (Not forceModelScan) Then GoTo NextLayout_1B
            Set blk = layout.Block
            For entIdx = 0 To blk.Count - 1
                Set ent = blk.Item(entIdx)
                Dim entNameScan As String
                entNameScan = UCase(ent.ObjectName)
                If entNameScan = "ACDBTEXT" Or entNameScan = "ACDBMTEXT" Then
                    Dim txtScan As String
                    txtScan = Trim(ent.TextString)
                    If txtScan = oldInv Then
                        foundOldInvText = True
                    ElseIf IsAllDigits(txtScan) And Len(txtScan) >= 4 And Len(txtScan) <= 8 Then
                        If txtScan <> oldInv And txtScan <> newInv And _
                           txtScan <> oldYear And txtScan <> newYear Then
                            If candidateInv = "" Then candidateInv = txtScan
                        End If
                    End If
                End If
            Next entIdx
NextLayout_1B:
        Next layout

        If foundOldInvText And candidateInv <> "" Then
            prevInv = candidateInv
        End If
    End If

    ' v24: diagnostic output
    If prevInv <> "" Then
        dwgDiag = "[prevInv=" & prevInv & "]"
    Else
        dwgDiag = "[prevInv=?]"
    End If

    ' ================================================
    ' STEP 2: Replace text in ALL spaces (Model + Layouts)
    ' v24: uses ApplyDwgReplacements helper + nested blocks
    ' v25 B-019: skip Model when CLD layouts exist
    ' ================================================
    Dim entType As String
    Dim oldVal As String
    Dim newVal As String

    For Each layout In dwgDoc.Layouts
        ' v25 B-019: skip Model if CLD layouts exist
        If hasCLDLayouts And layout.Name = "Model" And (Not forceModelScan) Then GoTo NextLayout_2
        Set blk = layout.Block
        For entIdx = 0 To blk.Count - 1
            Set ent = blk.Item(entIdx)
            entType = UCase(ent.ObjectName)

            Select Case entType
            Case "ACDBTEXT", "ACDBMTEXT"
                oldVal = ent.TextString
                newVal = ApplyDwgReplacements(oldVal, oldRev, newRev, cyrOldRev, _
                    oldPerm, newPerm, oldInv, newInv, _
                    oldDateFull, newDateFull, oldDateShort, newDateShort, _
                    oldYear, newYear, prevInv)
                If newVal <> oldVal Then
                    ent.TextString = newVal
                    changeCount = changeCount + 1
                End If

            Case "ACDBBLOCKREFERENCE"
                ' Process attributes of this block reference
                On Error Resume Next
                attribs = ent.GetAttributes
                If Err.Number = 0 Then
                    For ai = LBound(attribs) To UBound(attribs)
                        oldVal = attribs(ai).TextString
                        newVal = ApplyDwgReplacements(oldVal, oldRev, newRev, cyrOldRev, _
                            oldPerm, newPerm, oldInv, newInv, _
                            oldDateFull, newDateFull, oldDateShort, newDateShort, _
                            oldYear, newYear, prevInv)
                        If newVal <> oldVal Then
                            attribs(ai).TextString = newVal
                            changeCount = changeCount + 1
                        End If
                    Next ai
                End If
                Err.Clear
                On Error GoTo DwgErrHandler

                ' v24 B-013: process entities inside inner block definition
                On Error Resume Next
                innerBlkName = ent.Name
                If Err.Number = 0 And innerBlkName <> "" Then
                    Set innerBlk = dwgDoc.Blocks.Item(innerBlkName)
                    If Err.Number = 0 And Not innerBlk Is Nothing Then
                        For innerIdx = 0 To innerBlk.Count - 1
                            Set innerEnt = innerBlk.Item(innerIdx)
                            Dim innerEntType As String
                            innerEntType = UCase(innerEnt.ObjectName)

                            If innerEntType = "ACDBTEXT" Or innerEntType = "ACDBMTEXT" Then
                                oldVal = innerEnt.TextString
                                newVal = ApplyDwgReplacements(oldVal, oldRev, newRev, cyrOldRev, _
                                    oldPerm, newPerm, oldInv, newInv, _
                                    oldDateFull, newDateFull, oldDateShort, newDateShort, _
                                    oldYear, newYear, prevInv)
                                If newVal <> oldVal Then
                                    innerEnt.TextString = newVal
                                    changeCount = changeCount + 1
                                End If

                            ElseIf innerEntType = "ACDBBLOCKREFERENCE" Then
                                ' One more level of nesting (inner-inner block)
                                Dim innerAttribs2 As Variant
                                innerAttribs2 = innerEnt.GetAttributes
                                If Err.Number = 0 Then
                                    Dim ai2 As Long
                                    For ai2 = LBound(innerAttribs2) To UBound(innerAttribs2)
                                        oldVal = innerAttribs2(ai2).TextString
                                        newVal = ApplyDwgReplacements(oldVal, oldRev, newRev, cyrOldRev, _
                                            oldPerm, newPerm, oldInv, newInv, _
                                            oldDateFull, newDateFull, oldDateShort, newDateShort, _
                                            oldYear, newYear, prevInv)
                                        If newVal <> oldVal Then
                                            innerAttribs2(ai2).TextString = newVal
                                            changeCount = changeCount + 1
                                        End If
                                    Next ai2
                                End If
                                Err.Clear

                            ElseIf innerEntType = "ACDBTABLE" Then
                                Dim nR2 As Long, nC2 As Long
                                nR2 = innerEnt.Rows
                                nC2 = innerEnt.Columns
                                If Err.Number = 0 Then
                                    Dim tr2 As Long, tc2 As Long
                                    For tr2 = 0 To nR2 - 1
                                        For tc2 = 0 To nC2 - 1
                                            Dim cellT2 As String
                                            cellT2 = ""
                                            cellT2 = innerEnt.GetText(tr2, tc2)
                                            If Err.Number <> 0 Then
                                                Err.Clear
                                                GoTo NextInnerTableCell
                                            End If
                                            If cellT2 <> "" Then
                                                newVal = ApplyDwgReplacements(cellT2, oldRev, newRev, cyrOldRev, _
                                                    oldPerm, newPerm, oldInv, newInv, _
                                                    oldDateFull, newDateFull, oldDateShort, newDateShort, _
                                                    oldYear, newYear, prevInv)
                                                If newVal <> cellT2 Then
                                                    innerEnt.SetText tr2, tc2, newVal
                                                    changeCount = changeCount + 1
                                                End If
                                            End If
NextInnerTableCell:
                                        Next tc2
                                    Next tr2
                                End If
                                Err.Clear
                            End If
                        Next innerIdx
                    End If
                End If
                Err.Clear
                On Error GoTo DwgErrHandler

            Case "ACDBTABLE"
                On Error Resume Next
                Dim nRows As Long, nCols As Long
                nRows = ent.Rows
                nCols = ent.Columns
                If Err.Number = 0 Then
                    Dim tr As Long, tc As Long
                    For tr = 0 To nRows - 1
                        For tc = 0 To nCols - 1
                            Dim cellText As String
                            cellText = ""
                            cellText = ent.GetText(tr, tc)
                            If Err.Number <> 0 Then
                                Err.Clear
                                GoTo NextTableCell
                            End If
                            If cellText <> "" Then
                                newVal = ApplyDwgReplacements(cellText, oldRev, newRev, cyrOldRev, _
                                    oldPerm, newPerm, oldInv, newInv, _
                                    oldDateFull, newDateFull, oldDateShort, newDateShort, _
                                    oldYear, newYear, prevInv)
                                If newVal <> cellText Then
                                    ent.SetText tr, tc, newVal
                                    changeCount = changeCount + 1
                                End If
                            End If
NextTableCell:
                        Next tc
                    Next tr
                End If
                Err.Clear
                On Error GoTo DwgErrHandler

            End Select
        Next entIdx
NextLayout_2:
    Next layout

    ' ================================================
    ' v25 B-018: PDF EXPORT from DWG
    ' Collect CLD-named layouts; if none, use Model
    ' Uses saved plot settings in each layout (DWG To PDF.pc3)
    ' ================================================
    Dim pdfDwgCount As Long
    pdfDwgCount = 0

    If pdfFolder <> "" Then
        Dim pdfLayouts As New Collection
        Dim lo As Object

        ' Collect all CLD-layouts
        For Each lo In dwgDoc.Layouts
            If lo.Name <> "Model" And _
               UCase(Left(lo.Name, 3)) = "CLD" Then
                pdfLayouts.Add lo
            End If
        Next lo

        ' If no CLD-layouts — use Model
        If pdfLayouts.Count = 0 Then
            For Each lo In dwgDoc.Layouts
                If lo.Name = "Model" Then
                    pdfLayouts.Add lo
                    Exit For
                End If
            Next lo
        End If

        ' Extract base name from file path (without path and extension)
        Dim dwgBaseName As String
        dwgBaseName = Mid(filePath, InStrRev(filePath, "\") + 1)
        If InStrRev(dwgBaseName, ".") > 0 Then
            dwgBaseName = Left(dwgBaseName, InStrRev(dwgBaseName, ".") - 1)
        End If

        ' Print each layout to PDF
        Dim pli As Long
        For pli = 1 To pdfLayouts.Count
            Set lo = pdfLayouts(pli)
            Dim pdfFileName As String
            If lo.Name = "Model" Then
                pdfFileName = pdfFolder & dwgBaseName & ".pdf"
            Else
                pdfFileName = pdfFolder & dwgBaseName & "_" & lo.Name & ".pdf"
            End If

            On Error Resume Next
            ' Activate layout and use saved plot settings
            dwgDoc.ActiveLayout = lo
            acadApp.SetSystemVariable "BACKGROUNDPLOT", 0

            dwgDoc.Plot.PlotToFile pdfFileName, ""
            If Err.Number <> 0 Then
                Err.Clear
            Else
                pdfDwgCount = pdfDwgCount + 1
            End If
            On Error GoTo DwgErrHandler
        Next pli

        If pdfDwgCount > 0 Then
            dwgDiag = dwgDiag & " +PDF(" & pdfDwgCount & ")"
        End If
    End If

    ' Save DWG
    AppendTrace "ProcessDwg save"
    dwgDoc.Save
    AppendTrace "ProcessDwg close"
    dwgDoc.Close
    Set dwgDoc = Nothing
    AppendTrace "ProcessDwg done"

    ProcessDwg = changeCount
    Exit Function

DwgErrHandler:
    dwgDiag = dwgDiag & " [ERR " & Err.Number & ": " & Err.Description & "]"
    AppendTrace "ProcessDwg handler: " & dwgDiag
    On Error Resume Next
    If Not dwgDoc Is Nothing Then
        ' Avoid extra hangs in handler close on some Civil3D versions
        Set dwgDoc = Nothing
    End If
    On Error GoTo 0
    ProcessDwg = -1
End Function
