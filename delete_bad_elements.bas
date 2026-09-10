'==============================================================
' delete_bad_elements.bas
'--------------------------------------------------------------
' Deletes every element Nastran rejected for bad geometry.
' Reads the element IDs from the .f06 - nothing hardcoded.
'
' WHY DELETE RATHER THAN REPAIR
'   Of the 887 rejected elements in the current run:
'     381 have TWO or more zero-length edges -> zero area,
'         no membrane stiffness, no load path contribution
'     262 have ONE zero-length edge -> geometrically triangles
'         with negligible area
'     158 have real area but folded/crossed connectivity
'   Removing 887 of 608,526 elements is 0.15% of the mesh.
'   This is a DIAGNOSTIC run to get past SEMG, not a valid
'   analysis. Remesh properly before trusting any stress.
'
' SAFETY
'   - Refuses to run unless BACKUP_DONE = 1
'   - Counts elements before and after
'   - Verifies by re-fetching deleted IDs; any that still exist
'     are reported
'   - DRY_RUN = 1 reports what it would delete and stops
'==============================================================

Const DRY_RUN     As Integer = 1
Const BACKUP_DONE As Integer = 0

Const SCAN_DIR    As String = "C:\Users\nikhil.mohan\Desktop\Analysis testing\"
Const OUT_DIR     As String = "C:\Users\nikhil.mohan\Desktop\Analysis testing\audit\"

Const FT_ELEM     As Integer = 8
Const MAX_IDS     As Long = 20000

Dim femap As Object
Dim logF As Integer
Dim badID() As Long
Dim nBad As Long

Sub Main

    Dim activeLog As String

    Set femap = GetObject(, "femap.model")

    If EnsureDir(OUT_DIR) = 0 Then
        femap.feAppMessage 2, "Cannot create " & OUT_DIR
        End
    End If

    activeLog = FindLatestF06(SCAN_DIR)

    If activeLog = "" Then
        femap.feAppMessage 2, "No .f06 found in " & SCAN_DIR
        End
    End If

    If DRY_RUN = 0 Then
        If BACKUP_DONE <> 1 Then
            femap.feAppMessage 2, "REFUSING TO DELETE. Save a backup, then set BACKUP_DONE = 1."
            End
        End If
    End If

    logF = FreeFile
    Open OUT_DIR & "delete_bad.log" For Output As #logF

    WL "=== delete_bad_elements ==="
    WL "DRY_RUN = " & DRY_RUN
    WL "source  = " & activeLog
    WL ""

    Call CollectIDs(activeLog)
    Call DoDelete

    Close #logF

    femap.feAppMessage 0, "Done. See " & OUT_DIR & "delete_bad.log"

End Sub

Sub CollectIDs(pathStr As String)

    Dim f As Integer
    Dim ln As String
    Dim prev As String
    Dim eid As Long
    Dim k As Long
    Dim dup As Integer
    Dim fO As Integer

    ReDim badID(0 To MAX_IDS)
    nBad = 0
    prev = ""

    f = FreeFile
    Open pathStr For Input As #f

    Do While Not EOF(f)

        Line Input #f, ln
        eid = -1

        If InStr(ln, "ILLEGAL GEOMETRY FOR QUAD4/QUADR ELEMENT WITH ID") > 0 Then
            eid = GrabID(ln)
        ElseIf InStr(ln, "HAS AN INTERIOR ANGLE POSSIBLY GREATER") > 0 Then
            eid = GrabID(prev)
        ElseIf InStr(ln, "SPECIFIED FOR ELEMENT WITH ID") > 0 Then
            If InStr(prev, "INAPPROPRIATE GEOMETRY OR INCORRECT MATERIAL DATA") > 0 Then
                eid = GrabID(ln)
            End If
        End If

        If eid > 0 Then
            dup = 0
            For k = 0 To nBad - 1
                If badID(k) = eid Then
                    dup = 1
                End If
            Next k
            If dup = 0 Then
                If nBad <= MAX_IDS Then
                    badID(nBad) = eid
                    nBad = nBad + 1
                End If
            End If
        End If

        prev = ln

    Loop

    Close #f

    WL "elements to delete: " & nBad

    ' always dump the list, so a manual paste is possible if the
    ' API call turns out to be wrong for this build
    fO = FreeFile
    Open OUT_DIR & "delete_ids.txt" For Output As #fO
    For k = 0 To nBad - 1
        Print #fO, badID(k)
    Next k
    Close #fO

    WL "id list also written to delete_ids.txt"
    WL ""

End Sub

Sub DoDelete

    Dim s As Object
    Dim oEl As Object
    Dim i As Long
    Dim rc As Long
    Dim stillThere As Long
    Dim missingBefore As Long

    If nBad = 0 Then
        WL "Nothing to delete."
        Exit Sub
    End If

    Set s = femap.feSet                                   'VERIFIED
    s.Clear

    missingBefore = 0

    For i = 0 To nBad - 1
        Set oEl = femap.feElem                            'VERIFIED
        If oEl.Get(badID(i)) = -1 Then                    'VERIFIED
            s.Add badID(i)                                'VERIFIED
        Else
            missingBefore = missingBefore + 1
        End If
    Next i

    WL "present in model : " & (nBad - missingBefore)
    WL "already absent   : " & missingBefore
    WL ""

    If DRY_RUN = 1 Then
        WL "[DRY RUN] nothing deleted."
        WL "Set DRY_RUN = 0 and BACKUP_DONE = 1 to apply."
        Exit Sub
    End If

    rc = -1

    On Error Resume Next
    rc = femap.feDelete(FT_ELEM, s.ID)                    'UNVERIFIED
    If Err <> 0 Then
        rc = -999
    End If
    Err = 0
    On Error GoTo 0

    If rc = -999 Then
        WL "feDelete raised an error - the signature is wrong for"
        WL "this build. NOTHING was deleted."
        WL "Use delete_ids.txt with Delete > Model > Element instead."
        Exit Sub
    End If

    WL "feDelete returned: " & rc

    ' --- verify by re-fetching ---
    stillThere = 0
    For i = 0 To nBad - 1
        Set oEl = femap.feElem
        If oEl.Get(badID(i)) = -1 Then
            stillThere = stillThere + 1
        End If
    Next i

    WL ""
    WL "verification:"
    WL "  still present after delete : " & stillThere

    If stillThere = 0 Then
        WL "  ALL TARGET ELEMENTS REMOVED."
        WL ""
        WL "Next: re-solve. Expect a DIFFERENT error list -"
        WL "constraint and singularity errors live past SEMG and"
        WL "have never been reached."
    Else
        WL "  DELETE DID NOT WORK. Use delete_ids.txt manually."
    End If

End Sub

'==============================================================
' HELPERS
'==============================================================

Function FindLatestF06(dirPath As String) As String

    Dim fname As String
    Dim best As String
    Dim bestTime As Date
    Dim thisTime As Date
    Dim full As String
    Dim first As Integer

    FindLatestF06 = ""
    best = ""
    first = 1

    On Error GoTo BailFind

    fname = Dir(dirPath & "*.f06")

    Do While fname <> ""
        full = dirPath & fname
        thisTime = FileDateTime(full)
        If first = 1 Then
            best = full
            bestTime = thisTime
            first = 0
        ElseIf thisTime > bestTime Then
            best = full
            bestTime = thisTime
        End If
        fname = Dir()
    Loop

    FindLatestF06 = best
    Exit Function

BailFind:
    FindLatestF06 = best

End Function

Function GrabID(s As String) As Long

    Dim p As Integer
    Dim rest As String
    Dim i As Integer
    Dim c As String

    p = InStr(s, "ID =")

    If p = 0 Then
        GrabID = -1
        Exit Function
    End If

    rest = LTrim(Mid(s, p + 4))
    i = 1

    Do While i <= Len(rest)
        c = Mid(rest, i, 1)
        If c < "0" Or c > "9" Then
            Exit Do
        End If
        i = i + 1
    Loop

    If i = 1 Then
        GrabID = -1
    Else
        GrabID = CLng(Left(rest, i - 1))
    End If

End Function

Function EnsureDir(p As String) As Integer

    Dim i As Integer
    Dim c As String
    Dim partial As String

    EnsureDir = 0
    partial = ""

    For i = 1 To Len(p)
        c = Mid(p, i, 1)
        partial = partial & c
        If c = "\" Then
            If Len(partial) > 3 Then
                On Error Resume Next
                MkDir partial
                On Error GoTo 0
            End If
        End If
    Next i

    On Error Resume Next
    If Dir(p, 16) <> "" Then
        EnsureDir = 1
    End If
    On Error GoTo 0

End Function

Sub WL(s As String)

    Print #logF, s

End Sub
