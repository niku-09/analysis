'==============================================================
' scr_fix_final_v2.bas
'--------------------------------------------------------------
' Final fixer for the SCR upper-structure shell model.
' FEMAP 2022.2.2 + Simcenter Nastran, SOL 101, units MMGS.
'
' READS the .f06 to learn WHICH elements Nastran rejected, then
' MEASURES each one against the live model. Nothing is hardcoded.
' Point it at a new .f06 after each solve and it re-evaluates.
'
'--------------------------------------------------------------
' WHAT IT DOES
'
'   Step A  Parse .f06           -> list of rejected element IDs
'   Step B  Measure + classify   -> verdict, property, centroid
'   Step C  WHERE TO LOOK report -> per-property rollup + bbox
'   Step D  Apply the safe fixes -> only when RUN_MODE > 0
'   Step E  Summary + groups
'
'   Steps A/B/C are READ-ONLY and run in every mode.
'
'--------------------------------------------------------------
' WHAT IT FIXES, AND WHY THAT IS PROVABLY CORRECT
'
'   COLLAPSED  quad with exactly ONE zero-length edge.
'              Two corners occupy the same point, so the element
'              IS a triangle. The three surviving nodes are
'              unambiguous. Converted to CTRIA3.
'
'   BOWTIE     quad whose connectivity crosses itself, or which
'              has a corner at/over 180 deg. Reordering is
'              deterministic, and every candidate order is
'              geometrically validated BEFORE being written.
'
'--------------------------------------------------------------
' WHAT IT REFUSES TO TOUCH
'
'   SLIVER            all 4 nodes distinct, aspect ratio absurd.
'                     No rule recovers the intended shape.
'   COLLAPSED_TO_LINE two or more zero-length edges. That is a
'                     line, not a triangle.
'   BAD_TRIA          degenerate triangle. Cannot be reordered
'                     out of trouble.
'   loaded elements   reordering flips the element normal, and
'                     the Wx/Wy wind scripts pick windward faces
'                     BY NORMAL DIRECTION. Flipping a loaded
'                     shell silently reverses its pressure - no
'                     error, just a wrong answer.
'
'--------------------------------------------------------------
' API STATUS ON THIS BUILD
'
'   READS - VERIFIED. A prior script ran cleanly through all of
'           these before failing elsewhere: feElem, Get, vnode
'           (whole-array read), topology, propID, feNode, xyz,
'           feSet.
'   WRITES- UNVERIFIED. Therefore EVERY write is followed by a
'           READ-BACK CHECK: the element is re-fetched and the
'           change confirmed. Three consecutive failures aborts
'           the run. The realistic bad outcome is "nothing
'           happened, cleanly" - not silent mass corruption.
'
'   Note: oElem.vnode reads fine as a whole array even though
'   oProp.pval does not. Array-property behaviour differs per
'   object, so both indexed and whole-array writes are attempted.
'
'--------------------------------------------------------------
' RUN_MODE
'   0 = report only. No model change. THIS IS THE DEFAULT.
'   1 = fix ONE element of each kind, then stop. Inspect them.
'   2 = fix all eligible.
'
' >>> SAVE A COPY OF THE MODEL BEFORE RUN_MODE 1 OR 2 <<<
'     The script refuses to write unless BACKUP_DONE = 1.
'==============================================================

'--------------------------------------------------------------
' CONFIG
'--------------------------------------------------------------
Const RUN_MODE        As Integer = 0
Const BACKUP_DONE     As Integer = 0

' AUTO_FIND_LATEST
'   1 = scan SCAN_DIR for *.f06 and use the NEWEST one.
'       FEMAP increments run numbers (-000, -001, -002...), so a
'       fixed path silently analyses a stale run and reports
'       problems as fixed when they are not.
'   0 = use LOG_PATH exactly as written.
Const AUTO_FIND_LATEST As Integer = 1

Const SCAN_DIR        As String = "C:\Users\nikhil.mohan\Desktop\Analysis testing\"

' Used only when AUTO_FIND_LATEST = 0.
Const LOG_PATH        As String = "C:\Users\nikhil.mohan\Desktop\Analysis testing\testinggg-000.f06"

Const OUT_DIR         As String = "C:\Users\nikhil.mohan\Desktop\Analysis testing\audit\"

' Refuse to parse a .f06 that does not contain END OF JOB. A file
' the solver is still writing would otherwise be read as if it
' were complete, producing a short and misleading error list.
Const REQUIRE_JOB_END As Integer = 1

' Edge at or below this length counts as collapsed (mm).
' If COLLAPSED comes back near zero while FEMAP reported 3307
' blocked merge pairs, raise this to 0.5 and re-run Step B.
Const ZERO_EDGE_MM    As Double = 0.05

' All nodes distinct but aspect ratio above this = sliver, manual.
Const AR_HOPELESS     As Double = 500#

' Interior angle at or above this counts as folded (deg).
' Nastran's own fatal threshold is 180.
Const FOLD_ANGLE_DEG  As Double = 179#

Const PROTECT_LOADED  As Integer = 1
Const MAX_WRITE_FAILS As Integer = 3

Const MAX_IDS         As Long = 20000
Const MAX_PROPS       As Integer = 500

' FEMAP constants
Const TOPO_TRIA3      As Integer = 2   'UNVERIFIED - read-back catches it
Const TOPO_QUAD4      As Integer = 4   'VERIFIED
Const FT_ELEM         As Integer = 8   'VERIFIED

' Verdict codes
Const V_UNREADABLE    As Integer = 0
Const V_COLLAPSED     As Integer = 1
Const V_BOWTIE        As Integer = 2
Const V_SLIVER        As Integer = 3
Const V_LINE          As Integer = 4
Const V_CLEAN         As Integer = 5
Const V_BADTRIA       As Integer = 6

'--------------------------------------------------------------
' GLOBALS
'--------------------------------------------------------------
Dim femap As Object
Dim logF As Integer

Dim badID() As Long
Dim nBad As Long

Dim vVerdict() As Integer
Dim vProp() As Long
Dim vCX() As Double
Dim vCY() As Double
Dim vCZ() As Double
Dim vMinEdge() As Double
Dim vAR() As Double
Dim vMaxAng() As Double
Dim vKeep0() As Long
Dim vKeep1() As Long
Dim vKeep2() As Long

Dim writeFails As Integer
Dim abortRun As Integer
Dim activeLog As String

'==============================================================
' MAIN
'==============================================================
Sub Main

    Set femap = GetObject(, "femap.model")                'VERIFIED

    writeFails = 0
    abortRun = 0

    If EnsureDir(OUT_DIR) = 0 Then
        femap.feAppMessage 2, "Cannot create output folder: " & OUT_DIR
        End
    End If

    ' --- resolve which .f06 to read ------------------------------
    If AUTO_FIND_LATEST = 1 Then
        activeLog = FindLatestF06(SCAN_DIR)
        If activeLog = "" Then
            femap.feAppMessage 2, "No .f06 found in " & SCAN_DIR
            End
        End If
    Else
        activeLog = LOG_PATH
    End If

    If FileThere(activeLog) = 0 Then
        femap.feAppMessage 2, "Solver log not found: " & activeLog
        End
    End If

    If REQUIRE_JOB_END = 1 Then
        If JobFinished(activeLog) = 0 Then
            femap.feAppMessage 2, "This .f06 has no END OF JOB marker - the solver may still be running: " & activeLog
            End
        End If
    End If

    If RUN_MODE > 0 Then
        If BACKUP_DONE <> 1 Then
            femap.feAppMessage 2, "REFUSING TO WRITE. Save a backup copy of the model, then set BACKUP_DONE = 1."
            End
        End If
    End If

    logF = FreeFile
    Open OUT_DIR & "scr_fix_final.log" For Output As #logF

    WL "=============================================="
    WL " scr_fix_final"
    WL " RUN_MODE = " & RUN_MODE
    WL " source   = " & activeLog
    If AUTO_FIND_LATEST = 1 Then
        WL " (auto-selected as newest .f06 in " & SCAN_DIR & ")"
    End If
    WL "=============================================="
    WL ""

    Call StepA_Collect
    Call StepB_Classify
    Call StepC_WhereToLook

    If RUN_MODE > 0 Then
        Call StepD_ApplyFixes
    End If

    Call StepE_Summary

    Close #logF

    femap.feAppMessage 0, "scr_fix_final complete. Read " & OUT_DIR & "where_to_look.csv first."

End Sub

'==============================================================
' STEP A - collect rejected element IDs from the .f06
'  UFM 4298 (thickness) is deliberately NOT collected - that is
'  a property card fix, already done manually.
'==============================================================
Sub StepA_Collect

    Dim f As Integer
    Dim ln As String
    Dim prev As String
    Dim eid As Long
    Dim k As Long
    Dim dup As Integer
    Dim n4296 As Long
    Dim n4297 As Long
    Dim n4558 As Long

    WL "--- Step A: reading solver log ---"

    ReDim badID(0 To MAX_IDS)
    nBad = 0
    prev = ""
    n4296 = 0
    n4297 = 0
    n4558 = 0

    f = FreeFile
    Open activeLog For Input As #f

    Do While Not EOF(f)

        Line Input #f, ln
        eid = -1

        If InStr(ln, "ILLEGAL GEOMETRY FOR QUAD4/QUADR ELEMENT WITH ID") > 0 Then
            eid = GrabID(ln)
            n4296 = n4296 + 1
        ElseIf InStr(ln, "HAS AN INTERIOR ANGLE POSSIBLY GREATER") > 0 Then
            eid = GrabID(prev)
            n4297 = n4297 + 1
        ElseIf InStr(ln, "SPECIFIED FOR ELEMENT WITH ID") > 0 Then
            If InStr(prev, "INAPPROPRIATE GEOMETRY OR INCORRECT MATERIAL DATA") > 0 Then
                eid = GrabID(ln)
                n4558 = n4558 + 1
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

    WL "  UFM 4296 illegal winding messages : " & n4296
    WL "  UFM 4297 folded >=180 messages    : " & n4297
    WL "  UFM 4558 bad TRIA3 messages       : " & n4558
    WL "  unique elements to examine        : " & nBad
    WL ""

    If nBad = 0 Then
        WL "  Nothing to do. The log contains no geometry fatals."
        WL ""
    End If

End Sub

'==============================================================
' STEP B - measure and classify against the LIVE model
'  Read-only. Runs in every RUN_MODE.
'==============================================================
Sub StepB_Classify

    Dim i As Long
    Dim fC As Integer
    Dim nColl As Long
    Dim nBow As Long
    Dim nSlv As Long
    Dim nLine As Long
    Dim nCln As Long
    Dim nTri As Long
    Dim nUnr As Long

    WL "--- Step B: measuring against live model ---"

    ReDim vVerdict(0 To nBad)
    ReDim vProp(0 To nBad)
    ReDim vCX(0 To nBad)
    ReDim vCY(0 To nBad)
    ReDim vCZ(0 To nBad)
    ReDim vMinEdge(0 To nBad)
    ReDim vAR(0 To nBad)
    ReDim vMaxAng(0 To nBad)
    ReDim vKeep0(0 To nBad)
    ReDim vKeep1(0 To nBad)
    ReDim vKeep2(0 To nBad)

    nColl = 0
    nBow = 0
    nSlv = 0
    nLine = 0
    nCln = 0
    nTri = 0
    nUnr = 0

    fC = FreeFile
    Open OUT_DIR & "classification.csv" For Output As #fC
    Print #fC, "ElementID,Verdict,PropID,CentroidX,CentroidY,CentroidZ,MinEdge_mm,AspectRatio,MaxAngle_deg,Action"

    For i = 0 To nBad - 1

        Call MeasureOne(i)

        If vVerdict(i) = V_COLLAPSED Then
            nColl = nColl + 1
        ElseIf vVerdict(i) = V_BOWTIE Then
            nBow = nBow + 1
        ElseIf vVerdict(i) = V_SLIVER Then
            nSlv = nSlv + 1
        ElseIf vVerdict(i) = V_LINE Then
            nLine = nLine + 1
        ElseIf vVerdict(i) = V_CLEAN Then
            nCln = nCln + 1
        ElseIf vVerdict(i) = V_BADTRIA Then
            nTri = nTri + 1
        Else
            nUnr = nUnr + 1
        End If

        Print #fC, badID(i) & "," & VName(vVerdict(i)) & "," & vProp(i) & "," & Num(vCX(i)) & "," & Num(vCY(i)) & "," & Num(vCZ(i)) & "," & Num(vMinEdge(i)) & "," & Num(vAR(i)) & "," & Num(vMaxAng(i)) & "," & VAction(vVerdict(i))

    Next i

    Close #fC

    WL "  COLLAPSED  (-> CTRIA3, automatic)  : " & nColl
    WL "  BOWTIE     (-> reorder, automatic) : " & nBow
    WL "  SLIVER     (manual remesh)         : " & nSlv
    WL "  LINE       (manual delete)         : " & nLine
    WL "  BAD_TRIA   (manual remesh)         : " & nTri
    WL "  CLEAN      (no longer failing)     : " & nCln
    WL "  UNREADABLE (inspect by hand)       : " & nUnr
    WL ""
    WL "  automatic : " & (nColl + nBow)
    WL "  manual    : " & (nSlv + nLine + nTri + nUnr)
    WL ""

    If nColl = 0 Then
        If nSlv > 0 Then
            WL "  NOTE: zero COLLAPSED found. If FEMAP reported blocked"
            WL "  merge pairs, ZERO_EDGE_MM is too tight for this model."
            WL "  Raise it to 0.5 and re-run."
            WL ""
        End If
    End If

End Sub

'==============================================================
' MEASURE ONE ELEMENT
'  Fills the parallel arrays at index i. Read-only.
'==============================================================
Sub MeasureOne(i As Long)

    Dim eid As Long
    Dim oEl As Object
    Dim v As Variant
    Dim nd(0 To 3) As Long
    Dim x(0 To 3) As Double
    Dim y(0 To 3) As Double
    Dim z(0 To 3) As Double
    Dim nCorner As Integer
    Dim k As Integer
    Dim j As Integer
    Dim ok As Integer
    Dim d As Double
    Dim minEdge As Double
    Dim maxEdge As Double
    Dim nZero As Integer
    Dim dropIdx As Integer
    Dim maxAng As Double
    Dim reflexCount As Integer
    Dim w As Integer
    Dim sx As Double
    Dim sy As Double
    Dim sz As Double

    eid = badID(i)

    vVerdict(i) = V_UNREADABLE
    vProp(i) = 0
    vCX(i) = 0
    vCY(i) = 0
    vCZ(i) = 0
    vMinEdge(i) = 0
    vAR(i) = 0
    vMaxAng(i) = 0

    On Error GoTo BailMeasure

    Set oEl = femap.feElem                                'VERIFIED
    If oEl.Get(eid) <> -1 Then                            'VERIFIED
        Exit Sub
    End If

    vProp(i) = oEl.propID                                 'VERIFIED
    v = oEl.vnode                                         'VERIFIED

    If oEl.topology = TOPO_QUAD4 Then                     'VERIFIED
        nCorner = 4
    Else
        nCorner = 3
    End If

    ok = 1
    sx = 0
    sy = 0
    sz = 0

    For k = 0 To nCorner - 1
        nd(k) = CLng(v(k))
        If GetNodeXYZ(nd(k), x(k), y(k), z(k)) = 0 Then
            ok = 0
        Else
            sx = sx + x(k)
            sy = sy + y(k)
            sz = sz + z(k)
        End If
    Next k

    If ok = 0 Then
        Exit Sub
    End If

    vCX(i) = sx / nCorner
    vCY(i) = sy / nCorner
    vCZ(i) = sz / nCorner

    ' --- edges ---
    minEdge = 1E+30
    maxEdge = 0
    nZero = 0
    dropIdx = -1

    For k = 0 To nCorner - 1
        j = (k + 1) Mod nCorner
        d = Dist(x(k), y(k), z(k), x(j), y(j), z(j))
        If d < minEdge Then
            minEdge = d
        End If
        If d > maxEdge Then
            maxEdge = d
        End If
        If d <= ZERO_EDGE_MM Then
            nZero = nZero + 1
            dropIdx = j
        End If
    Next k

    vMinEdge(i) = minEdge

    If minEdge > 0.000001 Then
        vAR(i) = maxEdge / minEdge
    Else
        vAR(i) = 9.99E+9
    End If

    Call CornerAngles(x, y, z, nCorner, maxAng, reflexCount)
    vMaxAng(i) = maxAng

    ' --- verdict ---

    ' quad with exactly one collapsed edge IS a triangle
    If nCorner = 4 Then
        If nZero = 1 Then
            w = 0
            For k = 0 To 3
                If k <> dropIdx Then
                    If w = 0 Then
                        vKeep0(i) = nd(k)
                    ElseIf w = 1 Then
                        vKeep1(i) = nd(k)
                    ElseIf w = 2 Then
                        vKeep2(i) = nd(k)
                    End If
                    w = w + 1
                End If
            Next k
            vVerdict(i) = V_COLLAPSED
            Exit Sub
        End If
    End If

    ' two or more collapsed edges = a line
    If nZero >= 2 Then
        vVerdict(i) = V_LINE
        Exit Sub
    End If

    ' triangle with any collapsed edge = a line
    If nCorner = 3 Then
        If nZero >= 1 Then
            vVerdict(i) = V_LINE
            Exit Sub
        End If
    End If

    ' nodes all distinct, geometry hopeless
    If vAR(i) >= AR_HOPELESS Then
        vVerdict(i) = V_SLIVER
        Exit Sub
    End If

    If nCorner = 3 Then
        If maxAng >= FOLD_ANGLE_DEG Then
            vVerdict(i) = V_BADTRIA
        Else
            vVerdict(i) = V_CLEAN
        End If
        Exit Sub
    End If

    ' quad, repairable by reordering
    If reflexCount >= 1 Then
        vVerdict(i) = V_BOWTIE
        Exit Sub
    End If

    If maxAng >= FOLD_ANGLE_DEG Then
        vVerdict(i) = V_BOWTIE
        Exit Sub
    End If

    vVerdict(i) = V_CLEAN
    Exit Sub

BailMeasure:
    vVerdict(i) = V_UNREADABLE

End Sub

'==============================================================
' STEP C - WHERE TO LOOK
'  The most useful output in this script. 887 scattered element
'  IDs are unusable; a list of affected PROPERTIES, each with a
'  bounding box, tells you which physical parts to open.
'
'  A property normally maps to a plate or a part. Heavy
'  concentration in a few properties means a few sliver regions,
'  not hundreds of independent defects.
'==============================================================
Sub StepC_WhereToLook

    Dim pID(0 To MAX_PROPS) As Long
    Dim pTot(0 To MAX_PROPS) As Long
    Dim pAuto(0 To MAX_PROPS) As Long
    Dim pMan(0 To MAX_PROPS) As Long
    Dim pMinX(0 To MAX_PROPS) As Double
    Dim pMaxX(0 To MAX_PROPS) As Double
    Dim pMinY(0 To MAX_PROPS) As Double
    Dim pMaxY(0 To MAX_PROPS) As Double
    Dim pMinZ(0 To MAX_PROPS) As Double
    Dim pMaxZ(0 To MAX_PROPS) As Double

    Dim nP As Integer
    Dim i As Long
    Dim k As Integer
    Dim slot As Integer
    Dim fW As Integer
    Dim isAuto As Integer

    WL "--- Step C: where to look ---"

    nP = 0

    For i = 0 To nBad - 1

        If vVerdict(i) <> V_CLEAN Then

            slot = -1
            For k = 0 To nP - 1
                If pID(k) = vProp(i) Then
                    slot = k
                End If
            Next k

            If slot < 0 Then
                If nP < MAX_PROPS Then
                    slot = nP
                    pID(slot) = vProp(i)
                    pTot(slot) = 0
                    pAuto(slot) = 0
                    pMan(slot) = 0
                    pMinX(slot) = 1E+30
                    pMaxX(slot) = -1E+30
                    pMinY(slot) = 1E+30
                    pMaxY(slot) = -1E+30
                    pMinZ(slot) = 1E+30
                    pMaxZ(slot) = -1E+30
                    nP = nP + 1
                End If
            End If

            If slot >= 0 Then

                pTot(slot) = pTot(slot) + 1

                isAuto = 0
                If vVerdict(i) = V_COLLAPSED Then
                    isAuto = 1
                End If
                If vVerdict(i) = V_BOWTIE Then
                    isAuto = 1
                End If

                If isAuto = 1 Then
                    pAuto(slot) = pAuto(slot) + 1
                Else
                    pMan(slot) = pMan(slot) + 1
                End If

                If vCX(i) < pMinX(slot) Then
                    pMinX(slot) = vCX(i)
                End If
                If vCX(i) > pMaxX(slot) Then
                    pMaxX(slot) = vCX(i)
                End If
                If vCY(i) < pMinY(slot) Then
                    pMinY(slot) = vCY(i)
                End If
                If vCY(i) > pMaxY(slot) Then
                    pMaxY(slot) = vCY(i)
                End If
                If vCZ(i) < pMinZ(slot) Then
                    pMinZ(slot) = vCZ(i)
                End If
                If vCZ(i) > pMaxZ(slot) Then
                    pMaxZ(slot) = vCZ(i)
                End If

            End If

        End If

    Next i

    fW = FreeFile
    Open OUT_DIR & "where_to_look.csv" For Output As #fW
    Print #fW, "PropID,BadElements,AutoFixable,NeedsManual,MinX,MaxX,MinY,MaxY,MinZ,MaxZ,SpanX,SpanY,SpanZ"

    For k = 0 To nP - 1
        Print #fW, pID(k) & "," & pTot(k) & "," & pAuto(k) & "," & pMan(k) & "," & Num(pMinX(k)) & "," & Num(pMaxX(k)) & "," & Num(pMinY(k)) & "," & Num(pMaxY(k)) & "," & Num(pMinZ(k)) & "," & Num(pMaxZ(k)) & "," & Num(pMaxX(k) - pMinX(k)) & "," & Num(pMaxY(k) - pMinY(k)) & "," & Num(pMaxZ(k) - pMinZ(k))
    Next k

    Close #fW

    WL "  affected properties: " & nP
    WL ""
    WL "  properties with the most damage:"

    Call ReportTopProps(pID, pTot, pAuto, pMan, nP)

    WL ""
    WL "  Full table in where_to_look.csv."
    WL "  A SMALL span on a property = one localized sliver region."
    WL "  A LARGE span = the whole part meshed badly."
    WL ""

End Sub

' Prints the worst offenders without needing a real sort.
Sub ReportTopProps(pID() As Long, pTot() As Long, pAuto() As Long, pMan() As Long, nP As Integer)

    Dim shown As Integer
    Dim pass As Integer
    Dim k As Integer
    Dim best As Integer
    Dim bestVal As Long
    Dim used(0 To MAX_PROPS) As Integer

    For k = 0 To nP - 1
        used(k) = 0
    Next k

    shown = 0
    pass = 0

    Do While pass < 10

        best = -1
        bestVal = -1

        For k = 0 To nP - 1
            If used(k) = 0 Then
                If pTot(k) > bestVal Then
                    bestVal = pTot(k)
                    best = k
                End If
            End If
        Next k

        If best < 0 Then
            Exit Do
        End If

        used(best) = 1
        WL "    Prop " & pID(best) & "  -  " & pTot(best) & " bad  (" & pAuto(best) & " auto, " & pMan(best) & " manual)"
        shown = shown + 1
        pass = pass + 1

    Loop

    If nP > shown Then
        WL "    ... and " & (nP - shown) & " more properties, see CSV"
    End If

End Sub

'==============================================================
' STEP D - apply the safe fixes
'==============================================================
Sub StepD_ApplyFixes

    Dim i As Long
    Dim eid As Long
    Dim keep(0 To 2) As Long
    Dim nTriaDone As Long
    Dim nReorderDone As Long
    Dim nSkipLoaded As Long
    Dim nFailed As Long
    Dim didTria As Integer
    Dim didReorder As Integer
    Dim doIt As Integer

    WL "--- Step D: applying fixes ---"

    nTriaDone = 0
    nReorderDone = 0
    nSkipLoaded = 0
    nFailed = 0
    didTria = 0
    didReorder = 0

    For i = 0 To nBad - 1

        If abortRun = 1 Then
            Exit For
        End If

        eid = badID(i)

        If vVerdict(i) = V_COLLAPSED Then

            doIt = 0
            If RUN_MODE = 2 Then
                doIt = 1
            ElseIf didTria = 0 Then
                doIt = 1
            End If

            If doIt = 1 Then
                keep(0) = vKeep0(i)
                keep(1) = vKeep1(i)
                keep(2) = vKeep2(i)
                If ConvertToTria(eid, keep) = 1 Then
                    nTriaDone = nTriaDone + 1
                    didTria = 1
                    If RUN_MODE = 1 Then
                        WL "  [RUN_MODE 1] converted elem " & eid & " to TRIA3. Inspect it."
                    End If
                Else
                    nFailed = nFailed + 1
                End If
            End If

        ElseIf vVerdict(i) = V_BOWTIE Then

            doIt = 0
            If RUN_MODE = 2 Then
                doIt = 1
            ElseIf didReorder = 0 Then
                doIt = 1
            End If

            If doIt = 1 Then
                If PROTECT_LOADED = 1 And ElementHasLoad(eid) = 1 Then
                    nSkipLoaded = nSkipLoaded + 1
                Else
                    If ReorderQuad(eid) = 1 Then
                        nReorderDone = nReorderDone + 1
                        didReorder = 1
                        If RUN_MODE = 1 Then
                            WL "  [RUN_MODE 1] reordered elem " & eid & ". Inspect it."
                        End If
                    Else
                        nFailed = nFailed + 1
                    End If
                End If
            End If

        End If

    Next i

    WL "  converted to CTRIA3     : " & nTriaDone
    WL "  reordered               : " & nReorderDone
    WL "  skipped (carries load)  : " & nSkipLoaded
    WL "  failed                  : " & nFailed
    WL ""

    If nSkipLoaded > 0 Then
        WL "  Loaded elements were skipped on purpose. Reordering"
        WL "  flips the element normal, and the wind scripts select"
        WL "  windward faces by normal direction. Fix these by hand"
        WL "  and re-check the pressure direction afterwards."
        WL ""
    End If

End Sub

'==============================================================
' FIX A - collapsed quad -> CTRIA3, with read-back verification
'==============================================================
Function ConvertToTria(eid As Long, keep() As Long) As Integer

    Dim oEl As Object
    Dim v As Variant
    Dim wrote As Integer
    Dim okBack As Integer

    ConvertToTria = 0
    wrote = 0

    If keep(0) <= 0 Or keep(1) <= 0 Or keep(2) <= 0 Then
        Exit Function
    End If

    On Error GoTo BailTria

    Set oEl = femap.feElem
    If oEl.Get(eid) <> -1 Then
        Exit Function
    End If

    ' attempt 1 - indexed write
    On Error Resume Next
    oEl.topology = TOPO_TRIA3                             'UNVERIFIED
    oEl.vnode(0) = keep(0)                                'UNVERIFIED
    oEl.vnode(1) = keep(1)
    oEl.vnode(2) = keep(2)
    oEl.vnode(3) = 0
    If Err = 0 Then
        wrote = 1
    End If
    Err = 0
    On Error GoTo BailTria

    ' attempt 2 - whole-array write
    If wrote = 0 Then
        On Error Resume Next
        v = oEl.vnode
        v(0) = keep(0)
        v(1) = keep(1)
        v(2) = keep(2)
        v(3) = 0
        oEl.topology = TOPO_TRIA3
        oEl.vnode = v                                     'UNVERIFIED
        If Err = 0 Then
            wrote = 1
        End If
        Err = 0
        On Error GoTo BailTria
    End If

    If wrote = 0 Then
        Call NoteWriteFail(eid, "topology/vnode write rejected")
        Exit Function
    End If

    If oEl.Put(eid) <> -1 Then                            'VERIFIED pattern
        Call NoteWriteFail(eid, "Put failed")
        Exit Function
    End If

    ' --- READ BACK. the safety net. ---
    okBack = 0
    Set oEl = femap.feElem
    If oEl.Get(eid) = -1 Then
        If oEl.topology = TOPO_TRIA3 Then
            v = oEl.vnode
            If CLng(v(0)) = keep(0) Then
                If CLng(v(1)) = keep(1) Then
                    If CLng(v(2)) = keep(2) Then
                        okBack = 1
                    End If
                End If
            End If
        End If
    End If

    If okBack = 1 Then
        ConvertToTria = 1
        writeFails = 0
    Else
        Call NoteWriteFail(eid, "read-back mismatch - element did NOT change")
    End If

    Exit Function

BailTria:
    Call NoteWriteFail(eid, "runtime error during convert")

End Function

'==============================================================
' FIX B - reorder a bowtie / folded quad, with verification
'==============================================================
Function ReorderQuad(eid As Long) As Integer

    Dim oEl As Object
    Dim v As Variant
    Dim nd(0 To 3) As Long
    Dim tryOrder(0 To 2, 0 To 3) As Integer
    Dim t As Integer
    Dim k As Integer
    Dim ok As Integer
    Dim x(0 To 3) As Double
    Dim y(0 To 3) As Double
    Dim z(0 To 3) As Double
    Dim maxAng As Double
    Dim reflexCount As Integer
    Dim chosen As Integer
    Dim wrote As Integer
    Dim okBack As Integer

    ReorderQuad = 0
    chosen = -1

    On Error GoTo BailReorder

    Set oEl = femap.feElem
    If oEl.Get(eid) <> -1 Then
        Exit Function
    End If

    If oEl.topology <> TOPO_QUAD4 Then
        Exit Function
    End If

    v = oEl.vnode
    For k = 0 To 3
        nd(k) = CLng(v(k))
    Next k

    tryOrder(0, 0) = 0
    tryOrder(0, 1) = 1
    tryOrder(0, 2) = 3
    tryOrder(0, 3) = 2

    tryOrder(1, 0) = 0
    tryOrder(1, 1) = 2
    tryOrder(1, 2) = 1
    tryOrder(1, 3) = 3

    tryOrder(2, 0) = 0
    tryOrder(2, 1) = 3
    tryOrder(2, 2) = 2
    tryOrder(2, 3) = 1

    ' validate candidates BEFORE writing anything
    For t = 0 To 2
        If chosen < 0 Then
            ok = 1
            For k = 0 To 3
                If GetNodeXYZ(nd(tryOrder(t, k)), x(k), y(k), z(k)) = 0 Then
                    ok = 0
                End If
            Next k
            If ok = 1 Then
                Call CornerAngles(x, y, z, 4, maxAng, reflexCount)
                If reflexCount = 0 Then
                    If maxAng < FOLD_ANGLE_DEG Then
                        chosen = t
                    End If
                End If
            End If
        End If
    Next t

    If chosen < 0 Then
        Exit Function
    End If

    wrote = 0

    On Error Resume Next
    For k = 0 To 3
        oEl.vnode(k) = nd(tryOrder(chosen, k))            'UNVERIFIED
    Next k
    If Err = 0 Then
        wrote = 1
    End If
    Err = 0
    On Error GoTo BailReorder

    If wrote = 0 Then
        On Error Resume Next
        v = oEl.vnode
        For k = 0 To 3
            v(k) = nd(tryOrder(chosen, k))
        Next k
        oEl.vnode = v
        If Err = 0 Then
            wrote = 1
        End If
        Err = 0
        On Error GoTo BailReorder
    End If

    If wrote = 0 Then
        Call NoteWriteFail(eid, "vnode write rejected")
        Exit Function
    End If

    If oEl.Put(eid) <> -1 Then
        Call NoteWriteFail(eid, "Put failed")
        Exit Function
    End If

    okBack = 0
    Set oEl = femap.feElem
    If oEl.Get(eid) = -1 Then
        v = oEl.vnode
        If CLng(v(1)) = nd(tryOrder(chosen, 1)) Then
            If CLng(v(2)) = nd(tryOrder(chosen, 2)) Then
                okBack = 1
            End If
        End If
    End If

    If okBack = 1 Then
        ReorderQuad = 1
        writeFails = 0
    Else
        Call NoteWriteFail(eid, "read-back mismatch - element did NOT change")
    End If

    Exit Function

BailReorder:
    Call NoteWriteFail(eid, "runtime error during reorder")

End Function

'==============================================================
' WRITE FAILURE TRACKING
'  Three consecutive failures means the write API is wrong, not
'  that three elements were unusual. Stop rather than grind
'  through hundreds of elements achieving nothing.
'==============================================================
Sub NoteWriteFail(eid As Long, reason As String)

    writeFails = writeFails + 1
    WL "  WRITE FAILED on elem " & eid & " : " & reason

    If writeFails >= MAX_WRITE_FAILS Then
        abortRun = 1
        WL ""
        WL "  *** ABORTED: " & MAX_WRITE_FAILS & " consecutive write failures."
        WL "  *** The write API signature is wrong for this build."
        WL "  *** Nothing was modified by the failed attempts."
        WL "  ***"
        WL "  *** Tools > Programming > Record, change one element's"
        WL "  *** nodes in the GUI, stop recording, and compare the"
        WL "  *** recorded call against ConvertToTria."
        WL "  ***"
        WL "  *** classification.csv and where_to_look.csv are still"
        WL "  *** complete and valid - use them for manual work."
        WL ""
    End If

End Sub

'==============================================================
' STEP E - summary and groups
'==============================================================
Sub StepE_Summary

    Dim setManual As Object
    Dim setAuto As Object
    Dim i As Long
    Dim nMan As Long

    WL "--- Step E: summary ---"

    Set setManual = femap.feSet                           'VERIFIED
    Set setAuto = femap.feSet
    setManual.Clear
    setAuto.Clear

    nMan = 0

    For i = 0 To nBad - 1
        If vVerdict(i) = V_SLIVER Or vVerdict(i) = V_LINE Or vVerdict(i) = V_BADTRIA Or vVerdict(i) = V_UNREADABLE Then
            setManual.Add badID(i)                        'VERIFIED
            nMan = nMan + 1
        ElseIf vVerdict(i) = V_COLLAPSED Or vVerdict(i) = V_BOWTIE Then
            setAuto.Add badID(i)
        End If
    Next i

    Call MakeGroup(setManual, "FIX_MANUAL_REVIEW")
    Call MakeGroup(setAuto, "FIX_AUTOMATIC")

    WL ""

    If abortRun = 1 Then
        WL "  Run aborted on write failures. Reports are still valid."
    ElseIf RUN_MODE = 0 Then
        WL "  REPORT ONLY. Nothing was changed."
        WL ""
        WL "  Next:"
        WL "    1. Open where_to_look.csv. Which properties are hit?"
        WL "    2. Decide: is this secondary structure (patch and move"
        WL "       on) or primary connections (fix the CAD properly)?"
        WL "    3. Save a backup copy of the model."
        WL "    4. BACKUP_DONE = 1, RUN_MODE = 1. Inspect the two"
        WL "       elements it changes."
        WL "    5. If they look right, RUN_MODE = 2."
    Else
        WL "  Fixes applied."
        WL ""
        WL "  Next:"
        WL "    1. Group FIX_MANUAL_REVIEW holds " & nMan & " elements."
        WL "       Isolate it, then View > Autoscale to jump to them."
        WL "    2. Delete and remesh those locally."
        WL "    3. Re-run the solver."
        WL "    4. Re-run this script on the NEW .f06."
        WL ""
        WL "  Expect two or three rounds. Fixing one element can"
        WL "  expose a neighbour that was previously masked."
    End If

    WL ""
    WL "  Files in " & OUT_DIR
    WL "    where_to_look.csv    <- read this first"
    WL "    classification.csv   <- every element, verdict, location"
    WL "    scr_fix_final.log    <- this file"
    WL ""
    WL "  Analysed: " & activeLog

End Sub

'==============================================================
' GEOMETRY HELPERS - pure math, version independent
'==============================================================

Sub CornerAngles(x() As Double, y() As Double, z() As Double, n As Integer, ByRef maxAng As Double, ByRef reflexCount As Integer)

    Dim nx As Double
    Dim ny As Double
    Dim nz As Double
    Dim nl As Double
    Dim i As Integer
    Dim j As Integer
    Dim p As Integer
    Dim q As Integer
    Dim ax As Double
    Dim ay As Double
    Dim az As Double
    Dim bx As Double
    Dim by As Double
    Dim bz As Double
    Dim la As Double
    Dim lb As Double
    Dim dp As Double
    Dim ang As Double
    Dim cx As Double
    Dim cy As Double
    Dim cz As Double
    Dim skip As Integer

    maxAng = 0
    reflexCount = 0
    nx = 0
    ny = 0
    nz = 0

    ' Newell normal - robust on warped shells
    For i = 0 To n - 1
        j = (i + 1) Mod n
        nx = nx + (y(i) - y(j)) * (z(i) + z(j))
        ny = ny + (z(i) - z(j)) * (x(i) + x(j))
        nz = nz + (x(i) - x(j)) * (y(i) + y(j))
    Next i

    nl = Sqr(nx * nx + ny * ny + nz * nz)

    If nl < 0.000000001 Then
        maxAng = 180
        reflexCount = n
        Exit Sub
    End If

    nx = nx / nl
    ny = ny / nl
    nz = nz / nl

    For i = 0 To n - 1

        skip = 0
        p = (i + n - 1) Mod n
        q = (i + 1) Mod n

        ax = x(p) - x(i)
        ay = y(p) - y(i)
        az = z(p) - z(i)
        bx = x(q) - x(i)
        by = y(q) - y(i)
        bz = z(q) - z(i)

        la = Sqr(ax * ax + ay * ay + az * az)
        lb = Sqr(bx * bx + by * by + bz * bz)

        If la < 0.000000001 Or lb < 0.000000001 Then
            maxAng = 180
            reflexCount = reflexCount + 1
            skip = 1
        End If

        If skip = 0 Then

            dp = (ax * bx + ay * by + az * bz) / (la * lb)
            If dp > 1 Then
                dp = 1
            End If
            If dp < -1 Then
                dp = -1
            End If

            ang = Acos(dp) * 180 / 3.14159265358979

            cx = ay * bz - az * by
            cy = az * bx - ax * bz
            cz = ax * by - ay * bx

            If (cx * nx + cy * ny + cz * nz) < 0 Then
                ang = 360 - ang
                reflexCount = reflexCount + 1
            End If

            If ang > maxAng Then
                maxAng = ang
            End If

        End If

    Next i

End Sub

Function Dist(x1 As Double, y1 As Double, z1 As Double, x2 As Double, y2 As Double, z2 As Double) As Double

    Dim dx As Double
    Dim dy As Double
    Dim dz As Double

    dx = x2 - x1
    dy = y2 - y1
    dz = z2 - z1

    Dist = Sqr(dx * dx + dy * dy + dz * dz)

End Function

Function Acos(v As Double) As Double

    If v >= 1 Then
        Acos = 0
    ElseIf v <= -1 Then
        Acos = 3.14159265358979
    Else
        Acos = Atn(-v / Sqr(-v * v + 1)) + 2 * Atn(1)
    End If

End Function

Function GetNodeXYZ(nid As Long, ByRef x As Double, ByRef y As Double, ByRef z As Double) As Integer

    Dim oNd As Object
    Dim c As Variant

    GetNodeXYZ = 0

    If nid <= 0 Then
        Exit Function
    End If

    On Error GoTo BailNode

    Set oNd = femap.feNode                                'VERIFIED
    If oNd.Get(nid) <> -1 Then                            'VERIFIED
        Exit Function
    End If

    c = oNd.xyz                                           'VERIFIED
    x = c(0)
    y = c(1)
    z = c(2)

    GetNodeXYZ = 1
    Exit Function

BailNode:
    GetNodeXYZ = 0

End Function

'==============================================================
' LOAD CHECK
'  Fail-safe by design: any doubt returns 1 (carries load), so
'  the caller skips rather than risking a silent normal flip.
'==============================================================
Function ElementHasLoad(eid As Long) As Integer

    Dim oLM As Object
    Dim i As Long
    Dim found As Integer

    ElementHasLoad = 1

    On Error GoTo BailLoad

    found = 0

    Set oLM = femap.feLoadMesh                            'VERIFIED (no args)

    For i = 1 To oLM.CountLoads                           'UNVERIFIED
        If oLM.Get(i) = -1 Then
            If oLM.meshID = eid Then
                found = 1
            End If
        End If
    Next i

    ElementHasLoad = found
    Exit Function

BailLoad:
    ElementHasLoad = 1

End Function

'==============================================================
' GROUPS
'==============================================================
Sub MakeGroup(s As Object, nameStr As String)

    Dim gid As Long

    On Error GoTo BailGroup

    gid = femap.feGroupCreateFromSet(0, nameStr, FT_ELEM, s.ID)   'UNVERIFIED
    WL "  group created: " & nameStr
    Exit Sub

BailGroup:
    WL "  (group " & nameStr & " not created - CSV lists still valid)"

End Sub

'==============================================================
' UTILITY
'==============================================================

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

Function VName(v As Integer) As String

    If v = V_COLLAPSED Then
        VName = "COLLAPSED"
    ElseIf v = V_BOWTIE Then
        VName = "BOWTIE"
    ElseIf v = V_SLIVER Then
        VName = "SLIVER"
    ElseIf v = V_LINE Then
        VName = "COLLAPSED_TO_LINE"
    ElseIf v = V_CLEAN Then
        VName = "CLEAN"
    ElseIf v = V_BADTRIA Then
        VName = "BAD_TRIA"
    Else
        VName = "UNREADABLE"
    End If

End Function

Function VAction(v As Integer) As String

    If v = V_COLLAPSED Then
        VAction = "auto_convert_TRIA3"
    ElseIf v = V_BOWTIE Then
        VAction = "auto_reorder"
    ElseIf v = V_SLIVER Then
        VAction = "MANUAL_remesh"
    ElseIf v = V_LINE Then
        VAction = "MANUAL_delete"
    ElseIf v = V_CLEAN Then
        VAction = "none"
    ElseIf v = V_BADTRIA Then
        VAction = "MANUAL_remesh"
    Else
        VAction = "MANUAL_inspect"
    End If

End Function

Function Num(v As Double) As String

    Dim t As Double

    If v > 1E+15 Then
        Num = "9.99E+09"
        Exit Function
    End If

    If v < -1E+15 Then
        Num = "-9.99E+09"
        Exit Function
    End If

    t = Int(v * 1000 + 0.5) / 1000
    Num = CStr(t)

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

Function FileThere(p As String) As Integer

    FileThere = 0

    On Error Resume Next
    If Dir(p) <> "" Then
        FileThere = 1
    End If
    On Error GoTo 0

End Function

'--------------------------------------------------------------
' Returns the full path of the most recently modified .f06 in
' dirPath, or "" if none found.
'--------------------------------------------------------------
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

'--------------------------------------------------------------
' Returns 1 if the .f06 contains the END OF JOB marker, meaning
' the solver has finished writing it.
'--------------------------------------------------------------
Function JobFinished(p As String) As Integer

    Dim f As Integer
    Dim ln As String

    JobFinished = 0

    On Error GoTo BailFin

    f = FreeFile
    Open p For Input As #f

    Do While Not EOF(f)
        Line Input #f, ln
        If InStr(ln, "END OF JOB") > 0 Then
            JobFinished = 1
        End If
    Loop

    Close #f
    Exit Function

BailFin:
    JobFinished = 0

End Function

Sub WL(s As String)

    Print #logF, s

End Sub
