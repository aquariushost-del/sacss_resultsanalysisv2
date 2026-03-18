Attribute VB_Name = "modValidityEngine"

Option Explicit
'===========================================================
' Module: modValidityEngine
'
' PURPOSE:
'   Pure logic engine for interpreting grade distributions.
'   - Takes A1?F9 counts (array) + total N
'   - Computes:
'       * Validity flag  (LOW N / SKEWED / MIXED / VALID / NO DATA)
'       * Pattern type   (Clipped Top, Fat Middle, etc.)
'       * Teacher-facing text lines for the validity panel:
'           line1 = "What you see: ..."
'           line2 = "What it means: ..."
'           line3 = "What you can do: ..."
'
' ASSUMPTIONS:
'   - gradeCounts() is 1-based, indices 1..9:
'       1 = A1,  2 = A2,  3 = B3,
'       4 = B4,  5 = C5,  6 = C6,
'       7 = D7,  8 = E8,  9 = F9
'   - totalN = sum of all gradeCounts()
'
' HOW TO USE (from another module):
'
'   Dim flag As String, pattern As String
'   Dim l1 As String, l2 As String, l3 As String
'
'   Call EvaluateDistribution(gradeCounts, totalN, flag, pattern, l1, l2, l3)
'
'   ' Then pass l1, l2, l3 into your shape panel beside the chart.
'
'===========================================================

'-------------------
' Tunable thresholds
'-------------------
Private Const LOW_N_THRESHOLD          As Long = 10       ' N <= 10 ? LOW N
Private Const SKEW_TWO_BAND_WINDOW_PCT As Double = 70#    ' any 2-adjacent bands >= 70%
Private Const SKEW_TOP3_PCT            As Double = 60#    ' top 3 bands >= 60%
Private Const SKEW_BOTTOM3_PCT         As Double = 60#    ' bottom 3 bands >= 60%
Private Const MIXED_MIN_HIGH_PCT       As Double = 20#    ' A1?A2 >= 20%
Private Const MIXED_MIN_LOW_PCT        As Double = 20#    ' D7?F9 >= 20%
Private Const MID_MIN_FOR_FAT_MID      As Double = 50#    ' C5?C6 (and neighbours) >= 50%
Private Const MID_MAX_FOR_THIN_MID     As Double = 20#    ' mid band <= 20% for Thin Middle
Private Const CLIP_EDGE_LIMIT_PCT      As Double = 10#    ' opposite tail <= 10% for clipped
Private Const TIGHT_CLUSTER_WINDOW_PCT As Double = 80#    ' any 3 consecutive bands >= 80%
Private Const WIDE_SPREAD_MIN_BANDS    As Long = 6        ' at least 6 bands with >= 5%
Private Const WIDE_SPREAD_MIN_BAND_PCT As Double = 5#     ' band counted if >= 5%
Private Const STEPPED_DELTA_TOL        As Double = 5#     ' tolerance for monotonic decline (percentage points)
Private Const SPIKY_MIN_SIGN_CHANGES   As Long = 3        ' sign changes in deltas to call "Spiky"

'===========================================================
' PUBLIC ENTRY POINT
'===========================================================
Public Sub EvaluateDistribution( _
    ByRef gradeCounts() As Long, _
    ByVal totalN As Long, _
    ByRef validityFlag As String, _
    ByRef patternType As String, _
    ByRef line1 As String, _
    ByRef line2 As String, _
    ByRef line3 As String)

    Dim gradePercents() As Double

    ' Default outputs
    validityFlag = "NO DATA"
    patternType = "No Data"
    line1 = ""
    line2 = ""
    line3 = ""

    ' Guard: no candidates
    If totalN <= 0 Then
        validityFlag = "NO DATA"
        patternType = "No Data"
        GetInterpretationText validityFlag, patternType, line1, line2, line3
        Exit Sub
    End If

    ' Build percentage array from counts
    gradePercents = BuildPercentArray(gradeCounts, totalN)

    ' Decide flag
    validityFlag = GetValidityFlag(gradeCounts, gradePercents, totalN)

    ' Decide pattern type (even for VALID / MIXED / SKEWED where helpful)
    patternType = GetPatternType(gradePercents, totalN, validityFlag)

    ' Get teacher-facing text
    GetInterpretationText validityFlag, patternType, line1, line2, line3
End Sub

'===========================================================
' SCHEME-AWARE ENTRY POINT (G3 / G2 / G1)
'===========================================================
Public Sub EvaluateDistributionForScheme( _
    ByRef gradeCounts() As Long, _
    ByVal totalN As Long, _
    ByVal schemeKey As String, _
    ByRef validityFlag As String, _
    ByRef patternType As String, _
    ByRef line1 As String, _
    ByRef line2 As String, _
    ByRef line3 As String)

    Dim normCounts(1 To 9) As Long
    Dim i As Long, n As Long
    Dim sk As String

    sk = UCase$(Trim$(schemeKey))
    n = UBound(gradeCounts) - LBound(gradeCounts) + 1

    Select Case sk
        Case "G3"
            For i = 1 To 9
                If LBound(gradeCounts) + i - 1 <= UBound(gradeCounts) Then
                    normCounts(i) = gradeCounts(LBound(gradeCounts) + i - 1)
                End If
            Next i

        Case "G2"
            ' Map 6 bands to 9-band shape while preserving order/edges.
            ' 1->1, 2->2, 3->4, 4->6, 5->8, 6->9
            If n >= 1 Then normCounts(1) = gradeCounts(LBound(gradeCounts) + 0)
            If n >= 2 Then normCounts(2) = gradeCounts(LBound(gradeCounts) + 1)
            If n >= 3 Then normCounts(4) = gradeCounts(LBound(gradeCounts) + 2)
            If n >= 4 Then normCounts(6) = gradeCounts(LBound(gradeCounts) + 3)
            If n >= 5 Then normCounts(8) = gradeCounts(LBound(gradeCounts) + 4)
            If n >= 6 Then normCounts(9) = gradeCounts(LBound(gradeCounts) + 5)

        Case "G1"
            ' Map 5 bands to 9-band shape while preserving order/edges.
            ' A->1, B->3, C->5, D->7, E->9
            If n >= 1 Then normCounts(1) = gradeCounts(LBound(gradeCounts) + 0)
            If n >= 2 Then normCounts(3) = gradeCounts(LBound(gradeCounts) + 1)
            If n >= 3 Then normCounts(5) = gradeCounts(LBound(gradeCounts) + 2)
            If n >= 4 Then normCounts(7) = gradeCounts(LBound(gradeCounts) + 3)
            If n >= 5 Then normCounts(9) = gradeCounts(LBound(gradeCounts) + 4)

        Case Else
            For i = 1 To 9
                If LBound(gradeCounts) + i - 1 <= UBound(gradeCounts) Then
                    normCounts(i) = gradeCounts(LBound(gradeCounts) + i - 1)
                End If
            Next i
    End Select

    EvaluateDistribution normCounts, totalN, validityFlag, patternType, line1, line2, line3
End Sub

'===========================================================
' PERCENT ARRAY BUILDER
'===========================================================
Private Function BuildPercentArray( _
    ByRef gradeCounts() As Long, _
    ByVal totalN As Long) As Double()

    Dim lo As Long, hi As Long, i As Long
    Dim arr() As Double

    lo = LBound(gradeCounts)
    hi = UBound(gradeCounts)

    ReDim arr(lo To hi)

    If totalN <= 0 Then
        For i = lo To hi
            arr(i) = 0#
        Next i
        BuildPercentArray = arr
        Exit Function
    End If

    For i = lo To hi
        arr(i) = (CDbl(gradeCounts(i)) / CDbl(totalN)) * 100#
    Next i

    BuildPercentArray = arr
End Function

'===========================================================
' VALIDITY FLAG DECISION
'===========================================================
Public Function GetValidityFlag( _
    ByRef gradeCounts() As Long, _
    ByRef gradePercents() As Double, _
    ByVal totalN As Long) As String

    Dim lo As Long, hi As Long, i As Long
    Dim window2Max As Double
    Dim tmp As Double
    Dim sumTop3 As Double, sumBottom3 As Double
    Dim sumA1A2 As Double, sumD7F9 As Double

    lo = LBound(gradePercents)
    hi = UBound(gradePercents)

    '--------------------------
    ' 1. LOW N check (highest)
    '--------------------------
    If totalN <= LOW_N_THRESHOLD Then
        GetValidityFlag = "LOW N"
        Exit Function
    End If

    '--------------------------
    ' 2. Compute key aggregates
    '--------------------------
    ' top 3 = A1?B3 (1..3)
    sumTop3 = SafeSum(gradePercents, lo, lo + 2)      ' 1,2,3
    ' bottom 3 = D7?F9 (7..9) assuming 9 bands
    sumBottom3 = SafeSum(gradePercents, hi - 2, hi)   ' 7,8,9

    ' A1?A2
    sumA1A2 = SafeSum(gradePercents, lo, lo + 1)
    ' D7?F9
    sumD7F9 = sumBottom3

    ' 2-band moving window maximum
    window2Max = 0#
    For i = lo To hi - 1
        tmp = gradePercents(i) + gradePercents(i + 1)
        If tmp > window2Max Then window2Max = tmp
    Next i

    '--------------------------
    ' 3. SKEWED checks
    '--------------------------
    If window2Max >= SKEW_TWO_BAND_WINDOW_PCT _
       Or sumTop3 >= SKEW_TOP3_PCT _
       Or sumBottom3 >= SKEW_BOTTOM3_PCT Then

        GetValidityFlag = "SKEWED"
        Exit Function
    End If

    '--------------------------
    ' 4. MIXED checks
    '--------------------------
    If sumA1A2 >= MIXED_MIN_HIGH_PCT _
       And sumD7F9 >= MIXED_MIN_LOW_PCT Then

        GetValidityFlag = "MIXED"
        Exit Function
    End If

    '--------------------------
    ' 5. Default ? VALID
    '--------------------------
    GetValidityFlag = "VALID"
End Function

'===========================================================
' PATTERN TYPE DECISION (insight labels)
'===========================================================
Public Function GetPatternType( _
    ByRef gradePercents() As Double, _
    ByVal totalN As Long, _
    ByVal validityFlag As String) As String

    Dim lo As Long, hi As Long
    Dim sumTop3 As Double, sumMid As Double, sumBottom3 As Double
    Dim i As Long
    Dim isTightCluster As Boolean
    Dim countBands As Long
    Dim delta As Double, lastDelta As Double
    Dim signChanges As Long

    lo = LBound(gradePercents)
    hi = UBound(gradePercents)

    If totalN <= 0 Then
        GetPatternType = "No Data"
        Exit Function
    End If

    ' Aggregates
    sumTop3 = SafeSum(gradePercents, lo, lo + 2)        ' A1?B3
    sumMid = SafeSum(gradePercents, lo + 2, hi - 2)     ' B3?C6 approx middle
    sumBottom3 = SafeSum(gradePercents, hi - 2, hi)     ' D7?F9

    '---------------------------
    ' 1. If LOW N ? label it so
    '---------------------------
    If UCase$(validityFlag) = "LOW N" Then
        GetPatternType = "Small Cohort"
        Exit Function
    End If

    '----------------------------------
    ' 2. Clipped Top / Clipped Bottom
    '----------------------------------
    If sumTop3 >= SKEW_TOP3_PCT And sumBottom3 <= CLIP_EDGE_LIMIT_PCT Then
        GetPatternType = "Clipped Top"
        Exit Function
    End If

    If sumBottom3 >= SKEW_BOTTOM3_PCT And sumTop3 <= CLIP_EDGE_LIMIT_PCT Then
        GetPatternType = "Clipped Bottom"
        Exit Function
    End If

    '--------------------------
    ' 3. Thin Middle / Fat Middle
    '--------------------------
    If sumMid <= MID_MAX_FOR_THIN_MID _
       And sumTop3 >= MIXED_MIN_HIGH_PCT _
       And sumBottom3 >= MIXED_MIN_LOW_PCT Then

        GetPatternType = "Thin Middle"
        Exit Function
    End If

    If sumMid >= MID_MIN_FOR_FAT_MID _
       And sumTop3 <= (100# - MID_MIN_FOR_FAT_MID) _
       And sumBottom3 <= (100# - MID_MIN_FOR_FAT_MID) Then

        GetPatternType = "Fat Middle"
        Exit Function
    End If

    '--------------------------
    ' 4. Tight Cluster
    '--------------------------
    isTightCluster = False
    For i = lo To hi - 2
        If gradePercents(i) + gradePercents(i + 1) + gradePercents(i + 2) >= TIGHT_CLUSTER_WINDOW_PCT Then
            isTightCluster = True
            Exit For
        End If
    Next i

    If isTightCluster Then
        GetPatternType = "Tight Cluster"
        Exit Function
    End If

    '--------------------------
    ' 5. Wide Spread
    '--------------------------
    countBands = 0
    For i = lo To hi
        If gradePercents(i) >= WIDE_SPREAD_MIN_BAND_PCT Then
            countBands = countBands + 1
        End If
    Next i

    If countBands >= WIDE_SPREAD_MIN_BANDS Then
        GetPatternType = "Wide Spread"
        Exit Function
    End If

    '--------------------------
    ' 6. Stepped vs Spiky
    '--------------------------
    ' We look at percentage differences between adjacent bands.
    ' Mostly one direction (downwards) ? "Stepped"
    ' Many up-down sign changes ? "Spiky"

    lastDelta = 0#
    signChanges = 0

    For i = lo To hi - 1
        delta = gradePercents(i + 1) - gradePercents(i)

        ' Only care if the change is larger than tolerance
        If Abs(delta) > STEPPED_DELTA_TOL Then
            If lastDelta <> 0# Then
                If Sgn(delta) <> Sgn(lastDelta) Then
                    signChanges = signChanges + 1
                End If
            End If
            lastDelta = delta
        End If
    Next i

    If signChanges >= SPIKY_MIN_SIGN_CHANGES Then
        GetPatternType = "Spiky"
    ElseIf signChanges = 0 And gradePercents(lo) >= gradePercents(hi) Then
        GetPatternType = "Stepped"
    Else
        '--------------------------
        ' 7. Default: Balanced
        '--------------------------
        GetPatternType = "Balanced"
    End If
End Function

'===========================================================
' SAFE SUM HELPER
'===========================================================
Private Function SafeSum( _
    ByRef arr() As Double, _
    ByVal fromIdx As Long, _
    ByVal toIdx As Long) As Double

    Dim lo As Long, hi As Long, i As Long
    Dim s As Double

    lo = LBound(arr)
    hi = UBound(arr)

    If fromIdx < lo Then fromIdx = lo
    If toIdx > hi Then toIdx = hi

    s = 0#
    For i = fromIdx To toIdx
        s = s + arr(i)
    Next i

    SafeSum = s
End Function

'===========================================================
' TEXT GENERATOR ? FRIENDLY EXPLANATIONS
'===========================================================
Public Sub GetInterpretationText( _
    ByVal validityFlag As String, _
    ByVal patternType As String, _
    ByRef line1 As String, _
    ByRef line2 As String, _
    ByRef line3 As String)

    Dim f As String
    Dim p As String

    f = UCase$(Trim$(validityFlag))
    p = Trim$(patternType)

    Select Case f

        '---------------------------
        ' NO DATA
        '---------------------------
        Case "NO DATA"
            line1 = "What you see: No grades are available for this subject."
            line2 = "What it means: There is not enough information to form a pattern."
            line3 = "What you can do: Check if results have been entered correctly before interpreting."

        '---------------------------
        ' LOW N ? SMALL COHORT
        '---------------------------
        Case "LOW N"
            line1 = "What you see: Only a small number of students are represented in this chart."
            line2 = "What it means: The pattern can change a lot when one or two students move bands, so it is not a stable group trend."
            line3 = "What you can do: Use this mainly as a rough reference and focus more on individual student profiles than cohort comparisons."

        '---------------------------
        ' SKEWED ? ONE-SIDED
        '---------------------------
        Case "SKEWED"
            If p = "Clipped Top" Then
                line1 = "What you see: Many grades are concentrated in the highest bands, with few in the middle or lower bands."
                line2 = "What it means: The distribution is heavily top-weighted, so it is hard to see differences among high-performing students."
                line3 = "What you can do: Consider including more challenging tasks to stretch stronger learners and better differentiate high performance."
            ElseIf p = "Clipped Bottom" Then
                line1 = "What you see: Many grades are concentrated in the lowest bands, with few in the middle or higher bands."
                line2 = "What it means: The assessment may have been too demanding relative to current readiness, or key foundations are not secure."
                line3 = "What you can do: Revisit core concepts and scaffolds, and review whether future tasks can better match students? current level."
            Else
                line1 = "What you see: Grades are heavily concentrated at one end of the scale, with few students in the middle bands."
                line2 = "What it means: The distribution is one-sided and may reflect a narrow spread of readiness or an imbalanced level of difficulty."
                line3 = "What you can do: Review the range of task difficulty and consider how to either stretch or support the dominant group of learners."
            End If

        '---------------------------
        ' MIXED ? WIDE VARIATION
        '---------------------------
        Case "MIXED"
            If p = "Thin Middle" Then
                line1 = "What you see: There are noticeable groups at both the high and low ends, with relatively few students in the middle bands."
                line2 = "What it means: Performance is polarised, suggesting distinct groups of learners with different levels of readiness."
                line3 = "What you can do: Plan differentiated support for emerging learners and provide targeted extension for stronger learners."
            ElseIf p = "Wide Spread" Then
                line1 = "What you see: Grades are spread across many bands, with students at both high and low ends."
                line2 = "What it means: The cohort is diverse in performance, and the assessment distinguishes clearly between different levels of preparation."
                line3 = "What you can do: Use this spread to inform tiered support and extension, and ensure classroom tasks cater to varied starting points."
            Else
                line1 = "What you see: There are substantial numbers of students at both higher and lower grade bands."
                line2 = "What it means: The distribution suggests more than one performance group rather than a single, tightly clustered profile."
                line3 = "What you can do: Plan for differentiated teaching, with structured support and suitable challenge for learners at different levels."
            End If

        '---------------------------
        ' VALID ? BALANCED / HEALTHY
        '---------------------------
        Case "VALID"
            Select Case p
                Case "Fat Middle"
                    line1 = "What you see: Many students are in the middle bands, with smaller groups at the high and low ends."
                    line2 = "What it means: The assessment difficulty is generally well matched to the cohort, capturing typical understanding."
                    line3 = "What you can do: Strengthen core skills for the middle group and provide clear pathways for learners to move into higher bands."
                Case "Thin Middle"
                    line1 = "What you see: More students are represented at the high and low ends than in the middle bands."
                    line2 = "What it means: The cohort shows a more polarised profile, even though overall results are still interpretable."
                    line3 = "What you can do: Pay attention to both groups and consider how to support consolidation in the middle bands."
                Case "Tight Cluster"
                    line1 = "What you see: Most grades sit within a narrow range of bands."
                    line2 = "What it means: The cohort is relatively homogeneous in this subject, with similar levels of performance."
                    line3 = "What you can do: Use this to design focused instruction for the group, while still identifying individuals who need extra support or challenge."
                Case "Wide Spread"
                    line1 = "What you see: Grades stretch across many bands, with no single cluster dominating."
                    line2 = "What it means: There is a wide range of readiness and performance levels within the cohort."
                    line3 = "What you can do: Plan for tiered tasks and flexible grouping so that learners at different levels can progress meaningfully."
                Case "Stepped"
                    line1 = "What you see: The bars form a gradual decline from the stronger bands to the weaker bands."
                    line2 = "What it means: The distribution follows a natural pattern, and the assessment differentiates consistently across the cohort."
                    line3 = "What you can do: Use this distribution to identify where most students are currently performing and to target shifts into higher bands."
                Case "Spiky"
                    line1 = "What you see: The bars rise and fall sharply across different grade bands, without a smooth pattern."
                    line2 = "What it means: Some parts of the assessment may have been much easier or harder than others, leading to an uneven profile."
                    line3 = "What you can do: Review item types and sections to see which parts were unexpectedly difficult or easy, and adjust future tasks accordingly."
                Case Else  ' Balanced / default
                    line1 = "What you see: Grades are reasonably spread across several bands, without extreme clustering at any one end."
                    line2 = "What it means: The pattern is stable and interpretable, and the assessment appears to differentiate learners appropriately."
                    line3 = "What you can do: Use these results confidently for trend analysis, comparison across groups, and planning next steps in teaching."
            End Select

        '---------------------------
        ' FALLBACK (unknown flag)
        '---------------------------
        Case Else
            line1 = "What you see: The distribution does not match any predefined pattern clearly."
            line2 = "What it means: The results may require a closer look before drawing conclusions."
            line3 = "What you can do: Combine this chart with qualitative evidence from lessons and assessments to understand the cohort more fully."
    End Select
End Sub



