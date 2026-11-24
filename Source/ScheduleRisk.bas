Attribute VB_Name = "ScheduleRisk"
Option Explicit
Option Base 1

Private mActivityCount As Long
Private mActivityIds() As Variant
Private mDurations() As Double
Private mStart() As Double
Private mFinish() As Double
Private mSlack() As Double
Private mCritical() As Boolean
Private mCriticalCounts() As Long
Private mRecordedIteration As Long
Private mSignature As String

Private Function ScheduleVectorCount(Value As Variant) As Long
    On Error GoTo Finish
    If TypeName(Value) = "Range" Then
        ScheduleVectorCount = Value.Cells.Count
    Else
        Dim D1 As Long, D2 As Long
        D1 = UBound(Value, 1)
        On Error Resume Next
        D2 = UBound(Value, 2)
        On Error GoTo 0
        If D2 > 0 Then
            ScheduleVectorCount = IIf(D1 >= D2, D1, D2)
        Else
            ScheduleVectorCount = D1
        End If
    End If
Finish:
End Function

Private Function ScheduleVectorValue(Value As Variant, Index As Long) As Variant
    On Error GoTo Finish
    If TypeName(Value) = "Range" Then
        ScheduleVectorValue = Value.Cells(Index).Value
    Else
        Dim D1 As Long, D2 As Long
        D1 = UBound(Value, 1)
        On Error Resume Next
        D2 = UBound(Value, 2)
        On Error GoTo 0
        If D2 > 0 Then
            If D1 >= D2 Then
                ScheduleVectorValue = Value(Index, 1)
            Else
                ScheduleVectorValue = Value(1, Index)
            End If
        Else
            ScheduleVectorValue = Value(Index)
        End If
    End If
Finish:
End Function

Private Function ScheduleSignature(ActivityIds As Variant, Durations As Variant, Predecessors As Variant, _
                                  Optional DependencyTypes As Variant, Optional Lags As Variant) As String
    Dim IdText As String
    On Error Resume Next
    If TypeName(ActivityIds) = "Range" Then
        IdText = ActivityIds.Address(True, True, xlA1, True)
    Else
        IdText = CStr(ScheduleVectorValue(ActivityIds, 1))
    End If
    If TypeName(Durations) = "Range" Then IdText = IdText & "|" & Durations.Address(True, True, xlA1, True)
    If TypeName(Predecessors) = "Range" Then IdText = IdText & "|" & Predecessors.Address(True, True, xlA1, True)
    ScheduleSignature = IdText
End Function

Private Function FindActivityId(Id As Variant) As Long
    Dim I As Long
    For I = 1 To mActivityCount
        If CStr(mActivityIds(I)) = CStr(Id) Then
            FindActivityId = I
            Exit Function
        End If
    Next I
    FindActivityId = 0
End Function

Private Sub SplitScheduleText(Text As String, Count As Long, ByRef Parts() As String)
    Dim CleanText As String
    CleanText = Replace(Trim$(CStr(Text)), ";", ",")
    CleanText = Replace(CleanText, vbTab, ",")
    If CleanText = vbNullString Then
        Count = 0
        ReDim Parts(0)
    Else
        Parts = Split(CleanText, ",")
        Count = UBound(Parts) + 1
    End If
End Sub

Private Function ParseInlineLagAndType(Segment As String, ByRef Lag As Double, ByRef DependencyType As String) As String
    Dim I As Long
    Dim J As Long
    Dim Work As String
    Dim Ch As String
    Dim SignedNumber As String
    Dim Candidate As String

    Work = Trim$(Segment)
    Lag = 0
    DependencyType = "FS"
    For I = 1 To Len(Work)
        Candidate = UCase$(Mid$(Work, I, 2))
        If Candidate = "FS" Or Candidate = "SS" Or Candidate = "FF" Or Candidate = "SF" Then
            DependencyType = Candidate
            Work = Trim$(Replace(Work, Candidate, "", I, 1, vbTextCompare))
            Exit For
        End If
    Next I

    SignedNumber = vbNullString
    For J = Len(Work) To 1 Step -1
        Ch = Mid$(Work, J, 1)
        If Ch = "+" Or Ch = "-" Or (Ch >= "0" And Ch <= "9") Or Ch = "." Then
            SignedNumber = Ch & SignedNumber
        Else
            Exit For
        End If
    Next J
    If SignedNumber <> vbNullString Then
        On Error Resume Next
        Lag = CDbl(SignedNumber)
        On Error GoTo 0
        Work = Trim$(Left$(Work, Len(Work) - Len(SignedNumber)))
    End If
    ParseInlineLagAndType = Trim$(Work)
End Function

Private Sub ParsePredecessors(ActivityIndex As Long, Predecessors As Variant, DependencyTypes As Variant, _
                             Lags As Variant, ByRef PredCount As Long, ByRef PredIds() As Long, _
                             ByRef PredTypes() As String, ByRef PredLags() As Double)
    Dim Text As String
    Dim Parts() As String
    Dim PartCount As Long
    Dim I As Long
    Dim J As Long
    Dim Found As Long
    Dim SegmentId As String
    Dim DepText As String
    Dim LagText As String
    Dim LagParts() As String
    Dim LagCount As Long
    Dim DepParts() As String
    Dim DepCount As Long

    Text = CStr(ScheduleVectorValue(Predecessors, ActivityIndex))
    If Trim$(Text) = vbNullString Then
        PredCount = 0
        ReDim PredIds(0)
        ReDim PredTypes(0)
        ReDim PredLags(0)
        Exit Sub
    End If

    SplitScheduleText Text, PartCount, Parts
    DepText = vbNullString
    LagText = vbNullString
    On Error Resume Next
    If IsMissing(DependencyTypes) = False Then DepText = CStr(ScheduleVectorValue(DependencyTypes, ActivityIndex))
    If IsMissing(Lags) = False Then LagText = CStr(ScheduleVectorValue(Lags, ActivityIndex))
    On Error GoTo 0
    If Trim$(DepText) <> vbNullString Then SplitScheduleText DepText, DepCount, DepParts Else DepCount = 0
    If Trim$(LagText) <> vbNullString Then SplitScheduleText LagText, LagCount, LagParts Else LagCount = 0

    ReDim PredIds(1 To PartCount)
    ReDim PredTypes(1 To PartCount)
    ReDim PredLags(1 To PartCount)

    PredCount = 0
    For I = 0 To PartCount - 1
        Found = 0
        SegmentId = Trim$(Parts(I))
        For J = 1 To mActivityCount
            If Left$(UCase$(SegmentId), Len(CStr(mActivityIds(J)))) = UCase$(CStr(mActivityIds(J))) Then
                If Len(SegmentId) = Len(CStr(mActivityIds(J))) Or _
                   InStr(" ,;+-SF", Mid$(SegmentId, Len(CStr(mActivityIds(J))) + 1, 1)) > 0 Then
                    Found = J
                    Exit For
                End If
            End If
        Next J

        If Found > 0 Then
            PredCount = PredCount + 1
            PredIds(PredCount) = Found
            PredTypes(PredCount) = "FS"
            PredLags(PredCount) = 0
            If DepCount >= PredCount Then
                PredTypes(PredCount) = UCase$(Trim$(DepParts(PredCount - 1)))
                Select Case PredTypes(PredCount)
                    Case "FS", "SS", "FF", "SF"
                    Case Else
                        PredTypes(PredCount) = "FS"
                End Select
            Else
                Dim InlineLag As Double
                Dim InlineType As String
                SegmentId = ParseInlineLagAndType(SegmentId, InlineLag, InlineType)
                PredTypes(PredCount) = InlineType
                PredLags(PredCount) = InlineLag
            End If
            If LagCount >= PredCount Then
                On Error Resume Next
                PredLags(PredCount) = CDbl(Trim$(LagParts(PredCount - 1)))
                On Error GoTo 0
            End If
        End If
    Next I
End Sub

Private Sub BuildSchedule(ActivityIds As Variant, Durations As Variant, Predecessors As Variant, _
                         Optional DependencyTypes As Variant, Optional Lags As Variant)
    Dim I As Long
    Dim J As Long
    Dim Iter As Long
    Dim K As Long
    Dim Changed As Boolean
    Dim ProjectFinish As Double
    Dim PredCount As Long
    Dim PredIds() As Long
    Dim PredTypes() As String
    Dim PredLags() As Double
    Dim Candidate As Double

    mActivityCount = ScheduleVectorCount(ActivityIds)
    ReDim mActivityIds(1 To mActivityCount)
    ReDim mDurations(1 To mActivityCount)
    ReDim mStart(1 To mActivityCount)
    ReDim mFinish(1 To mActivityCount)
    ReDim mSlack(1 To mActivityCount)
    ReDim mCritical(1 To mActivityCount)

    For I = 1 To mActivityCount
        mActivityIds(I) = ScheduleVectorValue(ActivityIds, I)
        On Error Resume Next
        mDurations(I) = CDbl(ScheduleVectorValue(Durations, I))
        If Err <> 0 Or mDurations(I) < 0 Then mDurations(I) = 0
        Err.Clear
        On Error GoTo 0
        mStart(I) = 0
    Next I

    For Iter = 1 To mActivityCount
        Changed = False
        For I = 1 To mActivityCount
            ParsePredecessors I, Predecessors, DependencyTypes, Lags, PredCount, PredIds, PredTypes, PredLags
            For J = 1 To PredCount
                K = PredIds(J)
                Select Case PredTypes(J)
                    Case "SS"
                        Candidate = mStart(K) + PredLags(J)
                    Case "FF"
                        Candidate = mFinish(K) + PredLags(J) - mDurations(I)
                    Case "SF"
                        Candidate = mStart(K) + PredLags(J) - mDurations(I)
                    Case Else
                        Candidate = mFinish(K) + PredLags(J)
                End Select
                If Candidate > mStart(I) Then
                    mStart(I) = Candidate
                    Changed = True
                End If
            Next J
            mFinish(I) = mStart(I) + mDurations(I)
        Next I
        If Not Changed Then Exit For
    Next Iter

    ProjectFinish = 0
    For I = 1 To mActivityCount
        If mFinish(I) > ProjectFinish Then ProjectFinish = mFinish(I)
    Next I

    Dim LatestStart() As Double
    Dim LatestFinish() As Double
    ReDim LatestStart(1 To mActivityCount)
    ReDim LatestFinish(1 To mActivityCount)
    For I = 1 To mActivityCount
        LatestFinish(I) = ProjectFinish
        LatestStart(I) = LatestFinish(I) - mDurations(I)
    Next I

    For Iter = 1 To mActivityCount
        Changed = False
        For I = mActivityCount To 1 Step -1
            For J = 1 To mActivityCount
                ParsePredecessors J, Predecessors, DependencyTypes, Lags, PredCount, PredIds, PredTypes, PredLags
                Dim P As Long
                For P = 1 To PredCount
                    If PredIds(P) = I Then
                        Select Case PredTypes(P)
                            Case "SS"
                                Candidate = LatestStart(J) - PredLags(P)
                                If Candidate < LatestStart(I) Then LatestStart(I) = Candidate: Changed = True
                                LatestFinish(I) = LatestStart(I) + mDurations(I)
                            Case "FF"
                                Candidate = LatestFinish(J) - PredLags(P)
                                If Candidate < LatestFinish(I) Then LatestFinish(I) = Candidate: Changed = True
                                LatestStart(I) = LatestFinish(I) - mDurations(I)
                            Case "SF"
                                Candidate = LatestFinish(J) - PredLags(P)
                                If Candidate < LatestStart(I) Then LatestStart(I) = Candidate: Changed = True
                                LatestFinish(I) = LatestStart(I) + mDurations(I)
                            Case Else
                                Candidate = LatestStart(J) - PredLags(P)
                                If Candidate < LatestFinish(I) Then LatestFinish(I) = Candidate: Changed = True
                                LatestStart(I) = LatestFinish(I) - mDurations(I)
                        End Select
                    End If
                Next P
            Next J
        Next I
        If Not Changed Then Exit For
    Next Iter

    For I = 1 To mActivityCount
        mSlack(I) = LatestStart(I) - mStart(I)
        mCritical(I) = Abs(mSlack(I)) < 0.0000001
    Next I

    Dim NewSignature As String
    NewSignature = ScheduleSignature(ActivityIds, Durations, Predecessors, DependencyTypes, Lags)
    If NewSignature <> mSignature Then
        mSignature = NewSignature
        ReDim mCriticalCounts(1 To mActivityCount)
        mRecordedIteration = 0
    End If
End Sub

Public Sub ScheduleRiskReset()
    mActivityCount = 0
    mRecordedIteration = 0
    mSignature = vbNullString
End Sub

Public Sub ScheduleRiskRecordIteration(Iteration As Long)
    Dim I As Long
    If mActivityCount = 0 Or mRecordedIteration = Iteration Then Exit Sub
    For I = 1 To mActivityCount
        If mCritical(I) Then mCriticalCounts(I) = mCriticalCounts(I) + 1
    Next I
    mRecordedIteration = Iteration
End Sub

Public Function ScheduleRiskGetActivityCount() As Long
    ScheduleRiskGetActivityCount = mActivityCount
End Function

Public Function ScheduleRiskGetActivityId(Index As Long) As Variant
    If Index >= 1 And Index <= mActivityCount Then ScheduleRiskGetActivityId = mActivityIds(Index)
End Function

Public Function ScheduleRiskGetCriticalPercent(Index As Long) As Variant
    If Index < 1 Or Index > mActivityCount Then
        ScheduleRiskGetCriticalPercent = 0
    Else
        ScheduleRiskGetCriticalPercent = mCriticalCounts(Index) / IIf(mRecordedIteration > 0, mRecordedIteration, 1)
    End If
End Function

Public Function ScheduleRiskGetCriticalCount(Index As Long) As Long
    If Index >= 1 And Index <= mActivityCount Then ScheduleRiskGetCriticalCount = mCriticalCounts(Index)
End Function

Private Function EnsureSchedule(ActivityIds As Variant, Durations As Variant, Predecessors As Variant, _
                               Optional DependencyTypes As Variant, Optional Lags As Variant) As Boolean
    On Error GoTo Fail
    BuildSchedule ActivityIds, Durations, Predecessors, DependencyTypes, Lags
    EnsureSchedule = mActivityCount > 0
    Exit Function
Fail:
    EnsureSchedule = False
End Function

Public Function RiskScheduleFinish(ActivityIds As Variant, Durations As Variant, Predecessors As Variant, _
                                  Optional DependencyTypes As Variant, Optional Lags As Variant) As Variant
    Dim I As Long
    Dim ProjectFinish As Double
    If Not EnsureSchedule(ActivityIds, Durations, Predecessors, DependencyTypes, Lags) Then
        RiskScheduleFinish = CVErr(xlErrValue): Exit Function
    End If
    For I = 1 To mActivityCount
        If mFinish(I) > ProjectFinish Then ProjectFinish = mFinish(I)
    Next I
    RiskScheduleFinish = ProjectFinish
End Function

Public Function RiskScheduleActivityFinish(ActivityId As Variant, ActivityIds As Variant, Durations As Variant, Predecessors As Variant, _
                                          Optional DependencyTypes As Variant, Optional Lags As Variant) As Variant
    Dim I As Long
    If Not EnsureSchedule(ActivityIds, Durations, Predecessors, DependencyTypes, Lags) Then
        RiskScheduleActivityFinish = CVErr(xlErrValue): Exit Function
    End If
    I = FindActivityId(ActivityId)
    If I = 0 Then RiskScheduleActivityFinish = CVErr(xlErrValue) Else RiskScheduleActivityFinish = mFinish(I)
End Function

Public Function RiskScheduleActivityStart(ActivityId As Variant, ActivityIds As Variant, Durations As Variant, Predecessors As Variant, _
                                         Optional DependencyTypes As Variant, Optional Lags As Variant) As Variant
    Dim I As Long
    If Not EnsureSchedule(ActivityIds, Durations, Predecessors, DependencyTypes, Lags) Then
        RiskScheduleActivityStart = CVErr(xlErrValue): Exit Function
    End If
    I = FindActivityId(ActivityId)
    If I = 0 Then RiskScheduleActivityStart = CVErr(xlErrValue) Else RiskScheduleActivityStart = mStart(I)
End Function

Public Function RiskScheduleActivitySlack(ActivityId As Variant, ActivityIds As Variant, Durations As Variant, Predecessors As Variant, _
                                         Optional DependencyTypes As Variant, Optional Lags As Variant) As Variant
    Dim I As Long
    If Not EnsureSchedule(ActivityIds, Durations, Predecessors, DependencyTypes, Lags) Then
        RiskScheduleActivitySlack = CVErr(xlErrValue): Exit Function
    End If
    I = FindActivityId(ActivityId)
    If I = 0 Then RiskScheduleActivitySlack = CVErr(xlErrValue) Else RiskScheduleActivitySlack = mSlack(I)
End Function

Public Function RiskScheduleActivityCritical(ActivityId As Variant, ActivityIds As Variant, Durations As Variant, Predecessors As Variant, _
                                            Optional DependencyTypes As Variant, Optional Lags As Variant) As Variant
    Dim I As Long
    If Not EnsureSchedule(ActivityIds, Durations, Predecessors, DependencyTypes, Lags) Then
        RiskScheduleActivityCritical = CVErr(xlErrValue): Exit Function
    End If
    I = FindActivityId(ActivityId)
    If I = 0 Then
        RiskScheduleActivityCritical = CVErr(xlErrValue)
    ElseIf mCritical(I) Then
        RiskScheduleActivityCritical = 1
    Else
        RiskScheduleActivityCritical = 0
    End If
End Function

Public Function RiskScheduleDeadlineProbability(FinishValues As Range, TargetDay As Double) As Double
    Dim I As Long
    Dim Count As Long
    Dim ValidCount As Long
    Dim Value As Variant
    For I = 1 To FinishValues.Cells.Count
        Value = FinishValues.Cells(I).Value
        If IsNumeric(Value) Then
            ValidCount = ValidCount + 1
            If CDbl(Value) <= TargetDay Then Count = Count + 1
        End If
    Next I
    If ValidCount > 0 Then RiskScheduleDeadlineProbability = Count / ValidCount
End Function

Public Function RiskScheduleDeadline(FinishValues As Range, Confidence As Double) As Variant
    If Confidence < 0 Or Confidence > 1 Then RiskScheduleDeadline = CVErr(xlErrValue): Exit Function
    Dim Values() As Double
    Dim Count As Long
    Dim I As Long
    Dim Rank As Long
    Dim Value As Variant
    Count = 0
    ReDim Values(1 To FinishValues.Cells.Count)
    For I = 1 To FinishValues.Cells.Count
        Value = FinishValues.Cells(I).Value
        If IsNumeric(Value) Then Count = Count + 1: Values(Count) = CDbl(Value)
    Next I
    If Count = 0 Then RiskScheduleDeadline = CVErr(xlErrValue): Exit Function
    Values = ScheduleSmallToLarge(Values, Count)
    Rank = Application.WorksheetFunction.RoundUp(Confidence * Count, 0)
    If Rank < 1 Then Rank = 1
    If Rank > Count Then Rank = Count
    RiskScheduleDeadline = Values(Rank)
End Function

Private Function ScheduleSmallToLarge(Values() As Double, Count As Long) As Double()
    Dim I As Long
    Dim J As Long
    Dim Temp As Double
    For I = 1 To Count - 1
        For J = I + 1 To Count
            If Values(J) < Values(I) Then
                Temp = Values(I): Values(I) = Values(J): Values(J) = Temp
            End If
        Next J
    Next I
    ScheduleSmallToLarge = Values
End Function
