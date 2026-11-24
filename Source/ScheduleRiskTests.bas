Attribute VB_Name = "ScheduleRiskTests"
Option Explicit
Option Base 1

Public Sub ScheduleRisk_RunTests()
    Dim Passed As Long
    Dim Failed As Long

    If ScheduleTestChain(Passed, Failed) Then
        MsgBox "Schedule-risk tests passed: " & Passed & vbCrLf & "Failed: " & Failed, vbInformation, "XLRisk Schedule Tests"
    Else
        MsgBox "Schedule-risk tests failed: " & Failed, vbCritical, "XLRisk Schedule Tests"
    End If
End Sub

Private Function ScheduleTestChain(ByRef Passed As Long, ByRef Failed As Long) As Boolean
    Dim Ids As Variant
    Dim Durs As Variant
    Dim Preds As Variant
    Dim Value As Variant

    Ids = Array("A", "B", "C")
    Durs = Array(3, 4, 2)
    Preds = Array("", "A", "B")

    Value = RiskScheduleFinish(Ids, Durs, Preds)
    ScheduleAssert Abs(CDbl(Value) - 9) < 0.000001, "FS chain finish should be 9", Passed, Failed

    Value = RiskScheduleActivityStart("B", Ids, Durs, Preds)
    ScheduleAssert Abs(CDbl(Value) - 3) < 0.000001, "B should start at day 3", Passed, Failed

    Value = RiskScheduleActivityCritical("A", Ids, Durs, Preds)
    ScheduleAssert CInt(Value) = 1, "A should be critical in a three-task chain", Passed, Failed

    Value = RiskScheduleActivityCritical("C", Ids, Durs, Preds)
    ScheduleAssert CInt(Value) = 1, "C should be critical in a three-task chain", Passed, Failed

    Dim BranchDurs As Variant
    Dim BranchPreds As Variant
    BranchDurs = Array(3, 5, 2)
    BranchPreds = Array("", "", "A,B")
    Value = RiskScheduleFinish(Ids, BranchDurs, BranchPreds)
    ScheduleAssert Abs(CDbl(Value) - 7) < 0.000001, "Branch merge finish should be 7", Passed, Failed

    Value = RiskScheduleActivitySlack("A", Ids, BranchDurs, BranchPreds)
    ScheduleAssert Abs(CDbl(Value) - 2) < 0.000001, "A should have 2 days of total float", Passed, Failed

    Dim LagDurs As Variant
    Dim LagPreds As Variant
    Dim LagTypes As Variant
    Dim LagValues As Variant
    LagDurs = Array(3, 4)
    LagPreds = Array("", "A")
    LagTypes = Array("", "SS")
    LagValues = Array(0, 1)
    Value = RiskScheduleFinish(Array("A", "B"), LagDurs, LagPreds, LagTypes, LagValues)
    ScheduleAssert Abs(CDbl(Value) - 7) < 0.000001, "SS +1 schedule finish should be 7", Passed, Failed

    ScheduleTestChain = Failed = 0
End Function

Private Sub ScheduleAssert(Condition As Boolean, Message As String, ByRef Passed As Long, ByRef Failed As Long)
    If Condition Then
        Passed = Passed + 1
    Else
        Failed = Failed + 1
        Debug.Print "FAILED: " & Message
    End If
End Sub
