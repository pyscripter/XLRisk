Attribute VB_Name = "Correlations"
Option Explicit
Option Base 1

Public Function RiskIsValidCorrmat(CorrmatRng As Range, _
    Optional Tolerance As Double = 0.0000000001, Optional ShowMessages As Boolean = False) As Boolean
Attribute RiskIsValidCorrmat.VB_Description = "Checks whether a correlation matrix is valid.  Returns a Boolean value"
Attribute RiskIsValidCorrmat.VB_ProcData.VB_Invoke_Func = " \n20"
' Checks wether a range is a valid correlation matrix
    Dim I As Integer
    Dim J As Integer
    Dim NVars As Integer
    Dim Corrmat() As Double
    
    RiskIsValidCorrmat = True
    
    ' Must be square
    If CorrmatRng.Rows.Count <> CorrmatRng.Columns.Count Then
        If ShowMessages Then MsgBox "The correlation matrix in range " & CorrmatRng.Address(, , , True) _
            & " is not square, so the program cannot continue.", vbExclamation, "Invalid Correlation Matrix"
        RiskIsValidCorrmat = False
        Exit Function
    End If
    
    NVars = CorrmatRng.Rows.Count
    ReDim Corrmat(1 To NVars, 1 To NVars)
    
    ' Fill the array with values from the range.
    For I = 1 To NVars
        For J = 1 To NVars
            Corrmat(I, J) = CorrmatRng(I, J).Value
        Next
    Next
    
    ' Fill in the correlation array if it's missing the lower or upper part.
    For I = 1 To NVars
        For J = 1 To NVars
            If I <> J Then
                If Corrmat(I, J) = 0 Then
                    Corrmat(I, J) = Corrmat(J, I)
                ElseIf Corrmat(J, I) = 0 Then
                    Corrmat(J, I) = Corrmat(I, J)
                End If
            End If
        Next
    Next
    
    ' Obvious properties
    For I = 1 To NVars
        For J = 1 To I
            If J = I Then
                If Abs(Corrmat(I, J) - 1) > Tolerance Then
                    If ShowMessages Then MsgBox "The correlation matrix in range " & CorrmatRng.Address(, , , True) _
                        & " doesn't have all 1's on the diagonal.", _
                        vbExclamation, "Invalid Correlation Matrix"
                    RiskIsValidCorrmat = False
                    Exit Function
                End If
            ElseIf Not (Corrmat(I, J) > -1 And Corrmat(I, J) < 1) Then
                If ShowMessages Then MsgBox "The correlation matrix in range " & CorrmatRng.Address(, , , True) _
                    & " doesn't have all off-diagonal values between -1 and 1.", _
                    vbExclamation, "Invalid Correlation Matrix"
                RiskIsValidCorrmat = False
                Exit Function
            ElseIf Abs(Corrmat(I, J) - Corrmat(J, I)) > Tolerance Then
                If ShowMessages Then MsgBox "The correlation matrix in range " & CorrmatRng.Address(, , , True) _
                    & " is not symmetric about the diagonal.", _
                    vbExclamation, "Invalid Correlation Matrix"
                RiskIsValidCorrmat = False
                Exit Function
            End If
        Next
    Next
    
    ' Eigen Values positive
    If WorksheetFunction.Min(EigenValues(Corrmat, Tolerance)) < 0 Then
        If ShowMessages Then MsgBox "The correlation matrix in range " & CorrmatRng.Address(, , , True) _
            & " is not a 'valid' correlation matrix. Please use the RiskCorrectCorrmat to fix this issue.", _
            vbExclamation, "Invalid Correlation Matrix"
        RiskIsValidCorrmat = False
    End If
End Function

Function RiskCorrectCorrmat(CorrmatRng As Range, Optional Tolerance As Double = 0.0000000001) As Variant
Attribute RiskCorrectCorrmat.VB_Description = "Fixes an invalid correlation matrix and returns the corrected matrix"
Attribute RiskCorrectCorrmat.VB_ProcData.VB_Invoke_Func = " \n20"
'   Fixes an invalid correlation matrix  using a simple scaling methodology described in
'   https://kb.palisade.com/index.php?pg=kb.page&id=75
'   An alternative is described in https://www.avrahamadler.com/2013/08/19/correcting-a-pseudo-correlation-matrix-to-be-positive-semidefinite/
'   and https://www.risklatte.xyz/Articles/QuantitativeFinance/QF152.php
    Dim I As Integer
    Dim J As Integer
    Dim NVars As Integer
    Dim Corrmat() As Double
    Dim MinEigenValue As Double
    
    ' Must be square
    If CorrmatRng.Rows.Count <> CorrmatRng.Columns.Count Then
        RiskCorrectCorrmat = CVErr(xlErrValue)
        Exit Function
    End If
    
    NVars = CorrmatRng.Rows.Count
    ReDim Corrmat(1 To NVars, 1 To NVars)
    
    ' Fill the array with values from the range.
    For I = 1 To NVars
        For J = 1 To NVars
            Corrmat(I, J) = CorrmatRng(I, J).Value
        Next
    Next
    
    ' Fill in the correlation array if it's missing the lower or upper part.
    For I = 1 To NVars
        For J = 1 To NVars
            If I <> J Then
                If Corrmat(I, J) = 0 Then
                    Corrmat(I, J) = Corrmat(J, I)
                ElseIf Corrmat(J, I) = 0 Then
                    Corrmat(J, I) = Corrmat(I, J)
                End If
            End If
        Next
    Next
    
    ' Obvious properties
    For I = 1 To NVars
        For J = 1 To I
            If J = I Then
                If Abs(Corrmat(I, J) - 1) > Tolerance Then
                    RiskCorrectCorrmat = CVErr(xlErrValue)
                    Exit Function
                End If
            ElseIf Not (Corrmat(I, J) > -1 And Corrmat(I, J) < 1) Then
                RiskCorrectCorrmat = CVErr(xlErrValue)
                Exit Function
            ElseIf Abs(Corrmat(I, J) - Corrmat(J, I)) > Tolerance Then
                RiskCorrectCorrmat = CVErr(xlErrValue)
                Exit Function
            End If
        Next
    Next
    
    ' Eigen Values positive
    MinEigenValue = WorksheetFunction.Min(EigenValues(Corrmat, Tolerance))
    If MinEigenValue < 0 Then
        'Step 2 C' = C – EoI
        MinEigenValue = MinEigenValue - Tolerance
        For I = 1 To NVars
            Corrmat(I, I) = Corrmat(I, I) - MinEigenValue
        Next I
        'Step 3 C'' = (1/(1-Eo)) C'
        For I = 1 To NVars
            For J = 1 To NVars
                Corrmat(I, J) = Corrmat(I, J) / (1 - MinEigenValue)
            Next J
        Next I
    End If
    RiskCorrectCorrmat = Corrmat
End Function

Public Function RiskSCorrel(Array1 As Variant, Array2 As Variant) As Variant
Attribute RiskSCorrel.VB_Description = "Returns the Spearman's rank correlation of two arrays"
Attribute RiskSCorrel.VB_ProcData.VB_Invoke_Func = " \n20"
    Dim Ranks1() As Double
    Dim Ranks2() As Double
    Dim N As Long
    Dim I As Long
    With WorksheetFunction
        N = .Count(Array1)
        If N <> .Count(Array2) Then
            RiskSCorrel = CVErr(xlErrValue)
            Exit Function
        End If
        ReDim Ranks1(1 To N)
        ReDim Ranks2(1 To N)
        
        For I = 1 To N
            Ranks1(I) = .Rank_Avg(Array1(I), Array1, 1)
            Ranks2(I) = .Rank_Avg(Array2(I), Array2, 1)
        Next I
        RiskSCorrel = .Correl(Ranks1, Ranks2)
    End With
End Function
