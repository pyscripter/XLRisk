Attribute VB_Name = "Math"
Option Explicit
Option Compare Text
Option Base 1

'   The Iman-Conover method implementation is based on an excelent series of articles on
'   the topic at https://www.howardrudd.net/how-tos/vba-monte-carlo-risk-analysis-spreadsheet-with-correlation-part-1/
'   The following subroutines and functions:
'    1. ArrayRank(vArray)
'    2. Cholesky(A)
'    3. FSnvS(F, S)
'    4. ImanConover(X, C)
'    5. MatMult(A, B)
'    6. MatTransMult(A, B
'    7. NormalScores(N, M)
'    8. QS1dd(inLow, InHi, KeyArray, OtherArray)
'    9. QS1ds(inLow, InHi, KeyArray)
'   10. QS2dd(inLow, InHi, KeyArray, Column, OtherArray)
'   11. QS2ds(inLow, InHi, KeyArray, Column)
'   12. QuickSort(KeyArray, Optional Column, Optional OtherArray)
'   13. Shuffle(vArray, Optional column)
'
'   are subject to Copyright 2015 by Howard J Rudd and licensed under the Apache License, Version 2.0


Public Function IdentityMatrix(N As Integer) As Double()
'   Returns the (nxn) Identity Matrix
    Dim I As Integer
    Dim Imat() As Double
    ReDim Imat(1 To N, 1 To N)
    For I = 1 To N
        Imat(I, I) = 1
    Next I
    IdentityMatrix = Imat
End Function

Public Function NumberOfArrayDimensions(Arr As Variant) As Long
' This function is from http://www.cpearson.com/Excel/VBAArrays.htm

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' NumberOfArrayDimensions
' This function returns the number of dimensions of an array. An unallocated dynamic array
' has 0 dimensions.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim Ndx As Long
    Dim Res As Long
    On Error Resume Next
    ' Loop, increasing the dimension index Ndx, until an error occurs.
    ' An error will occur when Ndx exceeds the number of dimension
    ' in the array. Return Ndx - 1.
    Do
        Ndx = Ndx + 1
        Res = UBound(Arr, Ndx)
    Loop Until Err.Number <> 0
    
    NumberOfArrayDimensions = Ndx - 1
End Function

Public Sub Shuffle(vArray() As Double, Optional Column As Long)

' Usage: shuffle(vArray, column)
'
' shuffles vArray in place randomly. If vArray is 2-dimensional, then it shuffles
' the column specified by the optional variable "column"
'
' From http://www.howardrudd.net/pages/mc_maths_functions.html
' Copyright 2015 Howard J Rudd
' Licensed under the Apache License, Version 2.0

Dim J As Long
Dim k As Long
Dim startIndex As Long
Dim endIndex As Long
Dim vDims As Long

Dim temp As Double

' Check that dimensionality of "vArray" is either 1 or 2
vDims = NumberOfArrayDimensions(vArray)
If Not (vDims = 1 Or vDims = 2) Then
    Debug.Print "input argument not one or two dimensional"
    Debug.Print vDims
    Exit Sub
End If

If vDims = 1 Then
    startIndex = LBound(vArray)
    endIndex = UBound(vArray)
    J = endIndex
    Do While J > startIndex
        k = CLng(startIndex + Rnd() * (J - startIndex))
        temp = vArray(J)
        vArray(J) = vArray(k)
        vArray(k) = temp
        J = J - 1
    Loop
ElseIf vDims = 2 Then
    If IsMissing(Column) Then
        Err.Raise Number:=vbObjectError + 50, _
            Source:="Shuffle", _
            Description:="Argument 'column' not supplied"
        Exit Sub
    End If
' Check that the argument "column" points to one of the columns of "vArray"
    If Not (LBound(vArray, 2) <= Column And Column <= UBound(vArray, 2)) Then
        Err.Raise Number:=vbObjectError + 51, _
            Source:="Shuffle", _
            Description:="Argument 'column' does not point to one of the columns of 'vArray'"
        Exit Sub
    End If
    startIndex = LBound(vArray, 1)
    endIndex = UBound(vArray, 1)
    J = endIndex
    Do While J > startIndex
        k = CLng(startIndex + Rnd() * (J - startIndex))
        temp = vArray(J, Column)
        vArray(J, Column) = vArray(k, Column)
        vArray(k, Column) = temp
        J = J - 1
    Loop
End If
End Sub

Private Function MatrixUTSumSq(Matrix As Variant) As Double
'   Returns the Sum of Squares of the Upper Triangle of a symmetric Matrix
    Dim Sum As Double
    Dim I As Integer
    Dim J As Integer
    Dim N As Integer
    N = Sqr(Application.Count(Matrix))
    Sum = 0
    For I = 1 To N
        For J = I + 1 To N
            Sum = Sum + (Matrix(I, J) ^ 2)
        Next J
    Next I
    MatrixUTSumSq = Sum
End Function

Private Function JacobiRvec(N As Integer, Athis As Variant) As Variant
'   Returns vector containing MR, MC and Jrad
'   These are the row and column vectors and the angle of rotation for the P matrix
    Dim MaxVal As Double
    Dim Jrad As Double
    Dim I As Integer
    Dim J As Integer
    Dim MR As Integer
    Dim MC As Integer
    Dim Awork() As Double
    
    ReDim Awork(1 To N, 1 To N)
    
    MaxVal = -1
    MR = -1
    MC = -1
    For I = 1 To N
        For J = I + 1 To N
            Awork(I, J) = Abs(Athis(I, J))
            If Awork(I, J) > MaxVal Then
                MaxVal = Awork(I, J)
                MR = I
                MC = J
            End If
        Next J
    Next I
    If Athis(MR, MR) = Athis(MC, MC) Then
        Jrad = 0.25 * Application.Pi() * Sgn(Athis(MR, MC))
    Else
        Jrad = 0.5 * Atn(2 * Athis(MR, MC) / (Athis(MR, MR) - Athis(MC, MC)))
    End If
    JacobiRvec = Array(MR, MC, Jrad)
End Function

Private Function JacobiPmat(N As Integer, Rthis As Variant) As Double()
'   Returns the rotation Pthis matrix
'   Uses MatrixIdentity fn
    Dim Pthis() As Double
    
    Pthis = IdentityMatrix(N)
    Pthis(Rthis(1), Rthis(1)) = Cos(Rthis(3))
    Pthis(Rthis(2), Rthis(1)) = Sin(Rthis(3))
    Pthis(Rthis(1), Rthis(2)) = -Sin(Rthis(3))
    Pthis(Rthis(2), Rthis(2)) = Cos(Rthis(3))
    JacobiPmat = Pthis
End Function

Private Function JacobiAmat(N As Integer, Athis As Variant) As Variant
'   Returns Anext matrix, updated using the P rotation matrix
'   Uses Jacobirvec fn
'   Uses JacobiPmat fn
    Dim Rthis As Variant
    Dim Pthis As Variant
    Dim Anext As Variant
    
    Rthis = JacobiRvec(N, Athis)
    Pthis = JacobiPmat(N, Rthis)
    With WorksheetFunction
        Anext = .MMult(.Transpose(Pthis), .MMult(Athis, Pthis))
    End With
    JacobiAmat = Anext
End Function

Public Function EigenValues(Matrix As Variant, Optional Tolerance As Double = 0.0000000001) As Double()
'   Uses the Jacobi method to get the eigenvalues for a symmetric matrix
'   Amat is rotated (using the P matrix) until its off-diagonal elements are minimal
'   Uses MatrixUTSumSq fn
'   Uses JacobiAmat fn
    Dim SumSq As Double
    Dim I As Integer
    Dim N As Integer
    Dim EVec() As Double
    Dim AMat As Variant
    Dim Anext As Variant
    
    AMat = Matrix
    N = Sqr(WorksheetFunction.Count(AMat))
    SumSq = MatrixUTSumSq(AMat)
    Do While SumSq > Tolerance
        Anext = JacobiAmat(N, AMat)
        SumSq = MatrixUTSumSq(Anext)
        AMat = Anext
    Loop
    ReDim EVec(1 To N)
    For I = 1 To N
        EVec(I) = AMat(I, I)
    Next I
    EigenValues = EVec
End Function

Public Function Cholesky(A() As Double) As Double()

' Usage: Y = chol(A)
'
' Returns the upper triangular Cholesky root of A. A must by square, symmetric
' and positive definite.
' From http://www.howardrudd.net/pages/mc_maths_functions.html
' Copyright 2015 Howard J Rudd
' Licensed under the Apache License, Version 2.0

Dim I As Long
Dim J As Long
Dim k As Long
Dim M As Long
Dim N As Long
Dim G() As Double

'Determine the number of rows, n, and number of columns, m, in A
N = UBound(A, 1) - LBound(A, 1) + 1
M = UBound(A, 2) - LBound(A, 2) + 1

'TESTS

' Check that A is indexed from 1
If Not (LBound(A, 1) = 1 And LBound(A, 2) = 1) Then
    Err.Raise Number:=vbObjectError + 10, _
        Source:="chol", _
        Description:="Matrix not indexed from 1"
    Debug.Print "column index starts from" & LBound(A, 1) & " and row index starts from" & LBound(A, 1)
End If

'Check that A is square
If Not (N = M) Then
    Debug.Print "Attempted to perform Cholesky factorisation on a matrix" _
                & "that is not square"
    Err.Raise Number:=vbObjectError + 11, _
              Source:="Cholesky", _
              Description:="Matrix not square"
    Exit Function
End If

'Check that A is at least 2 x 2
If Not WorksheetFunction.Min(N, M) >= 2 Then
    Debug.Print "Attempted to perform Cholesky factorisation on a 1 x 1 matrix"
    Err.Raise Number:=vbObjectError + 12, _
              Source:="Cholesky", _
              Description:="Input matrix only has one element"
    Exit Function
End If

'Check that A is symmetric
For I = 1 To N
    For J = 1 To M
        If Not A(I, J) = A(J, I) Then
               Debug.Print "Attempted to perform Cholesky factorisation on a matrix" _
                         & "that is not symmetric"
                Err.Raise Number:=vbObjectError + 13, _
                          Source:="Cholesky", _
                          Description:="Matrix not symmetric"
            Exit Function
        End If
    Next J
Next I

' Check that the first diagonal element of A is >= 0
If A(1, 1) <= 0 Then
    Debug.Print "Attempted to perform Cholesky factorisation on a matrix" & _
                "that is not positive definite"
    Err.Raise Number:=vbObjectError + 14, _
                Source:="Cholesky", _
                Description:="Matrix not positive definite"
    Exit Function
End If

' END OF TESTS (almost). Actual maths starts here!

ReDim G(1 To N, 1 To M)

' Calculate 1st element
G(1, 1) = Sqr(A(1, 1))

' Calculate remainder of 1st row
For J = 2 To M
    G(1, J) = A(1, J) / G(1, 1)
Next J

' Calculate remaining rows
For I = 2 To N
' Calculate diagonal element of row i
    G(I, I) = A(I, I)
    For k = 1 To I - 1
        G(I, I) = G(I, I) - G(k, I) * G(k, I)
    Next k
' Check that g(i,i) is > 0
    If G(I, I) <= 0 Then
        Debug.Print "Attempted to perform Cholesky factorisation on a matrix" & _
                    "that is not positive definite"
        Err.Raise Number:=vbObjectError + 15, _
                  Source:="Cholesky", _
                  Description:="Matrix not positive definite"
        Exit Function
    End If
    G(I, I) = Sqr(G(I, I))
' Calculate remaining elements of row i
    For J = I + 1 To M
        G(I, J) = A(I, J)
        For k = 1 To I - 1
            G(I, J) = G(I, J) - G(k, I) * G(k, J)
        Next k
        G(I, J) = G(I, J) / G(I, I)
    Next J
Next I

' Calculate lower triangular half of G
For I = 2 To N
    For J = 1 To I - 1
        G(I, J) = 0
    Next J
Next I

Cholesky = G

End Function

Public Function ImanConover(Xascending() As Double, C() As Double) As Double()

' Usage: y = ic(Xascending, C)
'
' Performs the Iman-Conover method on Xind and returns a matrix Xcorr with the same
' dimensions as Xind.
'
' Xascending is an n-instance sample from an m-element random row-vector, with
' each column sorted in ascending order.
'
' If the columns of Xascending are in ascending order then the correlation
' matrix of Xcorr will be approximately equal to C. This function does not test
' Xascending to check that its columns are in ascending order. To do so would
' incur too large a computational burden.
'
' C must be square, symmetric and positive definite.
'
' The number of rows and columns of C must equal the number of columns Xascending.

    Dim nRowsC As Long
    Dim nColsC As Long
    Dim nRowsX As Long
    Dim nColsX As Long
    
    nRowsC = UBound(C, 1) - LBound(C, 1) + 1
    nColsC = UBound(C, 2) - LBound(C, 2) + 1
    nRowsX = UBound(Xascending, 1) - LBound(Xascending, 1) + 1
    nColsX = UBound(Xascending, 2) - LBound(Xascending, 2) + 1
    
    Dim I As Long
    Dim J As Long
    Dim k As Long
    
    Dim EX() As Double
    ReDim EX(1 To nColsX, 1 To nColsX)
    Dim FX() As Double
    ReDim FX(1 To nColsX, 1 To nColsX)
    Dim ZX() As Double
    ReDim ZX(1 To nColsX, 1 To nColsX)
    Dim S() As Double
    ReDim S(1 To nColsX, 1 To nColsX)
    
    Dim MX() As Double
    ReDim MX(1 To nRowsX, 1 To nColsX)
    Dim TX() As Double
    ReDim TX(1 To nRowsX, 1 To nColsX)
    Dim YX() As Double
    ReDim YX(1 To nRowsX, 1 To nColsX)
    Dim Ranks() As Long
    ReDim Ranks(1 To nRowsX, 1 To nColsX)
    
    ' TESTS!
    
    ' Test if C is square
    If Not nRowsC = nColsC Then
        MsgBox Title:="Iman-Conover Function", _
               prompt:="Correlation matrix is not square"
    End If
      
    ' Test if C is symmetric
    For I = 1 To nRowsC
        For J = I To nColsC
            If Abs(C(I, J) - C(J, I)) >= 10 ^ (-16) Then
                MsgBox Title:="Iman-Conover Function", _
                    prompt:="Correlation matrix is not symmetric"
                Exit Function
            End If
        Next J
    Next I
    
    ' Test if the number of rows of C is greater than the number of columns of X
    If nRowsC > nColsX Then
        MsgBox Title:="Iman-Conover Function", _
               prompt:="Correlation matrix too large"
    End If
    
    ' Test if the number of rows of C is less than the number of columns of X
    If nRowsC < nColsX Then
        MsgBox Title:="Iman-Conover Function", _
               prompt:="Correlation matrix too small"
    End If
    
    ' END OF TESTS: Actual maths starts here!
    
    ' Calculate the upper triangular Cholesky root of C
    S = Cholesky(C)
    
    ' Calculate the matrix, MX, of "normal scores"
    MX = NormalScores(nRowsX, nColsX)
    
    ' Calculate the matrix EX = MX' * MX
    EX = MatTransMult(MX, MX)
    
    ' Calculate Fx, the Cholesky root of Ex.
    FX = Cholesky(EX)
    
    ' Calculate ZX = FX^{-1) * S
    ZX = FInvS(FX, S)
    
    ' Calculate TX, the reordered matrix of "scores".
    TX = MatMult(MX, ZX)
    
    ' Calculate the rank orders of TX.
    Ranks = ArrayRank(TX)
    
    ' Reorder columns of X to match T.
    For J = 1 To nColsX
        For k = 1 To nRowsX
            YX(k, J) = Xascending(Ranks(k, J), J)
        Next k
    Next J
    
    ImanConover = YX
End Function

Public Function NormalScores(N As Long, M As Long) As Double()
    ' Produce Samples of the standard normal distribution.  Uses anthithetic sampling
    Dim Omega() As Double
    
    ReDim Omega(1 To N, 1 To M)
    
    Dim I As Long
    Dim J As Long
    Dim k As Long
    Dim x As Double
    Dim NormalisationFactor As Double
    
    x = 0
    For I = 1 To N \ 2
        Omega(I, 1) = WorksheetFunction.NormInv((I / (N + 1)), 0, 1)
        x = x + Omega(I, 1) ^ 2
    Next I
    
    NormalisationFactor = Sqr(2 * x / N)
    
    For I = 1 To N \ 2
        Omega(I, 1) = Omega(I, 1) / NormalisationFactor
    Next I
    
    k = Int(N / 2) + 1
    
    If Not 2 * Int(N / 2) = N Then
        Omega(k, 1) = 0
        For I = k + 1 To N
            Omega(I, 1) = -Omega(N - I + 1, 1)
        Next I
    Else
        For I = k To N
            Omega(I, 1) = -Omega(N - I + 1, 1)
        Next I
    End If
    
    For J = 2 To M
        For I = 1 To N
            Omega(I, J) = Omega(I, 1)
        Next I
    Next J
    
    For I = 1 To M
        Call Shuffle(Omega, I)
    Next I
    
    NormalScores = Omega
End Function

Public Function MatTransMult(A() As Double, B() As Double) As Double()

'Usage: C = mattransmult(A, B)
'
' Returns the product of A-transpose and B. Start indices of A and B
' can be arbitrary but start indices of the product C are both 1.

    Dim startIndexRowsA As Long
    Dim endIndexRowsA As Long
    Dim startIndexColsA As Long
    Dim endIndexColsA As Long
    Dim startIndexRowsB As Long
    Dim endIndexRowsB As Long
    Dim startIndexColsB As Long
    Dim endIndexColsB As Long
    Dim nRowsA As Long
    Dim nColsA As Long
    Dim nRowsB As Long
    Dim nColsB As Long
    Dim nRowsC As Long
    Dim nColsC As Long
    Dim I As Long
    Dim J As Long
    Dim k As Long
    
    startIndexRowsA = LBound(A, 1)
    endIndexRowsA = UBound(A, 1)
    startIndexColsA = LBound(A, 2)
    endIndexColsA = UBound(A, 2)
    startIndexRowsB = LBound(B, 1)
    endIndexRowsB = UBound(B, 1)
    startIndexColsB = LBound(B, 2)
    endIndexColsB = UBound(B, 2)
    
    nRowsA = endIndexRowsA - startIndexRowsA + 1
    nColsA = endIndexColsA - startIndexColsA + 1
    nRowsB = endIndexRowsB - startIndexRowsB + 1
    nColsB = endIndexColsB - startIndexColsB + 1
    
    ' Test that the two matrices are conformable
    If Not nRowsA = nRowsB Then
        Debug.Print "Attempted to multiply non conformable matrices"
        Err.Raise Number:=vbObjectError + 25, _
                  Source:="MatTransMult", _
                  Description:="Matrices not conformable"
        Exit Function
    End If
              
    nRowsC = nColsA
    nColsC = nColsB
    
    Dim C() As Double
    ReDim C(1 To nRowsC, 1 To nColsC)
    
    For I = 1 To nColsA
        For J = 1 To nColsB
            C(I, J) = 0
            For k = 1 To nRowsA
                C(I, J) = C(I, J) + A(k + startIndexRowsA - 1, I + startIndexColsA - 1) * _
                                    B(k + startIndexRowsB - 1, J + startIndexColsB - 1)
            Next k
        Next J
    Next I
    
    MatTransMult = C
End Function


Public Function FInvS(F() As Double, S() As Double) As Double()

' Usage: Y = FInvS(F, S)
'
' Returns the product of F^-1 and S for F and S both upper triangular. Actually
' solves FZ = S, i.e. finds Z such that FZ = S.

    Dim I As Long
    Dim J As Long
    Dim k As Long
    Dim M As Long
    Dim N As Long
    
    Dim nRowsF As Long
    Dim nColsF As Long
    Dim nRowsS As Long
    Dim nColsS As Long
    
    Dim z() As Double
    Dim w As Double
    
    nRowsF = UBound(F, 1) - LBound(F, 1) + 1
    nColsF = UBound(F, 2) - LBound(F, 2) + 1
    nRowsS = UBound(S, 1) - LBound(S, 1) + 1
    nColsS = UBound(S, 2) - LBound(S, 2) + 1
    
    ' TESTS
    
    ' Check that F is indexed from 1
    If Not (LBound(F, 1) = 1 And LBound(F, 2) = 1) Then
        Err.Raise Number:=vbObjectError + 16, _
            Source:="FInvS", _
            Description:="Matrix F not indexed from 1"
        Debug.Print "column index starts from" & LBound(F, 1) & " and row index starts from" & LBound(F, 1)
    End If
    
    ' Check that S is indexed from 1
    If Not (LBound(S, 1) = 1 And LBound(S, 2) = 1) Then
        Err.Raise Number:=vbObjectError + 17, _
            Source:="FInvS", _
            Description:="Matrix S not indexed from 1"
        Debug.Print "column index starts from" & LBound(S, 1) & " and row index starts from" & LBound(F, 1)
    End If
    
    ' Test whether F is square
    If Not nRowsF = nColsF Then
        Debug.Print "Matrix F is not square"
        Err.Raise Number:=vbObjectError + 18, _
                  Source:="FInvS", _
                  Description:="Matrix F is not square"
        Exit Function
    End If
    
    ' Test whether S is square
    If Not nRowsS = nColsS Then
        Debug.Print "Matrix S is not square"
        Err.Raise Number:=vbObjectError + 19, _
                  Source:="FInvS", _
                  Description:="Matrix S is not square"
        Exit Function
    End If
    
    ' Test whether F and S have same dimensions
    If Not nRowsF = nRowsS And nColsF = nColsS Then
        Debug.Print "Matrices F and S have different dimensions"
        Err.Raise Number:=vbObjectError + 20, _
                  Source:="FInvS", _
                  Description:="Matrices F and S have different dimensions"
        Exit Function
    End If
    
    ' Test whether F is upper triangular
    For I = 1 To nRowsF
        For J = 1 To I - 1
            If Not (F(I, J) = 0 And (Not F(J, I) = 0)) Then
                Debug.Print "Matrix F is not upper triangular"
                Err.Raise Number:=vbObjectError + 21, _
                          Source:="FInvS", _
                          Description:="Matrix F not upper triangular"
                Exit Function
            End If
        Next J
    Next I
    
    ' Test whether S is upper triangular
    For I = 2 To nRowsS
        For J = 1 To I - 1
            If S(I, J) > 10 ^ (-16) Then
                Debug.Print "Matrix S is not upper triangular"
                Err.Raise Number:=vbObjectError + 22, _
                          Source:="FInvS", _
                          Description:="Matrix S not upper triangular"
                Exit Function
            End If
        Next J
    Next I
    
    ' Test whether F has all non-zero diagonal elements
    For I = 1 To nRowsF
            If F(I, I) = 0 Then
                Debug.Print "Matrix F has at least one zero diagonal element and so is not invertible"
                Err.Raise Number:=vbObjectError + 23, _
                          Source:="FInvS", _
                          Description:="Matrix F has at least one zero diagonal element and so is not invertible"
                Exit Function
            End If
    Next I
    
    ' END OF TESTS. Actual maths starts here!
    
    N = nRowsF
    
    ReDim z(1 To N, 1 To N)
    
    ' Construct the nth row of Z
    For J = 1 To N - 1
        z(N, J) = 0
    Next J
        z(N, N) = S(N, N) / F(N, N)
    
    ' Construct the rows of Z above the nth
    For I = N - 1 To 1 Step -1
        For J = 1 To N
            w = 0
            For k = I + 1 To N
                w = w + F(I, k) * z(k, J)
            Next k
            z(I, J) = (S(I, J) - w) / F(I, I)
        Next J
    Next I
    
    FInvS = z
End Function

Public Function MatMult(A() As Double, B() As Double) As Double()

'Usage: C = matmult(A, B)
'
' Returns the product of A and B. Start indices of A and B can be arbitrary but
' start indices of the product C are both 1.

    Dim startIndexRowsA As Long
    Dim endIndexRowsA As Long
    Dim startIndexColsA As Long
    Dim endIndexColsA As Long
    Dim startIndexRowsB As Long
    Dim endIndexRowsB As Long
    Dim startIndexColsB As Long
    Dim endIndexColsB As Long
    Dim nRowsA As Long
    Dim nColsA As Long
    Dim nRowsB As Long
    Dim nColsB As Long
    Dim nRowsC As Long
    Dim nColsC As Long
    Dim I As Long
    Dim J As Long
    Dim k As Long
    
    startIndexRowsA = LBound(A, 1)
    endIndexRowsA = UBound(A, 1)
    startIndexColsA = LBound(A, 2)
    endIndexColsA = UBound(A, 2)
    startIndexRowsB = LBound(B, 1)
    endIndexRowsB = UBound(B, 1)
    startIndexColsB = LBound(B, 2)
    endIndexColsB = UBound(B, 2)
    
    nRowsA = endIndexRowsA - startIndexRowsA + 1
    nColsA = endIndexColsA - startIndexColsA + 1
    nRowsB = endIndexRowsB - startIndexRowsB + 1
    nColsB = endIndexColsB - startIndexColsB + 1
    
    ' Test that the two matrices are conformable
    If Not nColsA = nRowsB Then
        Debug.Print "Attempted to multiply non conformable matrices"
        Err.Raise Number:=vbObjectError + 24, _
                  Source:="MatMult", _
                  Description:="Matrices not conformable"
        Exit Function
    End If
              
    nRowsC = nRowsA
    nColsC = nColsB
    
    Dim C() As Double
    ReDim C(1 To nRowsC, 1 To nColsC)
    
    For I = 1 To nRowsC
        For J = 1 To nColsC
            C(I, J) = 0
            For k = 1 To nColsA
                C(I, J) = C(I, J) + A(I + startIndexRowsA - 1, k + startIndexColsA - 1) * _
                                    B(k + startIndexRowsB - 1, J + startIndexColsB - 1)
            Next k
        Next J
    Next I
    
    MatMult = C
End Function

Public Function ArrayRank(vArray() As Double) As Long()

' Usage: Y = arrayrank(vArray)
'
' Returns an array of longs representing the the rank orders of the elements
' of vArray. If vArray is two-dimensional, then it returns an array with
' the same number of rows and columns with the columns containing the ranks
' of the corresponding columns of "vArray".

    Dim vDims As Long
    Dim I As Long
    Dim J As Long
    Dim k As Long
    Dim M As Long
    Dim N As Long
    Dim vRank() As Long
    Dim tmpRank() As Long
    
    ' Check that dimensionality of vArray is either 1 or 2
    vDims = NumberOfArrayDimensions(vArray)
    If Not (vDims = 1 Or vDims = 2) Then
        Err.Raise Number:=vbObjectError + 7, _
                  Source:="ArrayRank", _
                  Description:="Input argument not one or two dimensional"
        Debug.Print "Arrayrank input argument not one or two dimensional"
        Debug.Print vDims
        Exit Function
    End If
    
    ' Check that array indices are numbered from 1
    If vDims = 1 Then
        If Not (LBound(vArray) = 1) Then
            Err.Raise Number:=vbObjectError + 8, _
                      Source:="ArrayRank", _
                      Description:="One dimensional array input argument not numbered from 1"
            Debug.Print "ArrayRank input argument start index = " & LBound(vArray)
            Exit Function
        End If
        N = UBound(vArray)
    ElseIf vDims = 2 Then
        If Not (LBound(vArray, 1) = 1 And LBound(vArray, 2) = 1) Then
            Err.Raise Number:=vbObjectError + 9, _
                      Source:="ArrayRank", _
                      Description:="Two dimensional array input argument not numbered from 1"
            Debug.Print "column index starts from" & LBound(vArray, 1) & " and row index starts from" & LBound(vArray, 1)
        End If
        N = UBound(vArray, 1) - LBound(vArray, 1) + 1
        M = UBound(vArray, 2) - LBound(vArray, 2) + 1
    End If
    
    ' Create array of consecutive longs from startIndex to endIndex
    If vDims = 1 Then
        ReDim vRank(1 To N)
        ReDim tmpRank(1 To N)
    ElseIf vDims = 2 Then
        ReDim vRank(1 To N, 1 To M)
        ReDim tmpRank(1 To N)
    End If
    
    ' Pass this array to QuickSort along with the array to be ranked
    If vDims = 1 Then
        For I = 1 To N
            tmpRank(I) = I
        Next I
            Call QuickSort(KeyArray:=vArray, OtherArray:=tmpRank)
        For I = 1 To N
            vRank(tmpRank(I)) = I
        Next I
    ElseIf vDims = 2 Then
        For J = 1 To M
            For I = 1 To N
                tmpRank(I) = I
            Next I
                Call QuickSort(KeyArray:=vArray, Column:=J, OtherArray:=tmpRank)
            For I = 1 To N
                vRank(tmpRank(I), J) = I
            Next I
        Next J
    End If
    
    ArrayRank = vRank
End Function

Private Function QS1ds(InLow As Long, InHi As Long, KeyArray() As Double)

' Usage: QuickSort(inLow, inHi, keyArray)
'
' "QS1ds" = quick sort one dimensional single array
'
' Sorts keyArray in place between indices inLow and inHi, keyArray must be of
' type Double
'
' This function is based on code suggested in:
' http://stackoverflow.com/questions/152319/vba-array-sort-function

    Dim tmpLow As Long
    Dim tmpHi As Long
    Dim pivot As Double
    Dim keyTmpSwap As Double
    
    tmpLow = InLow
    tmpHi = InHi
    pivot = KeyArray((InLow + InHi) \ 2)
    
    Do While (tmpLow <= tmpHi)
        Do While KeyArray(tmpLow) < pivot And tmpLow < InHi
            tmpLow = tmpLow + 1
        Loop
        Do While KeyArray(tmpHi) > pivot And tmpHi > InLow
            tmpHi = tmpHi - 1
        Loop
        If (tmpLow <= tmpHi) Then
            keyTmpSwap = KeyArray(tmpLow)
            KeyArray(tmpLow) = KeyArray(tmpHi)
            KeyArray(tmpHi) = keyTmpSwap
            tmpLow = tmpLow + 1
            tmpHi = tmpHi - 1
        End If
    Loop
    
    If (InLow < tmpHi) Then QS1ds InLow, tmpHi, KeyArray
    If (tmpLow < InHi) Then QS1ds tmpLow, InHi, KeyArray
End Function

Private Function QS1dd(InLow As Long, InHi As Long, KeyArray() As Double, OtherArray As Variant)

' Usage: QS1dd(inLow, inHi, keyArray, otherArray)
'
' "QS1dd" = quick sort one dimensional double array
'
' Sorts keyArray in place between indices inLow and inHi.
'
' An array, "otherArray" is sorted in parallel, also in-place. "otherArray" must
' be one dimensional and have the same start and end indices as the first dimen-
' sion of "keyArray".
'
' "keyArray" must be of type Double, "otherArray" must be of type long.
'
' This function is based on code suggested in:
' http://stackoverflow.com/questions/152319/vba-array-sort-function

    Dim pivot As Double
    Dim tmpLow As Long
    Dim tmpHi As Long
    Dim keyTmpSwap As Double
    Dim otherTmpSwap As Long
    Dim keyDims As Long
    Dim otherDims As Long
    
    tmpLow = InLow
    tmpHi = InHi
    pivot = KeyArray((InLow + InHi) \ 2)
    
    Do While (tmpLow <= tmpHi)
    
        Do While KeyArray(tmpLow) < pivot And tmpLow < InHi
            tmpLow = tmpLow + 1
        Loop
        
        Do While KeyArray(tmpHi) > pivot And tmpHi > InLow
            tmpHi = tmpHi - 1
        Loop
        
        If (tmpLow <= tmpHi) Then
        
            keyTmpSwap = KeyArray(tmpLow)
            otherTmpSwap = OtherArray(tmpLow)
            
            KeyArray(tmpLow) = KeyArray(tmpHi)
            OtherArray(tmpLow) = OtherArray(tmpHi)
            
            KeyArray(tmpHi) = keyTmpSwap
            OtherArray(tmpHi) = otherTmpSwap
            
            tmpLow = tmpLow + 1
            tmpHi = tmpHi - 1
            
        End If
        
    Loop
    
    If (InLow < tmpHi) Then QS1dd InLow, tmpHi, KeyArray, OtherArray
    If (tmpLow < InHi) Then QS1dd tmpLow, InHi, KeyArray, OtherArray
End Function

Private Function QS2ds(InLow As Long, InHi As Long, KeyArray() As Double, Column As Long)

' Usage: QS2ds(inLow, inHi, keyArray, column)
'
' "QS2ds" = quick sort two dimensional single array
'
' Sorts a column of "keyArray", in place ' between indices "inLow" and "inHi".
' "keyArray" must be is two-dimensional (i.e. a matrix). Only the column
' specified by the optional argument "column" will be sorted, the other columns
' are left untouched.
'
' "keyArray" must be of type Double, column and "column" must be of type
' long
'
' This function is based on code suggested in:
' http://stackoverflow.com/questions/152319/vba-array-sort-function

    Dim pivot As Double
    Dim tmpLow As Long
    Dim tmpHi As Long
    Dim keyTmpSwap As Double
    Dim keyDims As Long
    
    tmpLow = InLow
    tmpHi = InHi
    pivot = KeyArray((InLow + InHi) \ 2, Column)
    
    Do While (tmpLow <= tmpHi)
        Do While KeyArray(tmpLow, Column) < pivot And tmpLow < InHi
            tmpLow = tmpLow + 1
        Loop
        Do While KeyArray(tmpHi, Column) > pivot And tmpHi > InLow
            tmpHi = tmpHi - 1
        Loop
        If (tmpLow <= tmpHi) Then
            keyTmpSwap = KeyArray(tmpLow, Column)
            KeyArray(tmpLow, Column) = KeyArray(tmpHi, Column)
            KeyArray(tmpHi, Column) = keyTmpSwap
            tmpLow = tmpLow + 1
            tmpHi = tmpHi - 1
        End If
    Loop
    
    If (InLow < tmpHi) Then QS2ds InLow, tmpHi, KeyArray, Column
    If (tmpLow < InHi) Then QS2ds tmpLow, InHi, KeyArray, Column
End Function

Private Function QS2dd(InLow As Long, InHi As Long, KeyArray() As Double, Column As Long, OtherArray As Variant)

' Usage: qs2dd(inLow, inHi, keyArray, column, otherArray)
'
' "qs1dd" = quick sort two dimensional double array
'
' Sorts "keyArray" and "otherArray" in place between indices inLow and inHi.
'
' An array, "otherArray" is sorted in parallel, also in-place. "otherArray" must
' be one dimensional and have the same start and end indices as the first dimen-
' sion of "keyArray".
'
' "KeyArray" must be of type Double, "Column" and "OtherArray" must be of type
' long.
'
' This function is based on code suggested in:
' http://stackoverflow.com/questions/152319/vba-array-sort-function

    Dim pivot As Double
    Dim tmpLow As Long
    Dim tmpHi As Long
    Dim keyTmpSwap As Double
    Dim otherTmpSwap As Long
    Dim keyDims As Long
    Dim otherDims As Long
    
    tmpLow = InLow
    tmpHi = InHi
    pivot = KeyArray((InLow + InHi) \ 2, Column)
    
    Do While (tmpLow <= tmpHi)
    
        Do While KeyArray(tmpLow, Column) < pivot And tmpLow < InHi
            tmpLow = tmpLow + 1
        Loop
        
        Do While KeyArray(tmpHi, Column) > pivot And tmpHi > InLow
            tmpHi = tmpHi - 1
        Loop
        
        If (tmpLow <= tmpHi) Then
        
            keyTmpSwap = KeyArray(tmpLow, Column)
            otherTmpSwap = OtherArray(tmpLow)
            
            KeyArray(tmpLow, Column) = KeyArray(tmpHi, Column)
            OtherArray(tmpLow) = OtherArray(tmpHi)
            
            KeyArray(tmpHi, Column) = keyTmpSwap
            OtherArray(tmpHi) = otherTmpSwap
            
            tmpLow = tmpLow + 1
            tmpHi = tmpHi - 1
            
        End If
        
    Loop
    
    If (InLow < tmpHi) Then QS2dd InLow, tmpHi, KeyArray, Column, OtherArray
    If (tmpLow < InHi) Then QS2dd tmpLow, InHi, KeyArray, Column, OtherArray
End Function

Public Function QuickSort(KeyArray() As Double, Optional Column As Long, Optional OtherArray As Variant)

    ' Usage: QuickSort(keyArray, column, otherArray)
    '
    ' Sorts keyArray in place. If keyArray is two-dimensional (i.e. a matrix) then
    ' only the column specified by the optional argument "column" will be sorted.
    '
    ' An optional "otherArray" can be sorted in parallel, also in-place. If
    ' "otherArray" is used it must be one-dimensional and have the same start and
    ' end indices as the columns of "keyArray".
    '
    ' "keyArray" must be of type Double, "column" and otherArray must be of type
    ' long
    
    Dim keyDims As Long
    Dim otherDims As Long
    Dim InLow As Long
    Dim InHi As Long
    Dim I As Long
    
    ' TESTS
    
    ' Check that the dimensionality of "keyArray" is either 1 or 2
    keyDims = NumberOfArrayDimensions(KeyArray)
    If Not (keyDims = 1 Or keyDims = 2) Then
        Debug.Print "input argument not one or two dimensional"
        Exit Function
    End If
    
    If Not IsMissing(OtherArray) Then
    ' Check that "otherArray" is one-dimensional
        otherDims = NumberOfArrayDimensions(OtherArray)
        If Not otherDims = 1 Then
            Debug.Print "'otherArray' not one-dimensional"
            Exit Function
        End If
        If keyDims = 1 Then
    ' Check that "keyArray" and "otherArray" are conformable
            If Not ( _
                        UBound(KeyArray) = UBound(OtherArray) And _
                        LBound(KeyArray) = LBound(OtherArray) _
                    ) Then
                    Err.Raise Number:=vbObjectError + 40, _
                        Source:="QuickSort", _
                        Description:="'keyArray' and 'otherArray' not conformable"
                    Exit Function
            End If
        ElseIf keyDims = 2 Then
    ' Check that the argument "column" has been supplied
            If IsMissing(Column) Then
                Err.Raise Number:=vbObjectError + 41, _
                    Source:="QuickSort", _
                    Description:="'column' argument not passed"
                Exit Function
            End If
    ' Check that "otherArray" is conformable to columns of "keyArray"
            If Not ( _
                    UBound(KeyArray, 1) = UBound(OtherArray) And _
                    LBound(KeyArray, 1) = LBound(OtherArray) _
                ) Then
                Err.Raise Number:=vbObjectError + 42, _
                    Source:="QuickSort", _
                    Description:="'keyArray' and 'otherArray' not conformable"
                Exit Function
            End If
        End If
    End If
    
    ' Check that the argument "column" points to one of the columns of "keyArray"
    If keyDims = 2 Then
        If Not (LBound(KeyArray, 2) <= Column And Column <= UBound(KeyArray, 2)) Then
            Err.Raise Number:=vbObjectError + 43, _
                Source:="QuickSort", _
                Description:="Argument 'column' does not point to one of the columns of 'keyArray'"
            Exit Function
        End If
    End If
    
    ' END OF TESTS
    
    ' Calculate "inHi" and "inLow"
    
    If keyDims = 1 Then
        InLow = LBound(KeyArray)
        InHi = UBound(KeyArray)
    ElseIf keyDims = 2 Then
        InLow = LBound(KeyArray, 1)
        InHi = UBound(KeyArray, 1)
    End If
    
    ' Call appropriate sort function
    
    If keyDims = 1 And IsMissing(OtherArray) Then
        Call QS1ds(InLow, InHi, KeyArray)
            
    ElseIf keyDims = 1 And Not IsMissing(OtherArray) Then
        Call QS1dd(InLow, InHi, KeyArray, OtherArray)
    
    ElseIf keyDims = 2 And IsMissing(OtherArray) Then
        Call QS2ds(InLow, InHi, KeyArray, Column)
    
    ElseIf keyDims = 2 And Not IsMissing(OtherArray) Then
        Call QS2dd(InLow, InHi, KeyArray, Column, OtherArray)
    
    End If
End Function

