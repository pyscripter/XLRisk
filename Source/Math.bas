Attribute VB_Name = "Math"

Public Function NumberOfArrayDimensions(Arr() As Double) As Long
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

Public Function Shuffle(vArray() As Double, Optional Column As Variant)

' Usage: shuffle(vArray, column)
'
' shuffles vArray in place randomly. If vArray is 2-dimensional, then it shuffles
' the column specified by the optional variable "column"
'
' From http://www.howardrudd.net/pages/mc_maths_functions.html
' Copyright 2015 Howard J Rudd
' Licensed under the Apache License, Version 2.0

Dim j As Long
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
    Exit Function
End If

If vDims = 1 Then
    startIndex = LBound(vArray)
    endIndex = UBound(vArray)
    j = endIndex
    Do While j > startIndex
        k = CLng(startIndex + Rnd() * (j - startIndex))
        temp = vArray(j)
        vArray(j) = vArray(k)
        vArray(k) = temp
        j = j - 1
    Loop
ElseIf vDims = 2 Then
    If IsMissing(Column) Then
        'ERROR: Argument 'column' not supplied
        Exit Function
    End If
' Check that the argument "column" points to one of the columns of "vArray"
    If Not (LBound(vArray, 2) <= Column And Column <= UBound(vArray, 2)) Then
        ' ERROR: Argument "column" does not point to one of the columns of "vArray"
        Exit Function
    End If
    startIndex = LBound(vArray, 1)
    endIndex = UBound(vArray, 1)
    j = endIndex
    Do While j > startIndex
        k = CLng(startIndex + Rnd() * (j - startIndex))
        temp = vArray(j, Column)
        vArray(j, Column) = vArray(k, Column)
        vArray(k, Column) = temp
        j = j - 1
    Loop
End If

End Function
