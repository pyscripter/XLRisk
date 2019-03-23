Attribute VB_Name = "Utilities"
Option Explicit
Option Base 1

Sub CollectRiskInputs(Coll As Collection)
'  Adds all risk inputs of the ActiveWorkbook (cell with risk formulas) to Coll
'  On Exit The collection contains the risk input cells
    Dim Sht As Worksheet
    Dim Formulas As Range
    Dim Cell As Range
    Dim RiskFunction As Variant
    Dim FunctionList As Variant
        
    FunctionList = RiskFunctionList()
        
    For Each Sht In ActiveWorkbook.Worksheets 'loop through the sheets in the workbook
        On Error Resume Next 'in case there are no formulas
        'Limit the search to the UsedRange and use SpecialCells to reduce looping further
        Set Formulas = Sht.UsedRange.SpecialCells(xlCellTypeFormulas)
        If Err = 0 Then
            For Each Cell In Formulas 'loop through the SpecialCells only
                '  Check whether the formula contains a Risk function
                For Each RiskFunction In FunctionList
                    If Cell.HasFormula And InStr(1, Cell.Formula, RiskFunction, vbTextCompare) > 0 Then
                        Coll.Add Cell
                        Exit For
                    End If
               Next RiskFunction
            Next Cell
        End If
        Err.Clear
        Set Formulas = Nothing
    Next Sht
End Sub


Public Function InputCells() As Variant
' Returns an array of cells containing input formulas in the Active Workbook
' Can be used as an array function in Excel
    Dim I As Integer
    Dim Coll As New Collection
    Dim Cell As Range
    Dim Result() As Variant
    CollectRiskInputs Coll
    
    ' Convert collection to an array
    ReDim Result(Coll.Count, 2)
    For I = 1 To Coll.Count
      Set Cell = Coll(I)
      Result(I, 1) = QuoteIfNeeded(Cell.Parent.Name) & "!" & Cell.Address
      Result(I, 2) = Right(Cell.Formula, Len(Cell.Formula) - 1)
    Next I
    
    InputCells = Result
End Function

Sub CollectRiskOutputs(Coll As Collection)
'  Adds all risk outputs of the ActiveWorkbook to Coll
'  The collection contains pairs (name, output cells)
'  Assumes XLRisk sheet exists
    Dim Sht As Worksheet
    Dim R As Range
    Dim RiskOutput As Range
    Dim Row As Integer
    Dim Cell As Range
    
    Set Sht = ActiveWorkbook.Worksheets("XLRisk")
    Set R = Sht.Range("RiskOutputs").CurrentRegion
    
    For Row = 2 To R.Rows.Count
        Set RiskOutput = Range(R.Cells(Row, 1))
        For Each Cell In RiskOutput
          Coll.Add Array(R.Cells(Row, 2), Cell)
        Next Cell
    Next Row
End Sub


Sub ThickBorders(R As Range)
    With R
        R.Borders(xlEdgeTop).LineStyle = xlContinuous
        R.Borders(xlEdgeTop).Weight = xlMedium
        R.Borders(xlEdgeBottom).LineStyle = xlContinuous
        R.Borders(xlEdgeBottom).Weight = xlMedium
        R.Borders(xlEdgeLeft).LineStyle = xlContinuous
        R.Borders(xlEdgeLeft).Weight = xlMedium
        R.Borders(xlEdgeRight).LineStyle = xlContinuous
        R.Borders(xlEdgeRight).Weight = xlMedium
    End With
End Sub

Function QuoteIfNeeded(S As String) As String
    If InStr(S, " ") > 0 Then
        QuoteIfNeeded = "'" & S & "'"
    Else
        QuoteIfNeeded = S
    End If
End Function

Function NameOrAddress(R As Range) As String
    On Error Resume Next
    NameOrAddress = R.Name.Name
    If NameOrAddress = "" Then NameOrAddress = QuoteIfNeeded(R.Parent.Name) & "!" & R.Address
End Function

