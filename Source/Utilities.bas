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
    Dim RiskInput As ClsRiskInput
    Dim Formula As String
    Dim Pos As Long
        
    FunctionList = RiskFunctionList()
        
    For Each Sht In ActiveWorkBook.Worksheets 'loop through the sheets in the workbook
        On Error GoTo Error 'in case there are no formulas
        'Limit the search to the UsedRange and use SpecialCells to reduce looping further
        Set Formulas = Sht.UsedRange.SpecialCells(xlCellTypeFormulas)
        For Each Cell In Formulas 'loop through the SpecialCells only
            '  Check whether the formula contains a Risk function
            Formula = Cell.Formula
            Pos = InStr(1, Formula, "risk", vbTextCompare)
            If Pos > 0 Then
                For Each RiskFunction In FunctionList
                     If InStr(Pos, Formula, RiskFunction, vbTextCompare) > 0 Then
                         Set RiskInput = New ClsRiskInput
                         RiskInput.Init Cell
                         Call Coll.Add(RiskInput, AddressWithSheet(Cell))
                         Exit For
                     End If
                Next RiskFunction
            End If
        Next Cell
        Set Formulas = Nothing
NextSheet:
    Next Sht
    Exit Sub
Error:
    Resume NextSheet
End Sub

Public Function OneRiskFunctionPerCell(Coll As Collection) As Boolean
    Dim RiskInput As ClsRiskInput
    Dim FunctionList As Variant
    Dim RiskFunction As Variant
    Dim Count As Integer
    Dim Pos As Integer
    
    OneRiskFunctionPerCell = True
    FunctionList = RiskFunctionList()
    For Each RiskInput In Coll
        Count = 0
        For Each RiskFunction In FunctionList
            Pos = 1
            Do Until Pos = 0
                Pos = InStr(Pos, RiskInput.Cell.Formula, RiskFunction, vbTextCompare)
                If Pos > 0 Then
                    Count = Count + 1
                    Pos = Pos + Len(RiskFunction)
                End If
            If Count > 1 Then
                OneRiskFunctionPerCell = False
                MsgBox "XLRisk allows only one risk function per cell. Cell " & AddressWithSheet(RiskInput.Cell) & _
                    " contains more than one risk function", vbExclamation, _
                    "Multiple risk functions in cell"
                Exit Function
            End If
            Loop
        Next RiskFunction
    Next RiskInput
End Function

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
      Set Cell = Coll(I).Cell
      Result(I, 1) = AddressWithSheet(Cell)
      Result(I, 2) = Right(Cell.Formula, Len(Cell.Formula) - 1)
    Next I
    
    InputCells = Result
End Function

Public Sub CollectRiskOutputs(Coll As Collection)
'  Adds all risk outputs of the ActiveWorkbook to Coll
'  The collection contains pairs (name, output cells)
'  Assumes XLRisk sheet exists
    Dim Sht As Worksheet
    Dim R As Range
    Dim RiskOutputRange As Range
    Dim Row As Integer
    Dim Cell As Range
    Dim RiskOutput As ClsRiskOutput
    
    Set Sht = ActiveWorkBook.Worksheets("XLRisk")
    Set R = Sht.Range("RiskOutputs").CurrentRegion
    
    For Row = 2 To R.Rows.Count
        Set RiskOutputRange = Range(R.Cells(Row, 1).Value)
        For Each Cell In RiskOutputRange
            Set RiskOutput = New ClsRiskOutput
            RiskOutput.Init R.Cells(Row, 2).Value, Cell
            Coll.Add RiskOutput
        Next Cell
    Next Row
End Sub

Public Sub ThickBorders(R As Range)
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

Public Function QuoteIfNeeded(S As String) As String
    If S Like "*[!0-9a-zA-Z]*" Then
        QuoteIfNeeded = "'" & S & "'"
    Else
        QuoteIfNeeded = S
    End If
End Function

Public Function AddressWithSheet(R As Range) As String
    AddressWithSheet = QuoteIfNeeded(R.Parent.Name) & "!" & R.Address
End Function

Public Function NameOrAddress(R As Range) As String
    On Error Resume Next
    NameOrAddress = R.Name.Name
    If Len(NameOrAddress) = 0 Then NameOrAddress = AddressWithSheet(R)
End Function

Public Function SortedTable(Table As Range, Col As Long, Optional Ascending As Boolean = True, _
    Optional Absolute As Boolean = False) As Variant
    Dim ColVals() As Double
    Dim NRows As Long
    Dim NCols As Long
    Dim V As Variant
    Dim I As Long
    Dim J As Long
    Dim Ranks As Variant
    Dim CellVal As Variant
    
    NRows = Table.Rows.Count
    NCols = Table.Columns.Count
    
    'Error Checking
    If (NRows < 2) Or (NCols < 2) Or (Col < 1) Or (Col > NCols) Then
        SortedTable = CVErr(xlErrValue)
        Exit Function
    End If
    
    ReDim ColVals(1 To NRows)
    For I = 1 To NRows
        If Absolute Then
            ColVals(I) = Abs(Table.Cells(I, Col).Value)
        Else
            ColVals(I) = Table.Cells(I, Col).Value
        End If
    Next I
    
    ReDim V(1 To NRows, 1 To NCols)
    Ranks = ArrayRank(ColVals)
    For J = 1 To NCols
        For I = 1 To NRows
            CellVal = Table.Cells(I, J).Value
            '  Keep empty cells empty
            If Ascending Then
                V(Ranks(I), J) = IIf(IsEmpty(CellVal), "", CellVal)
            Else
                V(NRows + 1 - Ranks(I)) = IIf(IsEmpty(CellVal), "", CellVal)
            End If
        Next I
    Next J
    SortedTable = V
End Function

