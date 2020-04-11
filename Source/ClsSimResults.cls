VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsSimResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mOutSheet As Worksheet
Private mRiskInputs As Collection
Private mRiskOutputs As Collection
Private mIterations As Long
Private mSimResults() As Variant

Private Const ChartHeight As Integer = 22 'rows
Private Const ChartWidth As Integer = 10 'columns


Private Sub Class_Initialize()
    Set mOutSheet = CreateOutputSheet
End Sub

Public Sub Init(RiskInputs As Collection, RiskOutputs As Collection, Iterations As Long, SimResults() As Variant)
    Set mRiskInputs = RiskInputs
    Set mRiskOutputs = RiskOutputs
    mIterations = Iterations
    mSimResults = SimResults
End Sub

Public Sub ProduceResults()
    mOutSheet.Activate
    InitialiseResults
    'Produce Output
    mOutSheet.Cells(4, 2).Resize(mIterations, 1 + mRiskInputs.Count + mRiskOutputs.Count) = mSimResults
    ' Produce Statistics
    ProduceStatistics
    ' Calculate before producing the graphs
    Application.Calculate
    ' Produce Cumulative Distributions
    ProduceCumulativeDistributions
    ' Produce histograms only if Excel version > 16
    If Val(Application.Version) >= 16 Then ProduceHistograms
    ' Produce Tornado Diagram
    If mRiskInputs.Count > 1 Then ProduceTornadoDiagram
End Sub

Private Sub InitialiseResults()
    Dim RiskInput As ClsRiskInput
    Dim Curr As Range
    Dim RiskOutput As ClsRiskOutput
    
    With mOutSheet
        .Range("C1").Value = "Inputs"
        .Range("B3").Value = "Iterations"
    End With
      
    '  Setup risk inputs
    Set Curr = mOutSheet.Range("C2")
    mOutSheet.Names.Add Name:="InputResults", RefersTo:=Curr.Offset(2)
    For Each RiskInput In mRiskInputs
        Curr = NameOrAddress(RiskInput.Cell)
        Curr.Offset(1, 0) = Right(RiskInput.Cell.Formula, Len(RiskInput.Cell.Formula) - 1)
        Set Curr = Curr.Offset(0, 1)
    Next RiskInput
    'Format Input
    With mOutSheet.Range("C1").Resize(1, mRiskInputs.Count)
        If mRiskInputs.Count > 1 Then .Merge
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    With mOutSheet.Range("InputResults")
        ThickBorders .Offset(0, -1).Resize(mIterations, 1)
        ThickBorders .Offset(-2, -1).Resize(2, 1)
        ThickBorders .Offset(-3).Resize(1, mRiskInputs.Count)
        ThickBorders .Resize(mIterations, mRiskInputs.Count)
        ThickBorders .Offset(-2).Resize(2, mRiskInputs.Count)
    End With
    
    '   Setup risk outputs
    Curr.Offset(-1, 0) = "Outputs"
    mOutSheet.Names.Add Name:="OutputResults", RefersTo:=Curr.Offset(2)
    For Each RiskOutput In mRiskOutputs
        Curr = AddressWithSheet(RiskOutput.Cell)
        Curr.Offset(1, 0) = RiskOutput.Name
        Set Curr = Curr.Offset(0, 1)
    Next RiskOutput
    'Format Output
    With mOutSheet.Range("C1").Offset(0, 1).Resize(1, mRiskOutputs.Count)
        If mRiskOutputs.Count > 1 Then .Merge
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    '   AutoFit
    mOutSheet.Range("C2").Resize(1, mRiskInputs.Count).Columns.AutoFit
    With mOutSheet.Range("C2").Offset(0, mRiskInputs.Count).Resize(2, mRiskOutputs.Count)
        .Columns.AutoFit
        .HorizontalAlignment = xlCenter
    End With
    With mOutSheet.Range("OutputResults")
        ThickBorders .Offset(-3).Resize(1, mRiskOutputs.Count)
        ThickBorders .Resize(mIterations, mRiskOutputs.Count)
        ThickBorders .Offset(-2).Resize(2, mRiskOutputs.Count)
    End With
    
    '   Setup Simulation Statistics
    Set Curr = Curr.Offset(0, 2)
    Curr.Offset(-1, 0) = "Simulation Statistics"
    mOutSheet.Names.Add Name:="SimStats", RefersTo:=Curr.Offset(2, -1)
    Curr.Resize(2, mRiskOutputs.Count).Value = _
        mOutSheet.Range("OutputResults").Offset(-2).Resize(2, mRiskOutputs.Count).Value
    'Format Simulation Results
    With mOutSheet.Range("SimStats").Offset(-3).Resize(1, mRiskOutputs.Count + 1)
        .Merge
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    'AutoFit
    With mOutSheet.Range("SimStats").Offset(-2, 1).Resize(2, mRiskOutputs.Count)
        .Columns.AutoFit
        .HorizontalAlignment = xlCenter
    End With
    With mOutSheet.Range("SimStats")
        ThickBorders .Offset(-3).Resize(1, mRiskOutputs.Count + 1)
        ThickBorders .Resize(14, 1)
        ThickBorders .Offset(0, 1).Resize(14, mRiskOutputs.Count)
        ThickBorders .Offset(-2, 1).Resize(2, mRiskOutputs.Count)
        ThickBorders .Offset(-2).Resize(16, mRiskOutputs.Count + 1)
    End With
    
    '  Setup Output Cumulative Distributions
    Set Curr = Curr.Offset(0, 1 + mRiskOutputs.Count)
    Curr.Offset(-1, 0) = "Output Cumulative Distributions"
    mOutSheet.Names.Add Name:="OutCumDist", RefersTo:=Curr.Offset(1, 0)
    With Curr.Offset(-1, 0).Resize(1, ChartWidth)
        .Merge
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    
    If Val(Application.Version) >= 16 Then
        '  Setup Output Histograms
        Set Curr = Curr.Offset(0, ChartWidth + 1)
        Curr.Offset(-1, 0) = "Output Distributions"
        mOutSheet.Names.Add Name:="OutDist", RefersTo:=Curr.Offset(1, 0)
        With Curr.Offset(-1, 0).Resize(1, ChartWidth)
            .Merge
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
    End If
    
    If mRiskInputs.Count > 1 Then
        '  Setup Tornado Diagram
        Set Curr = Curr.Offset(0, ChartWidth + 1)
        Curr.Offset(-1, 0) = "Risk Tornado Diagram"
        mOutSheet.Names.Add Name:="Tornado", RefersTo:=Curr.Offset(1, 0)
        With Curr.Offset(-1, 0).Resize(1, ChartWidth)
            .Merge
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
    End If
    
    ' Add Hyperlinks
    Set Curr = mOutSheet.Range("A1")
    With Curr
        .Value = "Navigation"
        .Font.Bold = True
    End With
    
    With mOutSheet.Hyperlinks
        .Add Anchor:=Curr.Offset(1), Address:=vbNullString, _
            SubAddress:=mOutSheet.Range("OutputResults").Offset(-3).Address, TextToDisplay:="Risk Outputs"
        .Add Anchor:=Curr.Offset(2), Address:=vbNullString, _
            SubAddress:=mOutSheet.Range("SimStats").Offset(-3).Address, TextToDisplay:="Simulation Statistics"
        .Add Anchor:=Curr.Offset(3), Address:=vbNullString, _
            SubAddress:="Percentiles", TextToDisplay:="Risk Output Percentiles"
        .Add Anchor:=Curr.Offset(4), Address:=vbNullString, _
            SubAddress:="PercentRank", TextToDisplay:="Calculate Percent Rank"
        .Add Anchor:=Curr.Offset(5), Address:=vbNullString, _
            SubAddress:="Correlations", TextToDisplay:="Correlations"
        .Add Anchor:=Curr.Offset(6), Address:=vbNullString, _
            SubAddress:=mOutSheet.Range("OutCumDist").Offset(-2).Address, TextToDisplay:="Output Cumulative Distributions"
        If Val(Application.Version) >= 16 Then
            .Add Anchor:=Curr.Offset(7), Address:=vbNullString, _
                SubAddress:=mOutSheet.Range("OutDist").Offset(-2).Address, TextToDisplay:="Output Distributions"
        End If
        If mRiskInputs.Count > 1 Then
            .Add Anchor:=Curr.Offset(IIf(Val(Application.Version) >= 16, 8, 7)), Address:=vbNullString, _
                SubAddress:=mOutSheet.Range("Tornado").Offset(-2).Address, TextToDisplay:="Risk Tornado Diagram"
        End If
    End With
    mOutSheet.Columns(1).AutoFit
End Sub

Private Sub StatHelper(Cell As Range, StatName As String, StatFormula As String, Address As String)
    Cell = StatName
    Cell.Offset(0, 1).Formula = "=" & StatFormula & "(" & Address & ")"
End Sub

Private Sub StatAggregateHelper(Cell As Range, StatName As String, FormulaCode As Integer, Address As String)
    Cell = StatName
    Cell.Offset(0, 1).Formula = "=AGGREGATE(" & CStr(FormulaCode) & ",6," & Address & ")"
End Sub

Private Sub ProduceStatistics()
'  Use the Aggregate Excel function so that stats can be calculated even if the outputs contain a few errors
    Dim FirstOutput As Range
    Dim Cell As Range
    Dim Address As String
    Dim I As Long
    
    Set FirstOutput = mOutSheet.Range("OutputResults").Resize(mIterations, 1)
    Set Cell = mOutSheet.Range("SimStats")
    
    Address = FirstOutput.Address(True, False)
    StatAggregateHelper Cell, "Mean", 1, Address
    StatAggregateHelper Cell.Offset(1), "Median", 12, Address
    StatAggregateHelper Cell.Offset(2), "Mode", 13, Address
    StatAggregateHelper Cell.Offset(3), "Std. Deviation", 7, Address
    StatAggregateHelper Cell.Offset(4), "Variance", 10, Address
    StatHelper Cell.Offset(5), "Kurtosis", "KURT", Address
    StatHelper Cell.Offset(6), "Skewness", "Skew", Address
    StatAggregateHelper Cell.Offset(7), "Minimum", 5, Address
    StatAggregateHelper Cell.Offset(8), "Maximum", 4, Address
    Cell.Offset(9) = "Range"
    Cell.Offset(9, 1).Formula = "=" & Cell.Offset(8, 1).Address(False, False) & "-" & Cell.Offset(7, 1).Address(False, False)
    StatHelper Cell.Offset(10), "Count", "Count", Address
    Cell.Offset(11) = "Error Count"
    Cell.Offset(11, 1).FormulaArray = "=COUNT(IF(ISERROR(" & Address & "), 1, """"))"
    Cell.Offset(12) = "Std. Error"
    Cell.Offset(12, 1).Formula = "=" & Cell.Offset(3, 1).Address(False, False) & "/SQRT(" & Cell.Offset(10, 1).Address(False, False) & ")"
    Cell.Offset(13) = "Confidence Level (95%)"
    Cell.Offset(13, 1).Formula = "=CONFIDENCE.T(5%," & Cell.Offset(3, 1).Address(False, False) & "," & Cell.Offset(10, 1).Address(False, False) & ")"
    If mRiskOutputs.Count > 1 Then
        For I = 2 To mRiskOutputs.Count
            Cell.Offset(0, I).Resize(14, 1).FormulaR1C1 = Cell.Offset(0, 1).Resize(14, 1).FormulaR1C1
        Next I
    End If
    
    'Percentiles
    ProducePercentiles Cell.Offset(15)
    
    'Percent Rank
    ProducePercentRank Cell.Offset(60)
    
    'Correlations
    ProduceCorrelations Cell.Offset(66)
    
    'AutoFit
    mOutSheet.Columns(Cell.Column).AutoFit
End Sub

Private Sub ProducePercentiles(Cell As Range)
    Dim Perc As Integer
    Dim PCount As Integer
    Dim Address As String
    
    mOutSheet.Names.Add Name:="Percentiles", RefersTo:=Cell.Offset(3).Resize(41, mRiskOutputs.Count + 1)
    Perc = 0
    For PCount = 1 To 41
        With Cell.Offset(PCount + 2)
            .Value = Perc / 1000
            .NumberFormat = "0.0%"
            .HorizontalAlignment = xlRight
        End With
        Perc = Perc + 25
    Next PCount
    'Address of First Output with absolute rows
    Address = mOutSheet.Range("OutputResults").Resize(mIterations, 1).Address(True, False)
    Cell.Offset(3, 1).Formula = "=AGGREGATE(16,6," & Address & "," & Cell.Offset(3).Address(False, True) & ")"
    Cell.Offset(3, 1).Copy Cell.Offset(3, 1).Resize(41, mRiskOutputs.Count)
    ThickBorders Cell.Offset(3).Resize(41, 1)
    ThickBorders Cell.Offset(3, 1).Resize(41, mRiskOutputs.Count)
    mOutSheet.Range("SimStats").Offset(-3).Resize(3, 1 + mRiskOutputs.Count).Copy Cell
    Cell.Value = "Percentiles"
End Sub

Private Sub ProducePercentRank(Cell As Range)
    Dim Address As String
    
    mOutSheet.Names.Add Name:="PercentRank", RefersTo:=Cell
    Cell.Offset(3) = "Enter Value:"
    mOutSheet.Range("Percentiles").Cells(3, 2).Resize(1, mRiskOutputs.Count).Calculate
    Cell.Offset(3, 1).Resize(1, mRiskOutputs.Count).Value = _
        mOutSheet.Range("Percentiles").Cells(3, 2).Resize(1, mRiskOutputs.Count).Value '5th Percentile
    Cell.Offset(4) = "% rank:"
    Address = mOutSheet.Range("OutputResults").Resize(mIterations, 1).Address(True, False)
    Cell.Offset(4, 1).Formula = "=PERCENTRANK.INC(" & Address & "," & Cell.Offset(3, 1).Address(True, False) & ")"
    If mRiskOutputs.Count > 1 Then Cell.Offset(4, 1).Copy Cell.Offset(4, 2).Resize(1, mRiskOutputs.Count - 1)
    Cell.Offset(4, 1).Resize(1, mRiskOutputs.Count).NumberFormat = "0.00%"
    Cell.Offset(3).Resize(2, 1).HorizontalAlignment = xlRight
    ThickBorders Cell.Offset(3).Resize(2, 1)
    ThickBorders Cell.Offset(3, 1).Resize(2, mRiskOutputs.Count)
    mOutSheet.Range("SimStats").Offset(-3).Resize(3, 1 + mRiskOutputs.Count).Copy Cell
    Cell = "Percent Rank"
End Sub

Sub ProduceCorrelations(Cell As Range)
    Dim FirstInput As Range
    Dim OutAddress As String
    Dim I As Long
    
    Cell.Font.Bold = True
    mOutSheet.Names.Add Name:="Correlations", RefersTo:=Cell
    
    Cell.Offset(3).Resize(mRiskInputs.Count, 1).Value = WorksheetFunction.Transpose( _
        mOutSheet.Range("InputResults").Offset(-2).Resize(1, mRiskInputs.Count).Value)
        
    Set FirstInput = mOutSheet.Range("InputResults").Resize(mIterations, 1)
    OutAddress = mOutSheet.Range("OutputResults").Resize(mIterations, 1).Address(True, False)
    For I = 1 To mRiskInputs.Count
        Cell.Offset(2 + I, 1).Formula = _
        "=CORREL(" & FirstInput.Offset(0, I - 1).Address & "," & OutAddress & ")"
    Next I
    If mRiskOutputs.Count > 1 Then _
        Cell.Offset(3, 1).Resize(mRiskInputs.Count, 1).Copy _
        Cell.Offset(3, 2).Resize(mRiskInputs.Count, mRiskOutputs.Count - 1)
    ThickBorders Cell.Offset(3).Resize(mRiskInputs.Count, 1)
    ThickBorders Cell.Offset(3, 1).Resize(mRiskInputs.Count, mRiskOutputs.Count)
    mOutSheet.Range("SimStats").Offset(-3).Resize(3, 1 + mRiskOutputs.Count).Copy Cell
    Cell = "Correlations"
    Cell.Offset(2) = "Risk Inputs"
End Sub

Private Sub ProduceCumulativeDistributions()
    Dim I As Integer
    Dim R As Range
    Dim Percentiles As Range
    Dim NewChart As Chart
    Dim Source As Series
    
    Set Percentiles = mOutSheet.Range("Percentiles")
    Set R = mOutSheet.Range("OutCumDist").Resize(ChartHeight, ChartWidth)
    
    For I = 1 To mRiskOutputs.Count
        Set NewChart = mOutSheet.ChartObjects.Add(Left:=R.Left, Top:=R.Top, Width:=R.Width, Height:=R.Height).Chart
        
        With NewChart
            .ChartType = xlXYScatterLines
            .HasLegend = False
            Set Source = .SeriesCollection.NewSeries
            Source.Values = Percentiles.Columns(1)
            Source.XValues = Percentiles.Columns(I + 1)
            .Axes(xlValue).MaximumScale = 1
            If Percentiles.Cells(1, I + 1) > 0 Then
                .Axes(xlCategory).MinimumScale = WorksheetFunction.RoundDown(Percentiles.Cells(1, I + 1), 0)
            End If
            '.Axes(xlCategory).MaximumScale = WorksheetFunction.RoundUp(Percentiles.Cells(21, I + 1), 0)
            .SetElement (msoElementChartTitleAboveChart)
            .ChartTitle.text = "Cum. Distribution of " & mRiskOutputs(I).Name & _
                " (" & AddressWithSheet(mRiskOutputs(I).Cell) & ")"
            .ChartTitle.Font.Size = 14
            .ChartTitle.Font.Bold = False
        End With
        Set R = R.Offset(ChartHeight + 1)
    Next I
    mOutSheet.Range("A1").Select
End Sub

Private Sub ProduceHistograms()
    Dim SimOutput As Range
    Dim R As Range
    Dim ChartShape As Shape
    Dim NewChart As Chart
    Dim RiskOutput As ClsRiskOutput
    
    On Error Resume Next
    
    Set SimOutput = mOutSheet.Range("OutputResults").Resize(mIterations, 1)
    Set R = mOutSheet.Range("OutDist").Resize(ChartHeight, ChartWidth)
    
    For Each RiskOutput In mRiskOutputs
        ' Range needs to be selected
        'https://stackoverflow.com/questions/37912746/vba-why-xlhistogram-does-not-work
        SimOutput.Select
        Set ChartShape = ActiveSheet.Shapes.AddChart2(-1, xlHistogram, R.Left, R.Top, R.Width, R.Height)
        Set NewChart = ChartShape.Chart
        
        With NewChart
           .SetElement (msoElementChartTitleAboveChart)
           .ChartTitle.text = "Distribution of " & RiskOutput.Name & " (" & AddressWithSheet(RiskOutput.Cell) & ")"
        End With
        Set SimOutput = SimOutput.Offset(0, 1)
        Set R = R.Offset(ChartHeight + 1)
    Next RiskOutput
    mOutSheet.Range("A1").Select
End Sub

Sub ProduceTornadoDiagram()
    Dim Cell As Range
    Dim R As Range
    Dim SortedTable As Range
    Dim NewChart As Chart
    Dim Source As Series
    Dim I As Integer
    
    Set Cell = mOutSheet.Range("Tornado")
    Cell.Value = "Select Output:"
    Cell.Offset(1).Value = "Name:"
    
    Set SortedTable = Cell.Offset(0, ChartWidth + 1)
    
    mOutSheet.Range("Correlations").Resize(mRiskInputs.Count + 3, mRiskOutputs.Count + 1).Copy SortedTable.Offset(-1)
    SortedTable.Resize(mRiskInputs.Count + 1, mRiskOutputs.Count + 1).Columns.AutoFit
    SortedTable.Resize(mRiskInputs.Count + 1, mRiskOutputs.Count + 1).Columns.Hidden = True
        
    Cell.Offset(0, 1).Value = SortedTable.Offset(0, 1).Value 'First output cell address
    ThickBorders Cell.Offset(0, 1)
    With Cell.Offset(0, 1).Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="=" & SortedTable.Offset(0, 1).Resize(1, mRiskOutputs.Count).Address
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "Select Output Cell"
        .ShowInput = True
        .ShowError = True
    End With
    
    'Formulas
    Cell.Offset(1, 1).Formula = "=HLOOKUP(" & Cell.Offset(0, 1).Address & "," & _
        SortedTable.Offset(0, 1).Resize(2, mRiskOutputs.Count).Address & ",2,0)" 'Risk output name
    Cell.Offset(2).Formula = "=MATCH(" & Cell.Offset(0, 1).Address & "," & _
        SortedTable.Offset(0, 1).Resize(1, mRiskOutputs.Count).Address & ",0)" 'Column Index used in the formula below
    Cell.Offset(3).Formula = "=""Tornado Diagram for ""&" & Cell.Offset(1, 1).Address & "&"" (""&" & _
        Cell.Offset(0, 1).Address & "&"")""" ' Chart Title
    SortedTable.Offset(2).Resize(mRiskInputs.Count, mRiskOutputs.Count + 1).FormulaArray = _
        "=SortedTable(" & mOutSheet.Range("Correlations").Offset(3).Resize(mRiskInputs.Count, mRiskOutputs.Count + 1).Address _
        & "," & Cell.Offset(2).Address & "+1)" 'Sorted Table
    SortedTable.Offset(-1).Formula = "=""Correlations - Column ""&" & Cell.Offset(2).Address & "&"" sorted""" ' Table title
    
    Cell.Offset(2).NumberFormat = ";;;" 'Hide
    Cell.Resize(2, 2).Columns.AutoFit

    'Create Chart
    Set R = Cell.Offset(3).Resize(ChartHeight, ChartWidth)
    Set NewChart = mOutSheet.ChartObjects.Add(Left:=R.Left, Top:=R.Top, Width:=R.Width, Height:=R.Height).Chart
    
    mOutSheet.Names.Add Name:="TornadoY", RefersTo:="=OFFSET('" & mOutSheet.Name & "'!" & _
        SortedTable.Offset(2).Resize(mRiskInputs.Count, 1).Address & ",0,'" & _
        mOutSheet.Name & "'!" & Cell.Offset(2).Address & ")"
    
    With NewChart
        .ChartType = xlBarClustered
        .HasLegend = False
        .PlotVisibleOnly = False
        Set Source = .SeriesCollection.NewSeries
        Source.Values = "='" & mOutSheet.Name & "'!TornadoY"
        Source.XValues = SortedTable.Offset(2).Resize(mRiskInputs.Count, 1)
        .Axes(xlValue).MaximumScale = 1
        .Axes(xlValue).MinimumScale = -1
        .Axes(xlCategory).TickLabelPosition = xlLow
        '.ApplyDataLabels xlDataLabelsShowValue
        .SetElement (msoElementDataLabelInsideEnd)
        Source.DataLabels.NumberFormat = "0.00"
        Source.DataLabels.Position = xlLabelPositionInsideEnd
        .SetElement (msoElementChartTitleAboveChart)
        .ChartTitle.Caption = "='" & mOutSheet.Name & "'!" & Cell.Offset(3).Address
        .ChartTitle.Font.Size = 14
        .ChartTitle.Font.Bold = False
    End With
End Sub
