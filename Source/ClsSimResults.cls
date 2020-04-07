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
    mOutSheet.Activate
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
    
    '   Setup Simulation Statistics
    Set Curr = Curr.Offset(0, 2)
    Curr.Offset(-1, 0) = "Simulation Statistics"
    Curr.Offset(2, -1).Name = "SimStats"
    mOutSheet.Names.Add Name:="SimStats", RefersTo:=Curr.Offset(2, -1)
    For Each RiskOutput In mRiskOutputs
        Curr = AddressWithSheet(RiskOutput.Cell)
        Curr.Offset(1, 0) = RiskOutput.Name
        Set Curr = Curr.Offset(0, 1)
    Next RiskOutput
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
    
    '  Setup Output Cumulative Distributions
    Set Curr = Curr.Offset(0, 1)
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
            SubAddress:=mOutSheet.Range("OutCumDist").Offset(-2).Address, TextToDisplay:="Output Cumulative Distributions"
            
        If Val(Application.Version) >= 16 Then
            .Add Anchor:=Curr.Offset(6), Address:=vbNullString, _
                SubAddress:=mOutSheet.Range("OutDist").Offset(-2).Address, TextToDisplay:="Output Distributions"
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
    Dim Count As Integer
    Dim Perc As Integer
    Dim PCount As Integer
    
    Set FirstOutput = mOutSheet.Range("OutputResults").Resize(mIterations, 1)
    Set Cell = mOutSheet.Range("SimStats")
    
    Address = FirstOutput.Address(True, False)
    Count = mRiskOutputs.Count
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
    If Count > 1 Then Cell.Offset(0, 1).Resize(14, 1).Copy Cell.Offset(0, 2).Resize(14, Count - 1)
      
    'Percentiles
    Set Cell = Cell.Offset(15)
    Cell.Value = "Percentiles"
    Cell.Font.Bold = True
    Perc = 0
    For PCount = 1 To 41
        Cell.Offset(PCount) = Perc / 1000
        Cell.Offset(PCount).NumberFormat = "0.0%"
        Cell.Offset(PCount).HorizontalAlignment = xlRight
        Perc = Perc + 25
    Next PCount
    Cell.Offset(1, 1).Formula = "=AGGREGATE(16,6," & Address & "," & Cell.Offset(1).Address(False, True) & ")"
    mOutSheet.Names.Add Name:="Percentiles", RefersTo:=Cell.Offset(1).Resize(41, mRiskOutputs.Count + 1)
    Cell.Offset(1, 1).Copy Cell.Offset(1, 1).Resize(41, Count)
    
    ' Percent Rank
    Set Cell = Cell.Offset(43)
    Cell = "Percent Rank"
    Cell.Font.Bold = True
    mOutSheet.Names.Add Name:="PercentRank", RefersTo:=Cell
    Cell.Offset(1) = "Select Output Cell:"
    Cell.Offset(1, 1).Value = mOutSheet.Range("SimStats").Offset(-2, 1).Value
    ThickBorders Cell.Offset(1, 1)
    Cell.Offset(2) = "Output Name:"
    Address = mOutSheet.Range("SimStats").Offset(-2, 1).Resize(2, mRiskOutputs.Count).Address
    Cell.Offset(2, 1).Formula = "=HLOOKUP(" & Cell.Offset(1, 1).Address & "," & Address & ",2,0)"
    Cell.Offset(3) = "Value:"
    mOutSheet.Range("Percentiles").Cells(3, 2).Calculate
    Cell.Offset(3, 1).Value = mOutSheet.Range("Percentiles").Cells(3, 2).Value '5th Percentile
    ThickBorders Cell.Offset(3, 1)
    Cell.Offset(4) = "% rank:"
    Address = mOutSheet.Range("SimStats").Offset(-2, 1).Resize(1, mRiskOutputs.Count).Address
    Cell.Offset(4, 1) = "=PERCENTRANK.INC(OFFSET(" & FirstOutput.Address & ",0,MATCH(" & Cell.Offset(1, 1).Address _
        & "," & Address & ",0)-1)," & Cell.Offset(3, 1).Address & ")"
    Cell.Offset(4, 1).NumberFormat = "0.00%"
    Cell.Offset(1).Resize(4, 1).HorizontalAlignment = xlRight
    With Cell.Offset(1, 1).Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="=" & Address
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "Select Output Cell"
        .ShowInput = True
        .ShowError = True
    End With
    'AutoFit
    mOutSheet.Columns(Cell.Column).AutoFit
End Sub

Private Sub ProduceCumulativeDistributions()
    Dim I As Integer
    Dim R As Range
    Dim Percentiles As Range
    Dim ChartShape As ChartObject
    Dim NewChart As Chart
    Dim Source As Series
    
    Set Percentiles = mOutSheet.Range("Percentiles")
    Set R = mOutSheet.Range("OutCumDist").Resize(ChartHeight, ChartWidth)
    
    mOutSheet.Activate
    For I = 1 To mRiskOutputs.Count
        Set ChartShape = mOutSheet.ChartObjects.Add(Left:=R.Left, Top:=R.Top, Width:=R.Width, Height:=R.Height)
        Set NewChart = ChartShape.Chart
        
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
    
    mOutSheet.Activate
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

