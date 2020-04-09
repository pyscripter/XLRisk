Attribute VB_Name = "RiskFunctions"
Option Explicit
Option Base 1

Public ProduceRandomSample As Boolean

' TODO: Use a Rnd replacement for higher quality random number generation

Public Function RiskFunctionList() As Variant
' Returns a list of risk functions
' Needs to be updated as new risk functions are added
    RiskFunctionList = Array("RiskBernoulli", "RiskBeta", "RiskBinomial", _
    "RiskCumul", "RiskDiscrete", "RiskDUniform", "RiskErlang", "RiskExpon", _
    "RiskGamma", "RiskLogNorm", "RiskNormal", "RiskTriang", "RiskPert", _
    "RiskUniform", "RiskWeibull")
End Function

Public Function RiskBinomial(N As Long, P As Double, Optional Corrmat As Long = 0) As Long
Attribute RiskBinomial.VB_Description = "Generate random sample from a Binomial destribution"
Attribute RiskBinomial.VB_ProcData.VB_Invoke_Func = " \n20"
    Dim RndValue As Double
    Application.Volatile (ProduceRandomSample)
    
    'Error checking
    If (N < 0) Or (P < 0) Or (P > 1) Then
        RiskBinomial = CVErr(xlErrValue)
        Exit Function
    End If
    
    If (N = 0) Or (P = 0) Then
        RiskBinomial = 0
    ElseIf ProduceRandomSample Then
        If gSimulation Is Nothing Then
            RndValue = Rnd()
        Else
            RndValue = gSimulation.GetRndSample(Application.Caller)
        End If
        RiskBinomial = WorksheetFunction.Binom_Inv(N, P, RndValue)
    Else
        RiskBinomial = N * P
    End If
End Function

Public Function RiskBernoulli(P As Double, Optional Corrmat As Long = 0) As Long
    Application.Volatile (ProduceRandomSample)
    
    'Error checking
    If (P < 0) Or (P > 1) Then
        RiskBernoulli = CVErr(xlErrValue)
        Exit Function
    End If
    
    RiskBernoulli = RiskBinomial(1, P)
End Function

Public Function RiskExpon(Mean As Double, Optional Corrmat As Long = 0) As Double
    Dim RndValue As Double
    Application.Volatile (ProduceRandomSample)
    
    'Error checking
    If Mean <= 0 Then
        RiskExpon = CVErr(xlErrValue)
        Exit Function
    End If
    
    If ProduceRandomSample Then
        If gSimulation Is Nothing Then
            RndValue = Rnd()
        Else
            RndValue = gSimulation.GetRndSample(Application.Caller)
        End If
        RiskExpon = -Mean * Log(1 - RndValue)
    Else
        RiskExpon = Mean
    End If
End Function

Function RiskErlang(kappa As Long, lambda As Double, Optional Corrmat As Long = 0) As Double
'   Random Sample from an Erlang distribution (same as Gamma with an integer alpha)

    Application.Volatile (ProduceRandomSample)
    
    'Error checking
    If (kappa <= 0) Or (lambda <= 0) Then
      RiskErlang = CVErr(xlErrValue)
      Exit Function
    End If
    
    RiskErlang = RiskGamma(CDbl(kappa), lambda)
End Function

Function RiskGamma(alpha As Double, beta As Double, Optional Corrmat As Long = 0) As Double
'   Random sample from a Gamma distribution
    
    Dim RndValue As Double
    Application.Volatile (ProduceRandomSample)
    
    'Error checking
    If (alpha <= 0) Or (beta <= 0) Then
      RiskGamma = CVErr(xlErrValue)
      Exit Function
    End If
    
    If ProduceRandomSample Then
        If gSimulation Is Nothing Then
            RndValue = Rnd()
        Else
            RndValue = gSimulation.GetRndSample(Application.Caller)
        End If
        RiskGamma = WorksheetFunction.Gamma_Inv(RndValue, alpha, beta)
    Else
        RiskGamma = alpha * beta
    End If
End Function

Public Function RiskUniform(Min As Double, Max As Double, Optional Corrmat As Long = 0) As Double
Attribute RiskUniform.VB_Description = "Generate random sample from a uniform destribution"
Attribute RiskUniform.VB_ProcData.VB_Invoke_Func = " \n20"
'   Random sample from a Uniform distribution
    
    Dim RndValue As Double
    Application.Volatile (ProduceRandomSample)
    
    'Error checking
    If Max < Min Then
      RiskUniform = CVErr(xlErrValue)
      Exit Function
    End If
    
    If ProduceRandomSample Then
        If gSimulation Is Nothing Then
            RndValue = Rnd()
        Else
            RndValue = gSimulation.GetRndSample(Application.Caller)
        End If
        RiskUniform = Min + RndValue * (Max - Min)
    Else
        RiskUniform = (Min + Max) / 2
    End If
End Function

Public Function RiskDiscrete(Values As Variant, Probabilities As Variant, Optional Corrmat As Long = 0) As Double
'   Generates samples for a discrete distribution specified by the array Values of accending values and the respective
'   probabilities
    
    Dim Count As Long
    Dim I As Long
    Dim RndValue As Double
    Dim Cum As Double
    
    Application.Volatile (ProduceRandomSample)
    
    Count = WorksheetFunction.Count(Values)

    'Error checking
    If (Count < 1) Or WorksheetFunction.Count(Probabilities) <> Count Then
      RiskDiscrete = CVErr(xlErrValue)
      Exit Function
    End If
    
    If Round(WorksheetFunction.Sum(Probabilities), 10) <> 1 Then
      RiskDiscrete = CVErr(xlErrValue)
      Exit Function
    End If

    For I = 2 To Count
        If (Probabilities(I) < 0) Or (Values(I) - Values(I - 1) < 0) Then
          RiskDiscrete = CVErr(xlErrValue)
          Exit Function
        End If
    Next I

    If ProduceRandomSample Then
        If gSimulation Is Nothing Then
            RndValue = Rnd()
        Else
            RndValue = gSimulation.GetRndSample(Application.Caller)
        End If
        Cum = 0
        For I = 1 To Count
            Cum = Cum + Probabilities(I)
            If RndValue <= Cum Then
                RiskDiscrete = Values(I)
                Exit Function
            End If
        Next I
    Else
        RiskDiscrete = WorksheetFunction.SumProduct(Values, Probabilities)
    End If
End Function

Public Function RiskDUniform(Values As Variant, Optional Corrmat As Long = 0) As Variant
'   Random sample from a Discrete Uniform distribution
'   Values can be a range or an array of values
    
    Dim Count As Integer
    Dim RndValue As Double
    Application.Volatile (ProduceRandomSample)
    
    Count = WorksheetFunction.Count(Values)

    If ProduceRandomSample Then
        If gSimulation Is Nothing Then
            RndValue = Rnd()
        Else
            RndValue = gSimulation.GetRndSample(Application.Caller)
        End If
        RiskDUniform = Values(Int(RndValue * Count) + 1)
    Else
        RiskDUniform = WorksheetFunction.Sum(Values) / Count
    End If
End Function

Public Function RiskNormal(Mean As Double, StDev As Double, Optional Corrmat As Long = 0) As Double
Attribute RiskNormal.VB_Description = "Generate random sample from a normal destribution"
Attribute RiskNormal.VB_ProcData.VB_Invoke_Func = " \n20"
'   Random sample from a Normal distribution
    
    Dim RndValue As Double
    Application.Volatile (ProduceRandomSample)
    
    If ProduceRandomSample Then
        If gSimulation Is Nothing Then
            RndValue = Rnd()
        Else
            RndValue = gSimulation.GetRndSample(Application.Caller)
        End If
        RiskNormal = WorksheetFunction.Norm_Inv(RndValue, Mean, StDev)
    Else
        RiskNormal = Mean
    End If
End Function

Public Function RiskLogNorm(Mean As Double, StDev As Double, Optional Corrmat As Long = 0) As Double
Attribute RiskLogNorm.VB_Description = "Generate random sample from a lognormal destribution"
Attribute RiskLogNorm.VB_ProcData.VB_Invoke_Func = " \n20"
'   Random sample from a Log Normal distribution
    
    Dim RndValue As Double
    Application.Volatile (ProduceRandomSample)
    
    If ProduceRandomSample Then
        If gSimulation Is Nothing Then
            RndValue = Rnd()
        Else
            RndValue = gSimulation.GetRndSample(Application.Caller)
        End If
        RiskLogNorm = WorksheetFunction.LogNorm_Inv(RndValue, Mean, StDev)
    Else
        RiskLogNorm = Exp(Mean + 0.5 * StDev ^ 2)
    End If
End Function

Function RiskTriang(Min As Double, Mode As Double, Max As Double, Optional Corrmat As Long = 0) As Double
Attribute RiskTriang.VB_Description = "Generate random sample from a triangular destribution"
Attribute RiskTriang.VB_ProcData.VB_Invoke_Func = " \n20"
'   Random sample from a Triangular distribution
'   See https://en.wikipedia.org/wiki/Triangular_distribution
    
    Dim LowerRange As Double
    Dim HigherRange As Double
    Dim TotalRange As Double
    Dim CumulativeProb As Double
    Application.Volatile (ProduceRandomSample)
    
    'Error checking
    If (Mode <= Min) Or (Max <= Mode) Then
      RiskTriang = CVErr(xlErrValue)
      Exit Function
    End If
    
    If ProduceRandomSample Then
        LowerRange = Mode - Min
        HigherRange = Max - Mode
        TotalRange = Max - Min
        If gSimulation Is Nothing Then
            CumulativeProb = Rnd()
        Else
            CumulativeProb = gSimulation.GetRndSample(Application.Caller)
        End If
        If CumulativeProb < (LowerRange / TotalRange) Then
            RiskTriang = Min + Sqr(CumulativeProb * LowerRange * TotalRange)
        Else
            RiskTriang = Max - Sqr((1 - CumulativeProb) * HigherRange * TotalRange)
        End If
    Else
        RiskTriang = (Min + Mode + Max) / 3
    End If
End Function

Function RiskBeta(alpha As Double, beta As Double, Optional A As Double = 0, Optional B As Double = 1, Optional Corrmat As Long = 0) As Double
Attribute RiskBeta.VB_Description = "Generate random sample from a beta destribution"
Attribute RiskBeta.VB_ProcData.VB_Invoke_Func = " \n20"
'   Random Sample from a Beta distribution
    
    Dim RndValue As Double
    Application.Volatile (ProduceRandomSample)
    
    'Error checking
    If (B <= A) Then
      RiskBeta = CVErr(xlErrValue)
      Exit Function
    End If
    
    If ProduceRandomSample Then
        If gSimulation Is Nothing Then
            RndValue = Rnd()
        Else
            RndValue = gSimulation.GetRndSample(Application.Caller)
        End If
        RiskBeta = WorksheetFunction.Beta_Inv(RndValue, alpha, beta, A, B)
    Else
        RiskBeta = A + (alpha / (alpha + beta)) * (B - A)
    End If
End Function

Function RiskPert(Min As Double, Mode As Double, Max As Double, Optional Corrmat As Long = 0) As Double
Attribute RiskPert.VB_Description = "Generate random sample from a PERT destribution"
Attribute RiskPert.VB_ProcData.VB_Invoke_Func = " \n20"
'   Random sample from a Pert distribution a special case of the Beta distribution
'   A smoother version of the triangular distribution
'   See https://www.coursera.org/lecture/excel-vba-for-creative-problem-solving-part-3-projects/the-beta-pert-distribution-GJVsK

    Dim alpha As Double
    Dim beta As Double
    
    Application.Volatile (ProduceRandomSample)
    
    'Error checking
    If (Mode <= Min) Or (Max <= Mode) Then
      RiskPert = CVErr(xlErrValue)
      Exit Function
    End If
    
    ' Reparameterize the RiskBeta distribution as per above link
    alpha = (4 * Mode + Max - 5 * Min) / (Max - Min)
    beta = (5 * Max - Min - 4 * Mode) / (Max - Min)
    RiskPert = RiskBeta(alpha, beta, Min, Max)
End Function

Public Function RiskCumul(MinValue As Double, MaxValue As Double, _
                          XValues As Variant, YValues As Variant, Optional Corrmat As Long = 0) As Double
Attribute RiskCumul.VB_Description = "Generate random sample from a cumulative destribution"
Attribute RiskCumul.VB_ProcData.VB_Invoke_Func = " \n20"
'   Random sample from a distribution specifed by data points of its cumulative function
'   Values can be a range or an array of values
    
    Dim I As Integer
    Dim Count As Integer
    Dim ParamError As Boolean
    Dim RndValue As Double
    Dim Slope As Double
    
    Application.Volatile (ProduceRandomSample)
    
    Count = WorksheetFunction.Count(XValues)

    'Error checking
    ParamError = False
    If MinValue > MaxValue Then
        ParamError = True
    ElseIf Count < 1 Then
      ParamError = True
    ElseIf WorksheetFunction.Count(XValues) <> Count Then
        ParamError = True
    ElseIf (XValues(1) <= MinValue) Or (XValues(Count) >= MaxValue) Then
        ParamError = True
    Else
        'Check that XValues and YValues are in strict increasing order
        For I = 2 To Count
            If XValues(I) <= XValues(I - 1) Then
                ParamError = True
                Exit For
            End If
        Next I
        If Not ParamError Then
            For I = 2 To Count
                If YValues(I) <= YValues(I - 1) Then
                    ParamError = True
                    Exit For
                End If
            Next I
        End If
    End If
    
    If ParamError Then
        RiskCumul = CVErr(xlErrValue)
        Exit Function
    End If
    
    If ProduceRandomSample Then
        If gSimulation Is Nothing Then
            RndValue = Rnd()
        Else
            RndValue = gSimulation.GetRndSample(Application.Caller)
        End If
        If RndValue <= YValues(1) Then
            RiskCumul = MinValue + (XValues(1) - MinValue) * RndValue / YValues(1)
        ElseIf RndValue > YValues(Count) Then
            RiskCumul = XValues(Count) + (MaxValue - XValues(Count)) * _
                (RndValue - YValues(Count)) / (1 - YValues(Count))
        Else
          For I = 2 To Count
            If RndValue <= YValues(I) Then
                RiskCumul = XValues(I - 1) + (XValues(I) - XValues(I - 1)) * _
                    (RndValue - YValues(I - 1)) / (YValues(I) - YValues(I - 1))
                Exit For
            End If
          Next I
        End If
    Else
        Slope = YValues(1) / (XValues(1) - MinValue)
        RiskCumul = 0.5 * Slope * (XValues(1) ^ 2 - MinValue ^ 2)
        For I = 2 To Count
            Slope = (YValues(I) - YValues(I - 1)) / (XValues(I) - XValues(I - 1))
            RiskCumul = RiskCumul + 0.5 * Slope * (XValues(I) ^ 2 - XValues(I - 1) ^ 2)
        Next I
        Slope = (1 - YValues(Count)) / (MaxValue - XValues(Count))
        RiskCumul = RiskCumul + 0.5 * Slope * (MaxValue ^ 2 - XValues(Count) ^ 2)
    End If
End Function

Function RiskWeibull(alpha As Double, beta As Double, Optional Corrmat As Long = 0) As Double
'   Random Sample from a Weibull distribution (same as Gamma with an integer alpha)
    Dim RndValue As Double
    Application.Volatile (ProduceRandomSample)
    
    'Error checking
    If (alpha <= 0) Or (beta <= 0) Then
      RiskWeibull = CVErr(xlErrValue)
      Exit Function
    End If
    
   If ProduceRandomSample Then
        If gSimulation Is Nothing Then
            RndValue = Rnd()
        Else
            RndValue = gSimulation.GetRndSample(Application.Caller)
        End If
        RiskWeibull = beta * (-Log(1 - RndValue)) ^ (1 / alpha)
    Else
        RiskWeibull = beta * WorksheetFunction.Gamma(1 + 1 / alpha)
    End If
End Function

Private Sub CreateFunctionDescription(FuncName As String, FuncDesc As String, ArgDesc As Variant)
'   Creates a description for a function and its arguments
'   They are used by the Excel function wizard
    On Error Resume Next
    Call Application.MacroOptions( _
      Macro:=FuncName, _
      Description:=FuncDesc, _
      Category:="XLRisk", _
      ArgumentDescriptions:=ArgDesc)
End Sub

Public Sub FunctionDescriptions()
    Const CorrMatDesc As String = "Optional. Placeholder for linking to a correlation matrix using the RiskCorrmat function)"
    ' Distribution functions
    Call CreateFunctionDescription("RiskBernoulli", "Generate random sample from a Bernoulli destribution", _
        Array("Probability of success", CorrMatDesc))
    Call CreateFunctionDescription("RiskBeta", "Generate random sample from a beta destribution", _
        Array("Shape parameter", "Shape parameter", "Optional minimum - 0 if omitted", "Optional maximum - 1 if omitted", CorrMatDesc))
    Call CreateFunctionDescription("RiskBinomial", "Generate random sample from a Binomial destribution", _
        Array("Number of Trials", "Probability of success", CorrMatDesc))
    Call CreateFunctionDescription("RiskCumul", "Generate random sample from a cumulative destribution", _
        Array("Minimum Value", "Maximum Value", "Range or array of X coordinates", _
        "Range or Array of Y coordinates (cumulative probabilities)", CorrMatDesc))
    Call CreateFunctionDescription("RiskDiscrete", "Generate random sample from a discrete probability destribution", _
        Array("Range or array of ascending values", "Range or array of probabilities", CorrMatDesc))
    Call CreateFunctionDescription("RiskDUniform", "Generate random sample from a uniform discrete destribution", _
        Array("Range or array of values", CorrMatDesc))
    Call CreateFunctionDescription("RiskGamma", "Generate random sample from a Gamma destribution", _
        Array("Shape parameter", "Scale parameter", CorrMatDesc))
    Call CreateFunctionDescription("RiskErlang", "Generate random sample from an Erlang destribution", _
        Array("Integer shape parameter", "Scale parameter", CorrMatDesc))
    Call CreateFunctionDescription("RiskExpon", "Generate random sample from an exponential destribution", _
        Array("Mean", CorrMatDesc))
    Call CreateFunctionDescription("RiskNormal", "Generate random sample from a normal destribution", _
        Array("Mean", "Standard Deviation", CorrMatDesc))
    Call CreateFunctionDescription("RiskLogNorm", "Generate random sample from a lognormal destribution", _
        Array("Mean of Ln(X)", "Standard Deviation of Ln(X)", CorrMatDesc))
    Call CreateFunctionDescription("RiskPert", "Generate random sample from a PERT destribution", _
        Array("Minimum value", "Mode", "Maximum value", CorrMatDesc))
    Call CreateFunctionDescription("RiskTriang", "Generate random sample from a triangular destribution", _
        Array("Minimum value", "Mode", "Maximum value", CorrMatDesc))
    Call CreateFunctionDescription("RiskUniform", "Generate random sample from a uniform destribution", _
        Array("Minimum value", "Maximum Value", CorrMatDesc))
    Call CreateFunctionDescription("RiskWeibull", "Generate random sample from a Weibull destribution", _
        Array("Shape parameter", "Scale parameter", CorrMatDesc))
    ' Other functions
    Call CreateFunctionDescription("RiskCorrmat", "When added as an optional last argument to a Risk function it creates a link to a correlation matrix", _
        Array("Range containing a lower triangular correlation matrix", "Position of this risk input function in the correlation matrix"))
    Call CreateFunctionDescription("RiskCorrectCorrmat", "Fixes an invalid correlation matrix and returns the corrected matrix", _
        Array("Range containing a lower triangular correlation matrix", "Optional tolerance with default value 1.0E-10"))
    Call CreateFunctionDescription("RiskIsValidCorrmat", "Checks whether a correlation matrix is valid.  Returns a Boolean value", _
        Array("Range containing a lower triangular correlation matrix", "Optional tolerance with default value 1.0E-10"))
    Call CreateFunctionDescription("RiskSCorrel", "Returns the Spearman's rank correlation of two arrays", _
        Array("First Range or array", "Second Range or array"))
End Sub

Public Function RiskCorrmat(CorrmatRng As Range, Position As Long) As Long
Attribute RiskCorrmat.VB_Description = "When added as an optional last argument to a Risk function it creates a link to a correlation matrix"
Attribute RiskCorrmat.VB_ProcData.VB_Invoke_Func = " \n20"
    If Not (gSimulation Is Nothing) And (TypeOf Application.Caller Is Range) Then
        If gSimulation.ProcessingCorrmatInfo And (gSimulation.ActiveWorkBook Is Application.Caller.Parent.Parent) Then
            gSimulation.ProcessCorrmatInfo Application.Caller, CorrmatRng, Position
        End If
    End If
    RiskCorrmat = 0
End Function

Public Function RndM(Optional ByVal Number As Long) As Double
' Wichman-Hill Pseudo Random Number Generator: an alternative for VB Rnd() function
' http://www.vbforums.com/showthread.php?499661-Wichmann-Hill-Pseudo-Random-Number-Generator-an-alternative-for-VB-Rnd%28%29-function
' See also https://www.random.org/analysis/#visual
' Not currently used
    Static lngX As Long
    Static lngY As Long
    Static lngZ As Long
    Static blnInit As Boolean
    Dim dblRnd As Double
    ' if initialized and no input number given
    If blnInit And Number = 0 Then
        ' lngX, lngY and lngZ will never be 0
        lngX = (171 * lngX) Mod 30269
        lngY = (172 * lngY) Mod 30307
        lngZ = (170 * lngZ) Mod 30323
    Else
        ' if no initialization, use Timer, otherwise ensure positive Number
        If Number = 0 Then Number = Timer * 60 Else Number = Number And &H7FFFFFFF
        lngX = (Number Mod 30269)
        lngY = (Number Mod 30307)
        lngZ = (Number Mod 30323)
        ' lngX, lngY and lngZ must be bigger than 0
        If lngX <= 0 Then lngX = 171
        If lngY <= 0 Then lngY = 172
        If lngZ <= 0 Then lngZ = 170
        ' mark initialization state
        blnInit = True
    End If
    ' generate a random number
    dblRnd = CDbl(lngX) / 30269# + CDbl(lngY) / 30307# + CDbl(lngZ) / 30323#
    ' return a value between 0 and 1
    RndM = dblRnd - Int(dblRnd)
End Function
