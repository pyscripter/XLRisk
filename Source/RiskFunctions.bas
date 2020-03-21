Attribute VB_Name = "RiskFunctions"
Option Explicit
Option Base 1

Public ProduceRandomSample As Boolean

' TODO: Use a Rnd replacement for higher quality random number generation

Public Function RiskFunctionList()
' Returns a list of risk functions
' Needs to be updated as new risk functions are added
    RiskFunctionList = Array("RiskUniform", "RiskNormal", "RiskTriang", "RiskBeta", _
    "RiskPert", "RiskLogNorm", "RiskDUniform", "RiskCumul")
End Function

Public Function RiskUniform(Min As Double, Max As Double)
Attribute RiskUniform.VB_Description = "Generate random sample from a uniform destribution"
Attribute RiskUniform.VB_ProcData.VB_Invoke_Func = " \n20"
'  Random Sample from a Uniform distribution
    Application.Volatile (ProduceRandomSample)
    
    'Error checking
    If Max < Min Then
      RiskUniform = CVErr(xlErrValue)
      Exit Function
    End If
    
    If ProduceRandomSample Then
        RiskUniform = Min + Rnd() * (Max - Min)
    Else
        RiskUniform = (Min + Max) / 2
    End If
End Function

Public Function RiskDUniform(Values As Variant)
Attribute RiskDUniform.VB_Description = "Generate random sample from a uniform discrete destribution"
Attribute RiskDUniform.VB_ProcData.VB_Invoke_Func = " \n20"
'  Random Sample from a Discrete Uniform distribution
'  Values can be a range or an array of values
    Dim Count As Integer
    Application.Volatile (ProduceRandomSample)
    
    Count = WorksheetFunction.Count(Values)

    If ProduceRandomSample Then
        RiskDUniform = Values(Int(Rnd() * Count) + 1)
    Else
        RiskDUniform = WorksheetFunction.Sum(Values) / Count
    End If
End Function

Public Function RiskNormal(Mean As Double, StDev As Double)
Attribute RiskNormal.VB_Description = "Generate random sample from a normal destribution"
Attribute RiskNormal.VB_ProcData.VB_Invoke_Func = " \n20"
'  Random Sample from a Normal distribution
    Application.Volatile (ProduceRandomSample)
    
    If ProduceRandomSample Then
        RiskNormal = WorksheetFunction.Norm_Inv(Rnd(), Mean, StDev)
    Else
        RiskNormal = Mean
    End If
End Function

Public Function RiskLogNorm(Mean As Double, StDev As Double)
Attribute RiskLogNorm.VB_Description = "Generate random sample from a lognormal destribution"
Attribute RiskLogNorm.VB_ProcData.VB_Invoke_Func = " \n20"
'  Random Sample from a Log Normal distribution
    Application.Volatile (ProduceRandomSample)
    
    If ProduceRandomSample Then
        RiskLogNorm = WorksheetFunction.LogNorm_Inv(Rnd(), Mean, StDev)
    Else
        RiskLogNorm = Exp(Mean + 0.5 * StDev ^ 2)
    End If
End Function

Function RiskTriang(Min As Double, Mode As Double, Max As Double)
Attribute RiskTriang.VB_Description = "Generate random sample from a triangular destribution"
Attribute RiskTriang.VB_ProcData.VB_Invoke_Func = " \n20"
'  Random Sample from a Triangular distribution
'  See https://en.wikipedia.org/wiki/Triangular_distribution
    Dim LowerRange As Double, HigherRange As Double, TotalRange As Double, CumulativeProb As Double
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
        CumulativeProb = Rnd()
        If CumulativeProb < (LowerRange / TotalRange) Then
            RiskTriang = Min + Sqr(CumulativeProb * LowerRange * TotalRange)
        Else
            RiskTriang = Max - Sqr((1 - CumulativeProb) * HigherRange * TotalRange)
        End If
    Else
        RiskTriang = (Min + Mode + Max) / 3
    End If
End Function

Function RiskBeta(alpha As Double, beta As Double, Optional A As Double = 0, Optional B As Double = 1)
Attribute RiskBeta.VB_Description = "Generate random sample from a beta destribution"
Attribute RiskBeta.VB_ProcData.VB_Invoke_Func = " \n20"
'  Random Sample from a Beta distribution
    Application.Volatile (ProduceRandomSample)
    
    'Error checking
    If (B <= A) Then
      RiskBeta = CVErr(xlErrValue)
      Exit Function
    End If
    
    If ProduceRandomSample Then
        RiskBeta = WorksheetFunction.Beta_Inv(Rnd(), alpha, beta, A, B)
    Else
        RiskBeta = A + (alpha / (alpha + beta)) * (B - A)
    End If
End Function

Function RiskPert(Min As Double, Mode As Double, Max As Double)
Attribute RiskPert.VB_Description = "Generate random sample from a PERT destribution"
Attribute RiskPert.VB_ProcData.VB_Invoke_Func = " \n20"
'  Random Sample from a Pert distribution a special case of the Beta distribution
'  A smoother version of the triangular distribution
'  See https://www.coursera.org/lecture/excel-vba-for-creative-problem-solving-part-3-projects/the-beta-pert-distribution-GJVsK
    Dim alpha As Double, beta As Double
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
                          XValues As Variant, YValues As Variant)
Attribute RiskCumul.VB_Description = "Generate random sample from a cumulative destribution"
Attribute RiskCumul.VB_ProcData.VB_Invoke_Func = " \n20"
'  Random Sample from a Discrete Uniform distribution
'  Values can be a range or an array of values
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
        RndValue = Rnd()
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

Sub CreateFunctionDescription(FuncName, FuncDesc, ArgDesc)
'   Creates a description for a function and its arguments
'   They are used by the Excel function wizard
    On Error Resume Next
    Call Application.MacroOptions( _
      Macro:=FuncName, _
      Description:=FuncDesc, _
      Category:="XLRisk", _
      ArgumentDescriptions:=ArgDesc)
End Sub

Sub FunctionDescriptions()
  Call CreateFunctionDescription("RiskUniform", "Generate random sample from a uniform destribution", _
    Array("Minimum value", "Maximum Value"))
  Call CreateFunctionDescription("RiskNormal", "Generate random sample from a normal destribution", _
    Array("Mean", "Standard Deviation"))
  Call CreateFunctionDescription("RiskTriang", "Generate random sample from a triangular destribution", _
    Array("Minimum value", "Mode", "Maximum value"))
  Call CreateFunctionDescription("RiskBeta", "Generate random sample from a beta destribution", _
    Array("Shape parameter", "Shape parameter", "Optional minimum - 0 if omitted", "Optional maximum - 1 if omitted"))
  Call CreateFunctionDescription("RiskPert", "Generate random sample from a PERT destribution", _
    Array("Minimum value", "Mode", "Maximum value"))
  Call CreateFunctionDescription("RiskLogNorm", "Generate random sample from a lognormal destribution", _
    Array("Mean of Ln(X)", "Standard Deviation of Ln(X)"))
  Call CreateFunctionDescription("RiskDUniform", "Generate random sample from a uniform discrete destribution", _
    Array("Range or array of values"))
  Call CreateFunctionDescription("RiskCumul", "Generate random sample from a cumulative destribution", _
    Array("Minimum Value", "Maximum Value", "Range or array of X coordinates", _
    "Range or Array of Y coordinates (cumulative probabilities)"))
End Sub

Public Function RndM(Optional ByVal Number As Long) As Double
' Wichman-Hill Pseudo Random Number Generator: an alternative for VB Rnd() function
' http://www.vbforums.com/showthread.php?499661-Wichmann-Hill-Pseudo-Random-Number-Generator-an-alternative-for-VB-Rnd%28%29-function
' See also https://www.random.org/analysis/#visual
    Static lngX As Long, lngY As Long, lngZ As Long, blnInit As Boolean
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
        If lngX > 0 Then Else lngX = 171
        If lngY > 0 Then Else lngY = 172
        If lngZ > 0 Then Else lngZ = 170
        ' mark initialization state
        blnInit = True
    End If
    ' generate a random number
    dblRnd = CDbl(lngX) / 30269# + CDbl(lngY) / 30307# + CDbl(lngZ) / 30323#
    ' return a value between 0 and 1
    RndM = dblRnd - Int(dblRnd)
End Function
