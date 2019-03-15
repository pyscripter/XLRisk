Attribute VB_Name = "RiskFunctions"
Option Explicit
Option Base 1

Public ProduceRandomSample As Boolean

' TODO: Use a Rnd replacement for higher quality random number generation

Public Function RiskFunctionList()
' Returns a list of risk functions
' Needs to be updated as new risk functions are added
    RiskFunctionList = Array("RiskUniform", "RiskNormal", "RiskTriang")
End Function

Public Function RiskUniform(Min As Double, Max As Double)
Attribute RiskUniform.VB_Description = "Generate random sample from a uniform destribution"
Attribute RiskUniform.VB_ProcData.VB_Invoke_Func = " \n20"
'  Random Sample from a Uniform distribution
    Application.Volatile
    
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


Public Function RiskNormal(Mean As Double, StDev As Double)
Attribute RiskNormal.VB_Description = "Generate random sample from a normal destribution"
Attribute RiskNormal.VB_ProcData.VB_Invoke_Func = " \n20"
'  Random Sample from a Normal distribution
    Application.Volatile
    
    If ProduceRandomSample Then
        RiskNormal = WorksheetFunction.Norm_Inv(Rnd(), Mean, StDev)
    Else
        RiskNormal = Mean
    End If
End Function

Function RiskTriang(Min As Double, Mode As Double, Max As Double)
'  Random Sample from a Normal distribution
'  See https://en.wikipedia.org/wiki/Triangular_distribution
    Dim LowerRange As Double, HigherRange As Double, TotalRange As Double, CumulativeProb As Double
    Application.Volatile
    
    'Error checking
    If (Mode < Min) Or (Max < Mode) Then
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
    Array("Minimum", "Mode", "Max"))
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
