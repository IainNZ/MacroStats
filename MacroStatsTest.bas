Attribute VB_Name = "MacroStatsTest"
Option Explicit

Public Sub TestProbability()
    Dim CDF() As Double, Counter() As Variant
    Dim result As Variant
    Dim i As Integer, j As Integer, k As Integer
    
    '//////////////////////////////////////////////////////////////////////////
    ' 1.    Test CDF 1D
    ' 1.1   Random number test
    ReDim CDF(10 To 19) As Double
    ReDim Counter(10 To 19) As Variant
    CDF(10) = 0.1
    For i = 11 To 19
        CDF(i) = 0.1 + CDF(i - 1)
    Next i
    For i = 1 To 10000
        result = MacroStats.SampleDiscreteCDF(CDF, True)
        Counter(result) = Counter(result) + 1 / 10000
    Next i
    Debug.Print "Test 1.1 - 1D - Random number test - ";
    If Counter(10) > 0.09 And Counter(10) < 0.11 Then
        Debug.Print "Pass! ",
    Else
        Debug.Print "FAIL! ",
    End If
    Debug.Print Round(Counter(10), 3); Round(Counter(11), 3);
    Debug.Print Round(Counter(12), 3); Round(Counter(13), 3)
    
    ' 1.2   Test silent fail
    ReDim CDF(1 To 3) As Double
    CDF(1) = 0.33
    CDF(2) = 0.33
    CDF(3) = 0.33
    result = MacroStats.SampleDiscreteCDF(CDF, False)
    result = MacroStats.SampleDiscreteCDF(CDF, False)
    result = MacroStats.SampleDiscreteCDF(CDF, False)
    result = MacroStats.SampleDiscreteCDF(CDF, False)
    result = MacroStats.SampleDiscreteCDF(CDF, False)
    Debug.Print "Test 1.2 - 1D - Silent fail test - Pass!"
    
    '//////////////////////////////////////////////////////////////////////////
    ' 2.    Test CDF 2D
    ' 2.1   Random number test
    ReDim CDF(1 To 2, 3 To 6) As Double
    ReDim Counter(3 To 6) As Variant
    CDF(1, 3) = 0.1: CDF(2, 3) = 0.4
    CDF(1, 4) = 0.5: CDF(2, 4) = 0.5
    CDF(1, 5) = 0.9: CDF(2, 5) = 0.9
    CDF(1, 6) = 1#: CDF(2, 6) = 1#
    For i = 1 To 10000
        result = MacroStats.SampleDiscreteCDFon2D(CDF, 2, True)
        Counter(result) = Counter(result) + 1 / 10000
    Next i
    Debug.Print "Test 2.1 - 2D - Random number test - ";
    If Counter(3) > 0.39 And Counter(3) < 0.41 Then
        Debug.Print "Pass! ",
    Else
        Debug.Print "FAIL! ",
    End If
    Debug.Print Round(Counter(3), 3); Round(Counter(4), 3);
    Debug.Print Round(Counter(5), 3); Round(Counter(6), 3)

    '//////////////////////////////////////////////////////////////////////////
    ' 3.    Test CDF 3D
    ' 3.1   Random number test
    ReDim CDF(9, 1 To 2, 3 To 6) As Double
    ReDim Counter(3 To 6) As Variant
    CDF(9, 1, 3) = 0.1: CDF(9, 2, 3) = 0.4
    CDF(9, 1, 4) = 0.5: CDF(9, 2, 4) = 0.5
    CDF(9, 1, 5) = 0.9: CDF(9, 2, 5) = 0.9
    CDF(9, 1, 6) = 1#: CDF(9, 2, 6) = 1#
    For i = 1 To 10000
        result = MacroStats.SampleDiscreteCDFon3D(CDF, 9, 2, True)
        Counter(result) = Counter(result) + 1 / 10000
    Next i
    Debug.Print "Test 3.1 - 3D - Random number test - ";
    If Counter(3) > 0.39 And Counter(3) < 0.41 Then
        Debug.Print "Pass! ",
    Else
        Debug.Print "FAIL! ",
    End If
    Debug.Print Round(Counter(3), 3); Round(Counter(4), 3);
    Debug.Print Round(Counter(5), 3); Round(Counter(6), 3)
    
    '//////////////////////////////////////////////////////////////////////////
    ' 4.    Test FlipCoin
    ' 4.1   Random number test
    ReDim Counter(1 To 2) As Variant
    For i = 1 To 10000
        result = MacroStats.FlipCoin(0.4)
        If result Then Counter(1) = Counter(1) + 1 / 10000
        If Not result Then Counter(2) = Counter(2) + 1 / 10000
    Next i
    Debug.Print "Test 4.1 - FlipCoin - Random number test - ";
    If Counter(1) > 0.39 And Counter(1) < 0.41 Then
        Debug.Print "Pass! ",
    Else
        Debug.Print "FAIL! ",
    End If
    Debug.Print Round(Counter(1), 3); Round(Counter(2), 3)
    
End Sub

Public Sub TestDistributionFitting()
    Dim i As Integer

    '//////////////////////////////////////////////////////////////////////////
    ' 1.    Test Normal fitting
    ' 1.1   Fit to data
    Dim normalData(1 To 1000) As Double
    For i = 1 To 1000
        normalData(i) = MacroStats.RandomFromNormal(3, 2)
    Next i
    Dim normMean As Double, normStdDev As Double
    Debug.Print "Test 1.1 - Normal - Generate + FitToData test - ";
    If MacroStats.FitNormalDistributionToData(normalData, normMean, normStdDev) Then
        Debug.Print "Fitted! Mean [3] = "; Round(normMean, 2); ", SD [2] = "; Round(normStdDev, 2)
    Else
        Debug.Print "Fitting error!"
    End If
    ' 1.2   Fit to percentiles
    Debug.Print "Test 1.2 - Normal - Fit to percentile test"
    FitNormalDistributionToPercentiles 0, 0.5, 1, 0.841, normMean, normStdDev
    Debug.Print " 00.0 0.500 01.0 0.841 -> ", Round(normMean, 1), Round(normStdDev, 1)
    FitNormalDistributionToPercentiles 50, 0.5, 84.1, 0.841, normMean, normStdDev
    Debug.Print " 50.0 0.500 84.1 0.841 -> ", Round(normMean, 1), Round(normStdDev, 1)
    FitNormalDistributionToPercentiles 60, 0.5, 80, 0.841, normMean, normStdDev
    Debug.Print " 60.0 0.500 80.0 0.841 -> ", Round(normMean, 1), Round(normStdDev, 1)
    FitNormalDistributionToPercentiles 60, 0.5, 80, 0.841, normMean, normStdDev
    Debug.Print " 60.0 0.500 80.0 0.841 -> ", Round(normMean, 1), Round(normStdDev, 1)
    
    '//////////////////////////////////////////////////////////////////////////
    ' 2.    Test Gamma fitting
    ' 2.1   Fit to data
    Dim gammaData(1 To 10000) As Double
    For i = 1 To 10000
        gammaData(i) = WorksheetFunction.GammaInv(Rnd(), 10, 22)
    Next i
    Dim gammaShape As Double, gammaScale As Double
    Debug.Print "Test 2.1 - Gamma - FitToData test - ";
    If MacroStats.FitGammaDistributionToData(gammaData, gammaShape, gammaScale) Then
        Debug.Print "Fitted! Shape [10] = "; Round(gammaShape, 2); ", Scale [22] = "; Round(gammaScale, 2)
    Else
        Debug.Print "Fitting error!"
    End If
    ' 2.2   Fit to percentiles
    Debug.Print "Test 2.2 - Gamma - Fit to percentile test"
    FitGammaDistributionToPercentiles 5000, 0.5, 6500, 0.841, gammaShape, gammaScale
    Debug.Print "5000, 0.5, 6500, 0.841 -> [13.5, 378.9]", Round(gammaShape, 2), Round(gammaScale, 2)
    
End Sub
