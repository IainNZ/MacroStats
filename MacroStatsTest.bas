Attribute VB_Name = "MacroStatsTest"
Public Sub TestCDFs()
    Dim CDF() As Double, Counter() As Variant
    Dim result As Integer
    Dim i As Integer, j As Integer, k As Integer
    
    '//////////////////////////////////////////////////////////////////////////
    ' 1.    Test 1D
    ' 1.1   Random number tests
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
    Debug.Print "Test 1.1 - 1D - Random number tests - ";
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
    ' 2.    Test 2D
    ' 2.1   Random number tests
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
    Debug.Print "Test 2.1 - 2D - Random number tests - ";
    If Counter(3) > 0.39 And Counter(3) < 0.41 Then
        Debug.Print "Pass! ",
    Else
        Debug.Print "FAIL! ",
    End If
    Debug.Print Round(Counter(3), 3); Round(Counter(4), 3);
    Debug.Print Round(Counter(5), 3); Round(Counter(6), 3)

    '//////////////////////////////////////////////////////////////////////////
    ' 3.    Test 3D
    ' 3.1   Random number tests
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
    Debug.Print "Test 3.1 - 3D - Random number tests - ";
    If Counter(3) > 0.39 And Counter(3) < 0.41 Then
        Debug.Print "Pass! ",
    Else
        Debug.Print "FAIL! ",
    End If
    Debug.Print Round(Counter(3), 3); Round(Counter(4), 3);
    Debug.Print Round(Counter(5), 3); Round(Counter(6), 3)
End Sub

