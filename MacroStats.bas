Attribute VB_Name = "MacroStats"
'------------------------------------------------------------------------------
' MacroStats
' A collection of statistics-related functions for Excel's VBA macro language.
'
' By Iain Dunning, 2011
' http://www.iaindunning.com
' https://github.com/IainNZ/MacroStats
'
'------------------------------------------------------------------------------
'
' Licensing: MIT License
'
' Copyright (c) 2011 Iain Dunning, http://www.iaindunning.com
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in
' all copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
' THE SOFTWARE.
'
'------------------------------------------------------------------------------
'
' Project layout:
' MacroStats.bas
'   - The library itself, IMPORT THIS INTO YOUR PROJECTS!
' MacroStats.xlsm
'   - The workbook used to develop the libary.
'   - Don't need to add this to you to your project.
' MacroStatsTest.bas
'   - Mainly used to aid development, not necessary to include in your
'     projects.
'   - Used to aid development by ensuring that the functions work as
'     promised. Ff you encounter a problem and report it to me, I'll
'     add a test that should ensure the problem is never reintroduced.
'
'------------------------------------------------------------------------------
Option Explicit

Public Function SampleDiscreteCDF( _
    ByRef CDF As Variant, _
    Optional ByVal RaiseError As Boolean = False _
) As Long
'------------------------------------------------------------------------------
' SampleDiscreteCDF
' Generate a random integer from a cumulative distribution function.
' For more information about CDFs:
' http://en.wikipedia.org/wiki/Cumulative_distribution_function
'
' Input:
'   CDF() as Variant
'       - A one-dimensional array of numbers that describe a CDF.
'   [Optional] RaiseError as Boolean
'       - Default = False
'       - If there is a problem with the CDF, and this value is true, then
'         an error will be raised so you can address it
'
' Return:
'   A random integer, chosen using the CDF.
'   The integer is taken from the array indices.
'
' Example Usage:
'   Dim myCDF(3 to 6) as Double
'   myCDF(3) = 0.1
'   myCDF(4) = 0.5
'   myCDF(5) = 0.9
'   myCDF(6) = 1.0
'   Debug.Print SampleDiscreteCDF(myCDF)
'   ' 10% chance of returning a 3, 40% chance of a 4, etc.
'
'------------------------------------------------------------------------------

    Dim R As Single
    R = Rnd()
    
    Dim i As Long
    For i = LBound(CDF) To UBound(CDF)
        If R <= CDF(i) Then
            SampleDiscreteCDF = i
            Exit Function
        End If
    Next i
    
    ' If we got to this point, it must not have been a valid CDF because
    ' the last element in the array was not a 1.
    If RaiseError Then
        Err.Raise vbObjectError + 1, "MacroStats.SampleDiscreteCDF", _
                  "Unable to sample from CDF - check last element is 1."
    End If
End Function


Public Function SampleDiscreteCDFon2D( _
    ByRef CDF As Variant, _
    ByVal FirstIndex As Long, _
    Optional ByVal RaiseError As Boolean = False _
) As Long
'------------------------------------------------------------------------------
' SampleDiscreteCDFon2D
' Generate a random integer from a cumulative distribution function.
' Use this when you have a 2D array, where the CDF is in the second dimension.
' For more information about CDFs:
' http://en.wikipedia.org/wiki/Cumulative_distribution_function
'
' Input:
'   CDF() as Variant
'       - A one-dimensional array of numbers that describe a CDF.
'   FirstIndex as Long
'       - The index in the first dimension of the array.
'   [Optional] RaiseError as Boolean
'       - Default = False
'       - If there is a problem with the CDF, and this value is true, then
'         an error will be raised so you can address it
'
' Return:
'   A random integer, chosen using the CDF.
'   The integer is taken from the array indices of the second dimension.
'
' Example Usage:
'   Dim myCDF(1 to 2, 3 to 6) as Double
'   myCDF(1,3) = 0.1: myCDF(2,3) = 0.4
'   myCDF(1,4) = 0.5: myCDF(2,4) = 0.5
'   myCDF(1,5) = 0.9: myCDF(2,5) = 0.9
'   myCDF(1,6) = 1.0: myCDF(2,6) = 1.0
'   Debug.Print SampleDiscreteCDFon2D(myCDF, 2)
'   ' 40% chance of returning a 3, 10% chance of a 4, etc.
'
'------------------------------------------------------------------------------

    Dim R As Single
    R = Rnd()
    
    Dim i As Long
    For i = LBound(CDF, 2) To UBound(CDF, 2)
        If R <= CDF(FirstIndex, i) Then
            SampleDiscreteCDFon2D = i
            Exit Function
        End If
    Next i
    
    ' If we got to this point, it must not have been a valid CDF because
    ' the last element in the array was not a 1.
    If RaiseError Then
        Err.Raise vbObjectError + 1, "MacroStats.SampleDiscreteCDFon2D", _
                  "Unable to sample from CDF - check last element is 1."
    End If
End Function


Public Function SampleDiscreteCDFon3D( _
    ByRef CDF As Variant, _
    ByVal FirstIndex As Long, _
    ByVal SecondIndex As Long, _
    Optional ByVal RaiseError As Boolean = False _
) As Long
'------------------------------------------------------------------------------
' SampleDiscreteCDFon3D
' Generate a random integer from a cumulative distribution function.
' Use this when you have a 3D array, where the CDF is in the third dimension.
' For more information about CDFs:
' http://en.wikipedia.org/wiki/Cumulative_distribution_function
'
' Input:
'   CDF() as Variant
'       - A one-dimensional array of numbers that describe a CDF.
'   FirstIndex as Long
'       - The index in the first dimension of the array.
'   SecondIndex as Long
'       - The index in the second dimension of the array.
'   [Optional] RaiseError as Boolean
'       - Default = False
'       - If there is a problem with the CDF, and this value is true, then
'         an error will be raised so you can address it
'
' Return:
'   A random integer, chosen using the CDF.
'   The integer is taken from the array indices of the third dimension.
'
' Example Usage:
'   Dim myCDF(9 to 9, 1 to 2, 3 to 6) as Double
'   myCDF(9,1,3) = 0.1: myCDF(9,2,3) = 0.4
'   myCDF(9,1,4) = 0.5: myCDF(9,2,4) = 0.5
'   myCDF(9,1,5) = 0.9: myCDF(9,2,5) = 0.9
'   myCDF(9,1,6) = 1.0: myCDF(9,2,6) = 1.0
'   Debug.Print SampleDiscreteCDFon2D(myCDF, 9, 2)
'   ' 40% chance of returning a 3, 10% chance of a 4, etc.
'
'------------------------------------------------------------------------------

    Dim R As Single
    R = Rnd()
    
    Dim i As Long
    For i = LBound(CDF, 3) To UBound(CDF, 3)
        If R <= CDF(FirstIndex, SecondIndex, i) Then
            SampleDiscreteCDFon3D = i
            Exit Function
        End If
    Next i
    
    ' If we got to this point, it must not have been a valid CDF because
    ' the last element in the array was not a 1.
    If RaiseError Then
        Err.Raise vbObjectError + 1, "MacroStats.SampleDiscreteCDFon3D", _
                  "Unable to sample from CDF - check last element is 1."
    End If
End Function


Public Function FlipCoin( _
    ByVal Probability As Double _
) As Boolean
'------------------------------------------------------------------------------
' FlipCoin
' Returns a true with the provided probability.
'
' Input:
'   Probability as Double
'       - The probability we are using to generate the True/False
'
' Return:
'   True or False, with probability provided.
'
' Example Usage:
'   Debug.Print FlipCoin(0.5)
'   ' Like flipping a coin, 50/50 of being True
'   Debug.Print FlipCoin(0.6) ' 60/40..., etc.
'
'------------------------------------------------------------------------------
    
    FlipCoin = (Rnd() < Probability)
    
End Function
