Attribute VB_Name = "MacroStats"
'------------------------------------------------------------------------------
' MacroStats
' A collection of statistics-related helper functions for
' Excel's VBA macro language.
'
' By Iain Dunning, 2011
' http://www.iaindunning.com
' http://GITHUB
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
'   - This file
'   - The library itself, import this into your projects.
' MacroStatsTest.bas
'   - Mainly used to aid development, not necessary to include in your
'     projects.
'   - Used to aid development by ensuring that the functions work as
'     promised. Ff you encounter a problem and report it to me, I'll
'     add a test that should ensure the problem is never reintroduced.
'
'------------------------------------------------------------------------------

Public Function SampleDiscreteCDF( _
    ByRef CDF() As Variant, _
    Optional ByVal RaiseError As Boolean = False _
) As Integer
'------------------------------------------------------------------------------
' SampleDiscreteCDF
' Generate an random integer from a cumulative distribution function.
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
'   myCDF(4) = 0.4
'   myCDF(5) = 0.4
'   myCDF(6) = 0.1
'   Debug.Print SampleDiscreteCDF(myCDF)
'   ' 10% chance of returning a 3, 40% chance of a 4, etc.
'
'------------------------------------------------------------------------------

    Dim R As Single
    R = Rnd()
    
    Dim i As Long
    For i = LBound(CDF) To UBound(CDF)
        If R <= CDF(i) Then
            sampleDiscrete = i
            Exit Function
        End If
    Next i
    
    i = UBound(CDF)
End Function
