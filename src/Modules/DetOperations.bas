Attribute VB_Name = "DetOperations"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright © 2020, Rishat D. Kagirov (iEPCBM)
' All rights reserved.
'
' Redistribution and use in source and binary forms, with or without modification,
' are permitted provided that the following conditions are met:
'
' 1. Redistributions of source code must retain the above copyright notice, this
' list of conditions and the following disclaimer.
'
' 2. Redistributions in binary form must reproduce the above copyright notice,
' this list of conditions and the following disclaimer in the documentation and/or
' other materials provided with the distribution.
'
' 3. Neither the name of the copyright holder nor the names of its contributors may
' be used to endorse or promote products derived from this software without specific
' prior written permission.
'
' THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
' ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
' WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED.
' IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT,
' INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT
' NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA,
' OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY,
' WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE)
' ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY
' OF SUCH DAMAGE.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub detSumToLine(ByVal isRow As Boolean, ByVal lFrom As Integer, ByVal lTo As Integer, ByVal factor As Double)
    alphabet = Array("A", "B", "C", "D", "E", "F", "G", "H", "I")
    If isRow Then
        For i = 0 To 3
            Range(alphabet(i + 2) & Val(lTo + 2)).value = CVar(Range(alphabet(i + 2) & Val(lTo + 2)).value) + CVar(Range(alphabet(i + 2) & Val(lFrom + 2)).value) * factor
        Next i
    Else
        For i = 0 To 3
            Range(alphabet(lTo + 1) & Val(i + 3)).value = CVar(Range(alphabet(lTo + 1) & Val(i + 3)).value) + CVar(Range(alphabet(lFrom + 1) & Val(i + 3)).value) * factor
        Next i
    End If
End Sub

Public Function getDeterminant() As Variant
    ' Extarct matrix
    alphabet = Array("A", "B", "C", "D", "E", "F", "G", "H", "I")
    Dim matrix(4, 4) As Variant
    Dim matrix_3(4, 3, 3) As Variant
    Dim dets_3(4) As Variant
    Dim factors(4) As Variant
    Dim retVal As Variant
    
    For i = 0 To 3
        For j = 0 To 3
            matrix(i, j) = Range(alphabet(j + 2) & Val(i + 3)).value
        Next j
    Next i
    ' Extract matrix_3
    For i = 0 To 3
        d = 0
        For j = 0 To 3
            factors(j) = matrix(j, 0)
            For m = 0 To 2
                If j = i Then
                    d = 1
                End If
                matrix_3(i, j, m) = matrix(j + d, m + 1)
            Next m
        Next j
    Next i
    ' Getting dets_3
    For i = 0 To 3
        Dim t1 As Variant
        Dim t2 As Variant
        t1 = matrix_3(i, 0, 0) * matrix_3(i, 1, 1) * matrix_3(i, 2, 2) + matrix_3(i, 0, 1) * matrix_3(i, 1, 2) * matrix_3(i, 2, 0) + matrix_3(i, 1, 0) * matrix_3(i, 2, 1) * matrix_3(i, 0, 2)
        t2 = matrix_3(i, 0, 2) * matrix_3(i, 1, 1) * matrix_3(i, 2, 0) + matrix_3(i, 1, 0) * matrix_3(i, 0, 1) * matrix_3(i, 2, 2) + matrix_3(i, 0, 0) * matrix_3(i, 2, 1) * matrix_3(i, 1, 2)
        dets_3(i) = t1 - t2
    Next i
    ' Getting Det
    For i = 0 To 3
        retVal = retVal + (-1) ^ i * factors(i) * dets_3(i)
    Next i
    getDeterminant = retVal
End Function

Public Function isOptimizedByTask() As Boolean
    'Extarct matrix
    alphabet = Array("A", "B", "C", "D", "E", "F", "G", "H", "I")
    Dim matrix(4, 4) As Variant
    Dim matrix_3(4, 3, 3) As Variant
    Dim dets_3(4) As Variant
    Dim factors(4) As Variant
    Dim retVal As Variant
    
    For i = 0 To 3
        For j = 0 To 3
            matrix(i, j) = Range(alphabet(j + 2) & Val(i + 3)).value
        Next j
    Next i
    isOptimizedByTask = False
    If getSetting(stgSolIsRows) Then
        For i = 0 To 3
            zeros = 0
            For j = 0 To 3
                If matrix(j, i) = 0 Then
                    zeros = zeros + 1
                End If
            Next j
            If zeros >= getSetting(stgSolZerosCount) Then
                isOptimizedByTask = True
            End If
        Next i
    Else
        For i = 0 To 3
            zeros = 0
            For j = 0 To 3
                If matrix(i, j) = 0 Then
                    zeros = zeros + 1
                End If
            Next j
            If zeros >= getSetting(stgSolZerosCount) Then
                isOptimizedByTask = True
            End If
        Next i
    End If
End Function
