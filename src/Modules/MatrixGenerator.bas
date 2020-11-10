Attribute VB_Name = "MatrixGenerator"
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

Public solIsRows As Boolean
Public solZeros As Integer

Public Function generateDet() As Long()
    generateDet = twistMatrix(getPrevMatrix, getSetting(stgGenIterations))
End Function

Private Function random() As Long
    Dim num As Long
    If getSetting(stgPRNG) Then
        num = Rnd
    Else
        num = MersenneTwisterVBAModule.genrand_int31
    End If
    random = num
End Function

Private Function randomRange(ByVal min As Integer, ByVal max As Integer) As Integer
    randomRange = random Mod (max - min + 1) + min
End Function

Private Function coinFlip() As Boolean
    Dim isTrue As Boolean
    If getSetting(stgPRNG) Then
        isTrue = Rnd Mod 2
    Else
        isTrue = MersenneTwisterVBAModule.getCoinFlip
    End If
    coinFlip = isTrue
End Function

Private Function twistMatrix(matrix() As Long, ByVal iterations As Integer) As Long()
    Dim factor As Integer
    Dim lineFrom As Long
    Dim lineTo As Long
    Dim isRow As Boolean
    Dim procMatrix() As Long
    procMatrix = matrix
    Randomize
    For i = 0 To (iterations - 1)
        isRow = coinFlip
        factor = randomRange(1, 2)
        If coinFlip Then
            factor = -factor
        End If
        lineFrom = randomRange(0, 3)
        lineTo = randomRange(0, 3)
        If lineTo = lineFrom Then
            lineTo = (lineTo + 1) Mod 4
        End If
        
        procMatrix = sumLine(procMatrix, isRow, lineFrom, lineTo, factor)
    Next i
    twistMatrix = procMatrix
End Function

Private Function sumLine(matrix() As Long, ByVal isRow As Boolean, ByVal lFrom As Integer, ByVal lTo As Integer, ByVal factor As Integer) As Long()
    If isRow Then
        For i = 0 To 3
            matrix(i, lTo) = factor * matrix(i, lFrom) + matrix(i, lTo)
        Next i
    Else
        For i = 0 To 3
            matrix(lTo, i) = factor * matrix(lFrom, i) + matrix(lTo, i)
        Next i
    End If
    sumLine = matrix
End Function

Private Function getPrevMatrix() As Long()
    ' Cols or Rows
    Dim position As Integer
    Dim position2 As Integer
    Dim factor As Integer
    Dim maxZeros As Integer
    Dim minZeros As Integer
    maxZeros = CInt(getSetting(stgMaxZeros))
    minZeros = CInt(getSetting(stgMinZeros))
    
    Dim matrix(4, 4) As Long
    Dim cLine(4) As Long
    
    setSetting stgSolIsRows, CInt(coinFlip)
    setSetting stgSolZerosCount, randomRange(minZeros, maxZeros)
    position = randomRange(0, 3)
    position2 = randomRange(0, 3)
    If position = position2 Then
        position = (position + 1) Mod 4
    End If
    For i = 0 To 3
        For j = 0 To 3
            matrix(i, j) = randomRange(1, 9)
        Next j
    Next i
    
    For i = 0 To (getSetting(stgSolZerosCount) - 1)
        If Not getSetting(stgSolIsRows) Then
            matrix(position, i) = 0
        Else
            matrix(i, position) = 0
        End If
    Next i
    
    factor = randomRange(1, 2)
    If coinFlip Then
        factor = -factor
    End If
    getPrevMatrix = sumLine(matrix, getSetting(stgSolIsRows), position2, position, factor)
End Function

