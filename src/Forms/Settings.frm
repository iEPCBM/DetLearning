VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Settings 
   Caption         =   "Настройки"
   ClientHeight    =   4170
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   5760
   OleObjectBlob   =   "Settings.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private nColor As Long

Private Sub Confirm()
    setSetting stgSelColor, CStr(ActiveWorkbook.Colors(1))
    setSetting stgGenIterations, CStr(sliderIterations.value)
    setSetting stgMaxZeros, CStr(TextBoxZerosTo.value)
    setSetting stgMinZeros, CStr(TextBoxZerosFrom.value)
    setSetting stgPRNG, CInt(obtPRNGU.value)
    BookDetGame.Save
End Sub

Private Sub checkZerosCount(obj As Variant)
    If Not IsNumeric(obj.value) Then
        obj.value = ""
    ElseIf obj.value <= 0 Then
        obj.value = 1
    ElseIf obj.value > 4 Then
        obj.value = 4
    End If
End Sub

Private Sub btCancel_Click()
    Unload Me
End Sub

Private Sub btColorSelect_Click()
    Dim curColor As Long
    Dim curColorR As Integer
    Dim curColorG As Integer
    Dim curColorB As Integer
    Dim newColor As Long
    
    curColor = CLng(Common.getSetting(stgSelColor))
    curColorR = curColor Mod 2 ^ 8
    curColorG = (curColor \ 2 ^ 8) Mod 2 ^ 8
    curColorB = curColor \ 2 ^ 16
    Application.Dialogs(xlDialogEditColor).Show 1, curColorR, curColorG, curColorB
    lbColorDemo.BackColor = ActiveWorkbook.Colors(1)
End Sub

Private Sub btConfirm_Click()
    Confirm
End Sub

Private Sub btOK_Click()
    Confirm
    Unload Me
End Sub

Private Sub TextBoxZerosFrom_AfterUpdate()
    checkZerosCount TextBoxZerosFrom
End Sub

Private Sub TextBoxZerosTo_AfterUpdate()
    checkZerosCount TextBoxZerosTo
End Sub

Private Sub UserForm_Initialize()
    lbColorDemo.BackColor = Common.getSetting(stgSelColor)
    sliderIterations.value = Common.getSetting(stgGenIterations)
    TextBoxZerosFrom.value = Common.getSetting(stgMinZeros)
    TextBoxZerosTo.value = Common.getSetting(stgMaxZeros)
    If Common.getSetting(stgPRNG) Then
        obtPRNGU.value = True
    Else
        obtPRNGMT19937.value = True
    End If
    btOK.SetFocus
End Sub
