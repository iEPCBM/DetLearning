VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formAbout 
   Caption         =   "О программе"
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   OleObjectBlob   =   "formAbout.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formAbout"
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

Private Sub cmdOK_Click()
    Me.Hide
End Sub

Private Sub cmdViewLicense_Click()
    formLicense.Show
End Sub

Private Sub UserForm_Activate()
    lTitle.Caption = Common.getSetting(stgAboutTitle)
    lAutor.Caption = Common.getSetting(stgAboutAutor)
    lVersion.Caption = Common.getSetting(stgAboutVersion)
    lVersionCode.Caption = Common.getSetting(stgAboutVersionCode)
End Sub
