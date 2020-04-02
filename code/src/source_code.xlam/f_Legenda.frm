VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} f_Legenda 
   Caption         =   ":: Legenda ::"
   ClientHeight    =   1260
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "f_Legenda.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "f_Legenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()

    Dim iTop As Integer
    Dim iItens As Integer
    Dim iLeft As Integer
    Dim i As Integer
    
    Dim obj As Variant
    Dim c As New Collection
    Dim s() As String
    Dim oCor    As control
    Dim oTexto  As control
    
    iTop = 10
    iItens = oLegenda.Count
    iLeft = 10
    
    Me.Height = (iItens * 10) + iTop + 50
    Me.Width = 200
    
    For Each obj In oLegenda
    
        s() = Split(obj, ";")
        
        Set oCor = Me.Controls.Add("Forms.Label.1", "Cor", True)
        
        With oCor
            .Width = 10
            .Height = 10
            .BorderStyle = fmBorderStyleSingle
            .BorderColor = &H80000006
            .BackColor = s(2)
            .Top = iTop
            .Left = iLeft
        End With
        
        Set oTexto = Me.Controls.Add("Forms.Label.1", "Rótulo", True)
        
        With oTexto
            .Width = 50
            .Height = 18
            .Top = iTop
            .Left = iLeft + 15
            .Caption = s(0) & " - " & s(1)
        End With
        
        iTop = iTop + 9.75
        
    Next

End Sub
