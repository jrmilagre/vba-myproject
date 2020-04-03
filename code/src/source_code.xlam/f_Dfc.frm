VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} f_Dfc 
   Caption         =   ":: Selecione a conta DFC ::"
   ClientHeight    =   4425
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11145
   OleObjectBlob   =   "f_Dfc.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "f_Dfc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private oControle   As control
Private n           As Integer
Private optButton   As Boolean

Private Sub UserForm_Initialize()

    If Not IsNull(oGlobal.PesquisaID) Then
        
        For n = 1 To 10
        
            Set oControle = Controls("opt" & Format(n, "00"))
            
            If CInt(oControle.Tag) = oGlobal.PesquisaID Then
            
                oControle.Value = True
                
                Exit For
                
            End If
                
        Next n
    
    End If

End Sub
Private Sub btnLimparSelecao_Click()

    For n = 1 To 10
    
        Set oControle = Controls("opt" & Format(n, "00"))
        
        oControle.Value = False
            
    Next n

End Sub
Private Sub btnOK_Click()
    
    optButton = False
    
    For n = 1 To 10
    
        Set oControle = Controls("opt" & Format(n, "00"))
        
        If oControle.Value = True Then
            
            optButton = True
            
            oGlobal.PesquisaID = CInt(oControle.Tag)
            
            Exit For
            
        End If
            
    Next n
    
    If optButton = False Then
    
        oGlobal.PesquisaID = Null
        
    End If

    Unload Me

End Sub
Private Sub btnCancelar_Click()

    Unload Me

End Sub

