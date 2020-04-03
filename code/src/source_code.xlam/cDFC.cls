VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cDfc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Classe

Private Type TClass
    ID              As Integer
    Grupo           As String
    Subtotal        As Boolean
    Sequencia       As Integer
End Type

Private This As TClass

Private Const sTabela As String = "tbl_dfc"

' Propriedades

Public Property Get ID() As Integer: ID = This.ID: End Property
Public Property Let ID(ByVal Value As Integer): This.ID = Value: End Property

Public Property Get Grupo() As String: Grupo = This.Grupo: End Property
Public Property Let Grupo(ByVal Value As String): This.Grupo = Value: End Property

Public Property Get Subtotal() As Boolean: Subtotal = This.Subtotal: End Property
Public Property Let Subtotal(ByVal Value As Boolean): This.Subtotal = Value: End Property

Public Property Get Sequencia() As Integer: Sequencia = This.Sequencia: End Property
Public Property Let Sequencia(ByVal Value As Integer): This.Sequencia = Value: End Property

' Métodos

Public Sub CRUD(Acao As Integer, Optional ID As Long)

    On Error GoTo Erro
    
    If ID > 0 Then
    
        sSQL = "SELECT * FROM " & sTabela & " WHERE id = " & ID & " AND subtotal = FALSE"
        
    End If

    If Acao = eCrud.Read Then ' READ
    
        Set rst = New ADODB.Recordset
        
        With rst
        
            .Open sSQL, cnn, adOpenStatic, adLockReadOnly
        
            Me.ID = .Fields("id").Value
            Me.Grupo = .Fields("grupo").Value
            Me.Subtotal = .Fields("subtotal").Value
            Me.Sequencia = .Fields("sequencia").Value
            
        End With
            
        Set rst = Nothing
    
    End If
    
    Exit Sub
    
Erro:

    MsgBox err.Description, vbCritical, "Erro!"

End Sub
