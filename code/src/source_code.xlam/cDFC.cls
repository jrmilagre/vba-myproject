VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cDFC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Classe

Private Type TClass
    ID              As Long
    Grupo           As String
    Subtotal        As Boolean
    Sequencia       As Integer

    Filtro          As String
    Ordem           As String
End Type

Private This As TClass

Private Const sTabela As String = "tbl_dfc"

' Propriedades

Public Property Get ID() As Long: ID = This.ID: End Property
Public Property Let ID(ByVal Value As Long): This.ID = CLng(Value): End Property

Public Property Get Grupo() As String: Grupo = This.Grupo: End Property
Public Property Let Grupo(ByVal Value As String): This.Grupo = Value: End Property

Public Property Get Subtotal() As Boolean: Subtotal = This.Subtotal: End Property
Public Property Let Subtotal(ByVal Value As Boolean): This.Subtotal = Value: End Property

Public Property Get Sequencia() As Integer: Sequencia = This.Sequencia: End Property
Public Property Let Sequencia(ByVal Value As Integer): This.Sequencia = Value: End Property

Public Property Get Filtro() As String: Filtro = This.Filtro: End Property
Public Property Let Filtro(ByVal Value As String): This.Filtro = Value: End Property

Public Property Get Ordem() As String: Ordem = This.Ordem: End Property
Public Property Let Ordem(ByVal Value As String): This.Ordem = Value: End Property
' Métodos

Public Sub CRUD(Acao As Integer, Optional ID As Long)

    On Error GoTo Erro
    
    If ID > 0 Then
    
        sSQL = "SELECT * FROM " & sTabela & " WHERE id = " & ID
        
    End If

    Set rst = New ADODB.Recordset
    
    If Acao = eCrud.Read Then ' READ
        
        With rst
        
            .Open sSQL, cnn, adOpenStatic, adLockReadOnly
        
            Me.ID = .Fields("id").Value
            Me.Grupo = .Fields("grupo").Value
            Me.Subtotal = .Fields("subtotal").Value
            Me.Sequencia = .Fields("sequencia").Value
            
        End With
        
    ElseIf Acao = eCrud.Create Or Acao = eCrud.Update Then
        
        With rst
        
            If Acao = eCrud.Create Then
            
                .Open sTabela, cnn, adOpenDynamic, adLockOptimistic, adCmdTable
                
                .AddNew
                
                Me.ID = .Fields("id").Value
                
            ElseIf Acao = eCrud.Update Then
            
                .Open sSQL, cnn, adOpenDynamic, adLockOptimistic
                
            End If
            
            .Fields("grupo").Value = Me.Grupo
            .Fields("subtotal").Value = Me.Subtotal
            .Fields("sequencia").Value = Me.Sequencia

            .Update
            
        End With
        
    ElseIf Acao = eCrud.Delete Then
    
        With rst
        
            .Open sSQL, cnn, adOpenDynamic, adLockOptimistic
            
            .Delete
            
        End With
        
    End If
    
    Set rst = Nothing
    
    Exit Sub
    
Erro:

    MsgBox err.Description, vbCritical, "Erro!"

End Sub
Public Function Todos(Optional Ordem As String, Optional Filtro As String) As ADODB.Recordset

    On Error GoTo Erro:

    sSQL = "SELECT * "
    sSQL = sSQL & "FROM tbl_dfc "
    
    If Filtro <> "" Then
        sSQL = sSQL & "WHERE " & Filtro & " "
    End If
    
    If Ordem <> "" Then
        sSQL = sSQL & "ORDER BY " & Ordem
    End If
    
    Set Todos = New ADODB.Recordset
    
    With Todos
        .CursorLocation = adUseServer
        .CursorType = adOpenStatic
        .LockType = adLockReadOnly
        .PageSize = 24
        .Open sSQL, cnn
    End With
    
    Exit Function
    
Erro:

    MsgBox err.Description, vbCritical, "Erro!"
    
End Function
