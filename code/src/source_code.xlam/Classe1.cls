VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Classe1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Classe

Private Type TClass
    ID         As Long
    Abreviacao As String
    Descricao  As String
End Type

Private This As TClass

Private Const sTabela As String = "tbl_ums"

' Propriedades

Public Property Get ID() As Long: ID = This.ID: End Property
Public Property Let ID(ByVal Value As Long): This.ID = CLng(Value): End Property

Public Property Get Abreviacao() As String: Abreviacao = This.Abreviacao: End Property
Public Property Let Abreviacao(ByVal Value As String): This.Abreviacao = Value: End Property

Public Property Get Descricao() As String: Descricao = This.Descricao: End Property
Public Property Let Descricao(ByVal Value As String): This.Descricao = Value: End Property

' M�todos

Public Sub CRUD(Acao As Integer, Optional ID As Long, Optional Decisao As String)

    On Error GoTo Erro
    
    If ID > 0 Then
    
        sSQL = "SELECT * FROM " & sTabela & " WHERE id = " & ID
        
    End If

    Set rst = New ADODB.Recordset
    
    If Acao = eCrud.Read Then ' READ
        
        With rst
        
            .Open sSQL, cnn, adOpenStatic, adLockReadOnly
        
            Me.ID = .Fields("id").Value
            Me.Abreviacao = .Fields("abreviacao").Value
            Me.Descricao = .Fields("descricao").Value
            
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
            
            .Fields("abreviacao").Value = Me.Abreviacao
            .Fields("descricao").Value = Me.Descricao
            
            .Update
            
            MsgBox Decisao & " realizada com sucesso.", vbInformation, Decisao & " de registro"
            
        End With
        
    ElseIf Acao = eCrud.Delete Then
    
        With rst
        
            .Open sSQL, cnn, adOpenDynamic, adLockOptimistic
            
            .Delete
            
            MsgBox Decisao & " realizada com sucesso.", vbInformation, Decisao & " de registro"
            
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
    sSQL = sSQL & "FROM tbl_ums "
    
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
Public Function GetLegendas() As Collection

    Dim col As New Collection
    
    'col.Add "F;Feminino;&HFF80FF"
    'col.Add "M;Masculino;&HFF8080"
    
    Set GetLegendas = col

End Function
