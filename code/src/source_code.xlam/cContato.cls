VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cContato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Classe

Private Type TClass
    ID         As Long
    Nome       As String
    Nascimento As Variant   ' Date
    Salario    As Variant   ' Currency
    Genero     As Variant   ' String
End Type

Private This As TClass

Private Const sTabela As String = "tbl_contatos"

' Propriedades

Public Property Get ID() As Long: ID = This.ID: End Property
Public Property Let ID(ByVal Value As Long): This.ID = CLng(Value): End Property

Public Property Get Nome() As String: Nome = This.Nome: End Property
Public Property Let Nome(ByVal Value As String): This.Nome = Value: End Property

Public Property Get Nascimento() As Variant: Nascimento = This.Nascimento: End Property
Public Property Let Nascimento(ByVal Value As Variant): This.Nascimento = Value: End Property

Public Property Get Salario() As Variant: Salario = This.Salario: End Property
Public Property Let Salario(ByVal Value As Variant): This.Salario = Value: End Property

Public Property Get Genero() As Variant: Genero = This.Genero: End Property
Public Property Let Genero(ByVal Value As Variant): This.Genero = Value: End Property

' Métodos

Public Sub CRUD(Acao As Integer, Optional ID As Long)

    On Error GoTo Erro
    
    If Acao <> eCrud.Create Then
    
        sSQL = "SELECT * FROM " & sTabela & " WHERE id = " & ID
        
    End If

    Set rst = New ADODB.Recordset
    
    If Acao = eCrud.Read Then ' READ
        
        With rst
        
            .Open sSQL, cnn, adOpenStatic, adLockReadOnly
        
            Me.ID = .Fields("id").Value
            Me.Nome = .Fields("nome").Value
            If IsNull(.Fields("nascimento").Value) Then: Me.Nascimento = "": Else: Me.Nascimento = .Fields("nascimento").Value
            If IsNull(.Fields("salario").Value) Then: Me.Salario = 0: Else: Me.Salario = .Fields("salario").Value
            If IsNull(.Fields("genero").Value) Then: Me.Genero = Null: Else: Me.Genero = .Fields("genero").Value
            
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
            
            .Fields("nome").Value = Me.Nome
            .Fields("nascimento").Value = Me.Nascimento
            .Fields("salario").Value = Me.Salario
            .Fields("genero").Value = Me.Genero
            
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
    sSQL = sSQL & "FROM tbl_contatos "
    
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
    
    col.Add "F;Feminino;&HFF80FF"
    col.Add "M;Masculino;&HFF8080"
    
    Set GetLegendas = col

End Function
