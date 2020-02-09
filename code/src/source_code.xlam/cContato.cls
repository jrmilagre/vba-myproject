VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cContato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Classe

Private Type TClass
    ID         As Long
    Nome       As String
    Nascimento As Variant   ' Date
    Salario    As Variant   ' Currency
    Sexo       As Variant   ' String
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

Public Property Get Sexo() As Variant: Sexo = This.Sexo: End Property
Public Property Let Sexo(ByVal Value As Variant): This.Sexo = Value: End Property

' Métodos

Public Sub CRUD(Acao As Integer, Optional ID As Long)
    
    If ID > 0 Then
    
        sSQL = "SELECT * FROM " & sTabela & " WHERE id = " & ID
        
    End If

    Set rst = New ADODB.Recordset
    
    If Acao = Read Then ' READ
        
        With rst
        
            .Open sSQL, cnn, adOpenStatic, adLockReadOnly
        
            Me.ID = .Fields("id").Value
            Me.Nome = .Fields("nome").Value
            If IsNull(.Fields("nascimento").Value) Then: Me.Nascimento = "": Else: Me.Nascimento = .Fields("nascimento").Value
            If IsNull(.Fields("salario").Value) Then: Me.Salario = 0: Else: Me.Salario = .Fields("salario").Value
            If IsNull(.Fields("sexo").Value) Then: Me.Sexo = Null: Else: Me.Sexo = .Fields("sexo").Value
            
        End With
        
    ElseIf Acao = Create Or Acao = Update Then
        
        With rst
        
            If Acao = Create Then
            
                .Open sTabela, cnn, adOpenDynamic, adLockOptimistic, adCmdTable
                
                .AddNew
                
                Me.ID = .Fields("id").Value
                
            ElseIf Acao = Update Then
            
                .Open sSQL, cnn, adOpenDynamic, adLockOptimistic
                
            End If
            
            .Fields("nome").Value = Me.Nome
            .Fields("nascimento").Value = Me.Nascimento
            .Fields("salario").Value = Me.Salario
            .Fields("sexo").Value = Me.Sexo
            
            .Update
            
        End With
        
    ElseIf Acao = Delete Then
    
        With rst
        
            .Open sSQL, cnn, adOpenDynamic, adLockOptimistic
            
            .Delete
            
        End With
        
    End If
    
    Set rst = Nothing

End Sub
Public Function Todos(Optional Ordem As String) As ADODB.Recordset

    sSQL = "SELECT * "
    sSQL = sSQL & "FROM tbl_contatos "
    
'    sSQL = "SELECT nome, id, nascimento, salario "
'    sSQL = sSQL & "FROM tbl_contatos "
'    sSQL = sSQL & "WHERE nome LIKE ('%" & txbPesquisa.Text & "%') "
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
    
End Function
