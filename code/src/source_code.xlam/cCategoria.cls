VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cCategoria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Classe

Private Type TClass
    ID           As Long
    Categoria    As String
    Subcategoria As String
    Grupo        As String
    DfcID        As Variant
End Type

Private This As TClass

Private Const sTabela As String = "tbl_categorias"

' Propriedades

Public Property Get ID() As Long: ID = This.ID: End Property
Public Property Let ID(ByVal Value As Long): This.ID = CLng(Value): End Property

Public Property Get Categoria() As String: Categoria = This.Categoria: End Property
Public Property Let Categoria(ByVal Value As String): This.Categoria = Value: End Property

Public Property Get Subcategoria() As String: Subcategoria = This.Subcategoria: End Property
Public Property Let Subcategoria(ByVal Value As String): This.Subcategoria = Value: End Property

Public Property Get Grupo() As String: Grupo = This.Grupo: End Property
Public Property Let Grupo(ByVal Value As String): This.Grupo = Value: End Property

Public Property Get DfcID() As Variant: DfcID = This.DfcID: End Property
Public Property Let DfcID(ByVal Value As Variant): This.DfcID = Value: End Property

' Métodos

Public Sub CRUD(Acao As Integer, Optional ID As Long)
    
    If ID > 0 Then
    
        sSQL = "SELECT * FROM " & sTabela & " WHERE id = " & ID
        
    End If

    Set rst = New ADODB.Recordset
    
    If Acao = eCrud.Read Then ' READ
        
        With rst
        
            .Open sSQL, cnn, adOpenStatic, adLockReadOnly
        
            Me.ID = .Fields("id").Value
            Me.Categoria = .Fields("categoria").Value
            Me.Subcategoria = .Fields("subcategoria").Value
            Me.Grupo = .Fields("grupo").Value
            Me.DfcID = .Fields("dfc_id").Value
            
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
            
            .Fields("categoria").Value = Me.Categoria
            .Fields("subcategoria").Value = Me.Subcategoria
            .Fields("grupo").Value = Me.Grupo
            .Fields("dfc_id").Value = Me.DfcID
            
            .Update
            
        End With
        
    ElseIf Acao = eCrud.Delete Then
    
        With rst
        
            .Open sSQL, cnn, adOpenDynamic, adLockOptimistic
            
            .Delete
            
        End With
        
    End If
    
    Set rst = Nothing

End Sub
Public Function Todos(Optional Ordem As String, Optional Filtro As String) As ADODB.Recordset

    On Error GoTo Erro:

    sSQL = "SELECT * "
    sSQL = sSQL & "FROM tbl_categorias "
    
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
Public Function GetDFC(ID As Integer) As String

    Dim r As New ADODB.Recordset

    sSQL = "SELECT grupo FROM tbl_dfc WHERE id = " & ID
    
    Set r = cnn.Execute(sSQL)
    
    GetDFC = r.Fields("grupo").Value
    
    Set r = Nothing

End Function
Public Function GetGrupos() As Collection

    Dim col As New Collection
    
    col.Add "Receitas,R"
    col.Add "Despesas,D"
    
    Set GetGrupos = col

End Function
