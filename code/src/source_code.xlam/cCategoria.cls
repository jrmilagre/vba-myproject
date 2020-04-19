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
    Movimento    As String
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

Public Property Get Movimento() As String: Movimento = This.Movimento: End Property
Public Property Let Movimento(ByVal Value As String): This.Movimento = Value: End Property

Public Property Get DfcID() As Variant: DfcID = This.DfcID: End Property
Public Property Let DfcID(ByVal Value As Variant): This.DfcID = Value: End Property

' Métodos

Public Sub CRUD(Acao As Integer, Optional ID As Long, Optional Decisao As String)
    
    On Error GoTo Erro
    
    If Acao <> eCrud.Create Then
    
        sSQL = "SELECT * FROM " & sTabela & " WHERE id = " & ID
        
    End If

    Set rst = New ADODB.Recordset
    
    If Acao = eCrud.Read Then ' READ
        
        With rst
        
            .Open sSQL, cnn, adOpenStatic, adLockReadOnly
        
            Me.ID = .Fields("id").Value
            Me.Categoria = .Fields("categoria").Value
            Me.Subcategoria = .Fields("subcategoria").Value
            Me.Movimento = .Fields("movimento").Value
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
            .Fields("movimento").Value = Me.Movimento
            .Fields("dfc_id").Value = Me.DfcID
            
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
Public Function GetMovimentos() As Collection

    Dim col As New Collection
    
    col.Add "ENTRADA;E"
    col.Add "SAÍDA;S"
    
    Set GetMovimentos = col

End Function
Public Function GetLegendas() As Collection

    Dim col As New Collection
    
    col.Add "E;Entrada;&HFF0000"
    col.Add "S;Saída;&HFF"
    
    Set GetLegendas = col

End Function
Public Function Listar(OrderBy As String) As Collection

    Dim col As New Collection

    sSQL = "SELECT id, categoria, subcategoria "
    sSQL = sSQL & "FROM " & sTabela & " "
    sSQL = sSQL & "ORDER BY " & OrderBy

    Set rst = New ADODB.Recordset
    
    rst.Open sSQL, cnn, adOpenStatic

    Do Until rst.EOF = True
        col.Add rst.Fields("id").Value & ";" & rst.Fields("categoria").Value & ";" & rst.Fields("subcategoria").Value
        rst.MoveNext
    Loop

    Set rst = Nothing
    
    Set Listar = col

End Function

