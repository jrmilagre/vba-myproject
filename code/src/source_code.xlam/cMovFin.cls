VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cMovFin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Classe

Private Type TClass
    ID              As Long
    Data            As Date
    Valor           As Currency
    Movimento       As String
    ContaID         As Long
    CategoriaID     As Long
    DfcID           As Variant
    FornecedorID    As Variant ' Usa-se Variant quando é permitido valor nulo
    LojaID          As Variant
    Historico       As String
    DataCompra      As Variant
End Type

Private This As TClass

Private Const sTabela As String = "tbl_mov_fin"

' Propriedades

Public Property Get ID() As Long: ID = This.ID: End Property
Public Property Let ID(ByVal Value As Long): This.ID = CLng(Value): End Property

Public Property Get Data() As String: Data = This.Data: End Property
Public Property Let Data(ByVal Value As String): This.Data = Value: End Property

Public Property Get Valor() As Currency: Valor = This.Valor: End Property
Public Property Let Valor(ByVal Value As Currency): This.Valor = Value: End Property

Public Property Get Movimento() As String: Movimento = This.Movimento: End Property
Public Property Let Movimento(ByVal Value As String): This.Movimento = Value: End Property

Public Property Get ContaID() As Long: ContaID = This.ContaID: End Property
Public Property Let ContaID(ByVal Value As Long): This.ContaID = Value: End Property

Public Property Get CategoriaID() As Long: CategoriaID = This.CategoriaID: End Property
Public Property Let CategoriaID(ByVal Value As Long): This.CategoriaID = Value: End Property

Public Property Get DfcID() As Variant: DfcID = This.DfcID: End Property
Public Property Let DfcID(ByVal Value As Variant): This.DfcID = Value: End Property

Public Property Get FornecedorID() As Variant: FornecedorID = This.FornecedorID: End Property
Public Property Let FornecedorID(ByVal Value As Variant): This.FornecedorID = Value: End Property

Public Property Get LojaID() As Variant: LojaID = This.LojaID: End Property
Public Property Let LojaID(ByVal Value As Variant): This.LojaID = Value: End Property

Public Property Get Historico() As String: Historico = This.Historico: End Property
Public Property Let Historico(ByVal Value As String): This.Historico = Value: End Property

Public Property Get DataCompra() As Variant: DataCompra = This.DataCompra: End Property
Public Property Let DataCompra(ByVal Value As Variant): This.DataCompra = Value: End Property
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
            Me.Data = .Fields("data").Value
            Me.Valor = .Fields("valor").Value
            Me.Movimento = .Fields("movimento").Value
            Me.ContaID = .Fields("conta_id").Value
            Me.CategoriaID = .Fields("categoria_id").Value
            If IsNull(.Fields("dfc_id").Value) Then: Me.DfcID = Null: Else: Me.DfcID = .Fields("dfc_id").Value
            If IsNull(.Fields("fornecedor_id").Value) Then: Me.FornecedorID = Null: Else: Me.FornecedorID = .Fields("fornecedor_id").Value
            If IsNull(.Fields("loja_id").Value) Then: Me.LojaID = Null: Else: Me.LojaID = .Fields("loja_id").Value
            Me.Historico = .Fields("historico").Value
            If IsNull(.Fields("data_compra").Value) Then: Me.DataCompra = Null: Else: Me.DataCompra = .Fields("data_compra").Value
            
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
            
            .Fields("data").Value = Me.Data
            .Fields("valor").Value = Me.Valor
            .Fields("movimento").Value = Me.Movimento
            .Fields("conta_id").Value = Me.ContaID
            .Fields("categoria_id").Value = Me.CategoriaID
            .Fields("dfc_id").Value = Me.DfcID
            .Fields("fornecedor_id").Value = Me.FornecedorID
            .Fields("loja_id").Value = Me.LojaID
            .Fields("historico").Value = Me.Historico
            .Fields("data_compra").Value = Me.DataCompra
            
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

    Dim sOrdem As String

    On Error GoTo Erro:

    sSQL = "SELECT tbl_mov_fin.* "
    sSQL = sSQL & "FROM tbl_mov_fin "
    
    Select Case Replace(Replace(Ordem, " ASC", ""), " DESC", "")
        Case "categoria"
            sSQL = sSQL & "INNER JOIN tbl_categorias ON tbl_categorias.id = tbl_mov_fin.categoria_id "
            sOrdem = Replace(Ordem, "categoria ", "")
            Ordem = "tbl_categorias.categoria " & sOrdem & ", tbl_categorias.subcategoria " & sOrdem
        Case "conta"
            sSQL = sSQL & "INNER JOIN tbl_contas ON tbl_contas.id = tbl_mov_fin.conta_id "
            sOrdem = Replace(Ordem, "conta ", "")
            Ordem = "tbl_contas.conta " & sOrdem & " "
    End Select
    
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
    
    col.Add "E;Entrada;&HFF0000"
    col.Add "S;Saída;&HFF"
    
    Set GetLegendas = col

End Function
