VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTransferencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Classe

Private Type TClass
    ID          As Long
    Data        As Date
    Valor       As Currency
    CtaOrigID   As Long
    CtaDestID   As Long
End Type

Private This As TClass

Private Const sTabela As String = "tbl_transferencias"

' Propriedades

Public Property Get ID() As Long: ID = This.ID: End Property
Public Property Let ID(ByVal Value As Long): This.ID = CLng(Value): End Property

Public Property Get Data() As Date: Data = This.Data: End Property
Public Property Let Data(ByVal Value As Date): This.Data = Value: End Property

Public Property Get Valor() As Currency: Valor = This.Valor: End Property
Public Property Let Valor(ByVal Value As Currency): This.Valor = Value: End Property

Public Property Get CtaOrigID() As Long: CtaOrigID = This.CtaOrigID: End Property
Public Property Let CtaOrigID(ByVal Value As Long): This.CtaOrigID = Value: End Property

Public Property Get CtaDestID() As Long: CtaDestID = This.CtaDestID: End Property
Public Property Let CtaDestID(ByVal Value As Long): This.CtaDestID = Value: End Property

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
            Me.Data = .Fields("data").Value
            Me.Valor = .Fields("valor").Value
            Me.CtaOrigID = .Fields("ctaorig_id").Value
            Me.CtaDestID = .Fields("ctadest_id").Value
            
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
            .Fields("ctaorig_id").Value = Me.CtaOrigID
            .Fields("ctadest_id").Value = Me.CtaDestID
            
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
