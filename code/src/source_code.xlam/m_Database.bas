Attribute VB_Name = "m_Database"
Option Explicit

' Desenvolvedor: Jairo Milagre da Fonseca Jr

' BIBLIOTECAS necessárias:
' ---> Microsoft ActiveX Data Objects 2.8 Library
' ---> Microsoft ADO Ext. 2.8 for DDL and Security

' Declaração de variáveis públicas
Enum CRUD
    Create = 1
    Read = 2
    Update = 3
    Delete = 4
End Enum

Public cnn  As ADODB.Connection
Public rst  As ADODB.Recordset
Public cat  As ADOX.Catalog
Public sSQL As String
Public Function Conecta() As Boolean
    
    ' Armazena o caminho do banco de dados
    Dim sCaminho As String
    
    sCaminho = Mid(wbCode.Path, 1, Len(wbCode.Path) - 5) & _
           Application.PathSeparator & "data" & _
           Application.PathSeparator & "banco.mdb"
    
    ' Cria objeto de conexão com o banco de dados
    Set cnn = New ADODB.Connection
    Set cat = New ADOX.Catalog
    
    ' Inicia a função com o valor Falso, pois a conexão ainda não aconteceu
    Conecta = False
    
    With cnn
        .Provider = "Microsoft.Jet.OLEDB.4.0"   ' Escolhe o provedor da conexão
        On Error GoTo Erro                      ' Se a conexão der problema, desvia para o rótulo Erro
        .Open sCaminho                          ' Abre a conexão com o banco de dados
        Set cat.ActiveConnection = cnn          ' Seta catálogo
    End With
    
    ' Se a conexão for um sucesso, retorna Verdadeiro
    Conecta = True
    
    ' Sai da função
    Exit Function
    
Erro:
    ' Mensagem caso a conexão com o banco de dados der problema
    MsgBox "Banco de dados não existe ou não está acessível.", vbInformation
    
End Function
Public Sub Desconecta()
    
    cnn.Close           ' Fecha o objeto de conexão
    Set cat = Nothing

End Sub

Public Function ExecutaSQL(ByVal sSQL As String) As ADODB.Recordset
    
    ' Cria o objeto Recordset
    Set ExecutaSQL = New ADODB.Recordset

    ' Abre a tabela
    With ExecutaSQL
        .CursorLocation = adUseServer
        .Open Source:=sSQL, _
              ActiveConnection:=cnn, _
              CursorType:=adOpenDynamic, _
              LockType:=adLockOptimistic, _
              Options:=adCmdText
    End With
                 
End Function


