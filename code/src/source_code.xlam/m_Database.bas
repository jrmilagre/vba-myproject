Attribute VB_Name = "m_Database"
Option Explicit

' Desenvolvedor: Jairo Milagre da Fonseca Jr

' BIBLIOTECAS necessárias:
' ---> Microsoft ActiveX Data Objects 2.8 Library
' ---> Microsoft ADO Ext. 2.8 for DDL and Security

' Declaração de variáveis públicas
Public Enum eCrud
    Create = 1
    Read = 2
    Update = 3
    Delete = 4
End Enum

Public cnn      As ADODB.Connection
Public rst      As ADODB.Recordset
Public cat      As ADOX.Catalog
Public sSQL     As String
Public oFiltro  As New c_Filtro
Public oConfig  As New c_Config
Public Function Conecta() As Boolean

    ' Declara variável(is)
    Dim vbResposta As VBA.VbMsgBoxResult
    Dim sCaminho   As String
    Dim sProvedor  As String
    
    ' Cria objeto de conexão com o banco de dados
    Set cnn = New ADODB.Connection
    Set cat = New ADOX.Catalog
    
    sCaminho = oConfig.GetCaminhoBD
    sProvedor = oConfig.GetProvedorDB
    
    ' Inicia a função com o valor Falso, pois a conexão ainda não aconteceu
    Conecta = False
    
    With cnn
        .Provider = sProvedor               ' Escolhe o provedor da conexão
        On Error GoTo Erro                  ' Se a conexão der problema, desvia para o rótulo Erro
        .Open sCaminho                      ' Abre a conexão com o banco de dados
        Set cat.ActiveConnection = cnn      ' Seta catálogo
    End With
    
    ' Se a conexão for um sucesso, retorna Verdadeiro
    Conecta = True
    
    ' Sai da função
    Exit Function
    
Erro:
    ' Mensagem caso a conexão com o banco de dados der problema
    vbResposta = MsgBox("Banco de dados não existe ou não está acessível:" & vbNewLine & _
           vbNewLine & "Caminho do banco procurado: " & vbNewLine & _
           vbNewLine & sCaminho & vbNewLine & vbNewLine & _
           "Deseja criar o arquivo de banco de dados?", vbInformation + vbYesNo)
           
    If vbResposta = vbYes Then
    
        Call CriaBancoDeDados(sCaminho, sProvedor)
        
        Call AtualizaTabelas
        
    Else
        
        Exit Function
        
    End If
    
End Function
Public Sub Desconecta()
    
    cnn.Close           ' Fecha o objeto de conexão
    Set cat = Nothing

End Sub
' Procedimento para criar o banco de dados
' - É necessário habilitar a biblioteca "Microsoft ADO Ext. 2.8 for DDL and Security"
' - para o funcionamento deste procedimento.
Private Sub CriaBancoDeDados(Caminho As String, Provedor As String)
     
    ' Declara variável
    Dim oCatalogo As New ADOX.Catalog
     
    ' Cria o banco de dados
    oCatalogo.Create "Provider=" & Provedor & ";Data Source=" & Caminho
    
    ' Mensagem de conclusão
    MsgBox "Banco de dados criado com sucesso!", vbInformation
    
End Sub

' +------------------------------------------+
' |Tipos de Dados SQL |Tipos de dados do JET |
' +------------------------------------------+
' | BIT               | YES/NO               |
' | BYTE              | NUMERIC - BYTE       |
' | COUNTER           | COUNTER -contador    |
' | CURRENCY          | CURRENCY - Moeda     |
' | DateTime          | DATE/TIME            |
' | SINGLE            | NUMERIC - SINGLE     |
' | DOUBLE            | NUMERIC - DOUBLE     |
' | SHORT             | NUMERIC - INTEGER    |
' | LONG              | NUMERIC - LONG       |
' | LONGTEXT          | MEMO                 |
' | LONGBINARY        | OLE OBJECTS          |
' | Text              | Text                 |
' +------------------------------------------+
Public Sub AtualizaTabelas()

    ' Declara variáveis
    Dim oCatalogo       As New ADOX.Catalog
    Dim sCaminho        As String
    Dim sProvedor       As String
    
    ' Cria o banco de dados se não existir
    On Error GoTo Conecta
    
    sCaminho = oConfig.GetCaminhoBD
    sProvedor = oConfig.GetProvedorDB
    
    oCatalogo.Create "Provider=" & sProvedor & ";Data Source=" & sCaminho

Conecta:
    Set cnn = New ADODB.Connection
    
    ' Abre catálogo
    With cnn
        .Provider = sProvedor                       ' Provedor
        .Open sCaminho
        Set oCatalogo.ActiveConnection = cnn        ' Instancia o catálogo
    End With
    
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '
    Dim FilePath As String
    Dim sText As String
    Dim myArray() As String
    Dim sTableName As String
    
    FilePath = oConfig.GetCaminhoDicionario
    
    Open FilePath For Input As #1
    
    ' Laço para percorrer o arquivo csv que contém o dicionário de dados
    Do Until EOF(1)
    
        Line Input #1, sText
        
        ' Ignora o cabeçalho
        If Mid(sText, 1, 5) <> "table" Then
            
            myArray = Split(sText, ";")
                        
            ' --- VERIFICA SE EXISTE TABELA -------------------------+
            If sTableName <> myArray(0) Then                        '|
                                                                    '|
                Dim oTabela         As New ADOX.Table               '|
                Dim bExisteTabela   As Boolean                      '|
                                                                    '|
                bExisteTabela = False                               '|
                                                                    '|
                For Each oTabela In oCatalogo.Tables                '|
                    If oTabela.Type = "TABLE" Then                  '|
                        If oTabela.Name = myArray(0) Then           '|
                            bExisteTabela = True                    '|
                            Exit For                                '|
                        End If                                      '|
                    End If                                          '|
                Next oTabela                                        '|
            Else                                                    '|
                bExisteTabela = True                                '|
            End If                                                  '|
                                                                    '|
            sTableName = myArray(0)                                 '|
            '--------------------------------------------------------+
            
            ' --- CRIA TABELA SE NÃO EXISTIR ------------------------+
            If bExisteTabela = False Then                           '|
                                                                    '|
                With oTabela                                        '|
                    .Name = myArray(0)                              '|
                    Set .ParentCatalog = oCatalogo                  '|
                End With                                            '|
                                                                    '|
                oCatalogo.Tables.Append oTabela                     '|
            End If                                                  '|
            '--------------------------------------------------------+
            
            '--- VERIFICA SE EXISTE CAMPO ---------------------------+
            Dim oCampo          As ADOX.Column                      '|
            Dim bExisteCampo    As Boolean                          '|
                                                                    '|
            Set oCampo = New ADOX.Column                            '|
            bExisteCampo = False                                    '|
                                                                    '|
            For Each oCampo In oCatalogo.Tables(myArray(0)).Columns '|
                                                                    '|
                If oCampo.Name = myArray(1) Then                    '|
                    bExisteCampo = True                             '|
                    Exit For                                        '|
                End If                                              '|
                                                                    '|
            Next oCampo                                             '|
            '--------------------------------------------------------+
            
            Set oCampo = Nothing
            
            ' Cria o campo na tabela, caso não exista
            If bExisteCampo = False Then
            
                Set oCampo = New ADOX.Column
                
                With oCampo
                    Set .ParentCatalog = oCatalogo
                    .Name = myArray(1)
                    .Type = CInt(myArray(2))
                    
                    If CInt(myArray(2)) = 202 Then
                        .DefinedSize = CInt(myArray(3))
                    End If
                    
                    If CInt(myArray(3)) <> 13 Then
                        .Properties("Nullable").Value = CBool(myArray(4))
                        .Properties("Autoincrement").Value = CBool(myArray(5))
                        .Properties("Description").Value = CStr(myArray(6))
                    End If
                    
                End With
                
                oCatalogo.Tables(myArray(0)).Columns.Append oCampo
                
                ' Cria chave primária
                If CBool(myArray(7)) = True Then
                        
                    Dim idx As ADOX.Index
                    
                    Set idx = New ADOX.Index

                    With idx
                        .Name = "PK_" & myArray(0)
                        .IndexNulls = adIndexNullsAllow
                        .PrimaryKey = True
                        .Unique = True
                        .Columns.Append myArray(1)
                    End With
                    
                    oCatalogo.Tables(myArray(0)).Indexes.Append idx
                    
                    Set idx = Nothing

                End If
                
                ' Cria chave estrangeira
                If myArray(8) <> "False" Then

                    Dim fk As ADOX.Key
                    
                    Set fk = New ADOX.Key
                    
                    Dim fkArr() As String

                    fkArr = Split(myArray(8), ".")

                    With fk
                       .Name = "FK_" & fkArr(0) & "->" & fkArr(1) & "=" & myArray(0) & "->" & myArray(1)
                       .Type = adKeyForeign
                       .RelatedTable = fkArr(0)
                       .Columns.Append myArray(1)
                       .Columns(myArray(1)).RelatedColumn = fkArr(1)
                       .UpdateRule = adRICascade
                    End With

                    oCatalogo.Tables(myArray(0)).Keys.Append fk
                    
                    Set fk = Nothing
                    
                End If
                
                Set oCampo = Nothing
                
            End If
            
            If myArray(0) = "tbl_dfc" And myArray(1) = "sequencia" And bExisteCampo = False Then
        
                Call PopulaTblDFC
            
            End If
        
        End If
        
        ' Se estiver sendo criado o último campo da tabela tbl_dfc,
        ' inclui os dados permanentemente na tabela

    
    Loop
    
    Close #1
    
    ' CONSULTAS (VIEW)
    Dim cmd As ADODB.Command
    Dim vw  As ADOX.View
    Dim bVw As Boolean
    
    FilePath = oConfig.GetCaminhoConsultas
    
    Open FilePath For Input As #1
    
    Do Until EOF(1)
    
        Line Input #1, sText
    
        ' Explode texto da linha pelo delimitador
        myArray() = Split(sText, ";")
        
        If myArray(0) <> "name" Then
        
            bVw = False
            
            For Each vw In oCatalogo.Views
            
                If myArray(0) = vw.Name Then
                    
                    bVw = True
                    
                    Exit For
                    
                End If
                
            Next
            
            If bVw = False Then
            
                ' Instancia novo objeto ADODB.Command
                Set cmd = New ADODB.Command
            
                ' Cria o comando representando a Consulta (View)
                cmd.CommandText = myArray(1)
            
                ' Cria uma nova Consulta (View)
                oCatalogo.Views.Append myArray(0), cmd
                
                ' Limpa objeto
                Set cmd = Nothing
                
            End If
            
        End If
    
    Loop
    
    Close #1
    
    Set oCatalogo = Nothing
    
    Call Desconecta
    
    MsgBox "Tabelas atualizadas com sucesso!", vbInformation

End Sub
Private Function ExisteTabela(Tabela As String) As Boolean
    
    ' Inicia retorno da função como tabela não existente
    ExisteTabela = False
    
    ' Armazena esquema de tabelas no Recordset
    Set rst = cnn.OpenSchema(adSchemaTables)
    
    ' Laço para percorrer todas as tabelas
    Do Until rst.EOF
    
        ' Se a tabela do laço for igual a tabela verificada
        ' significa que existe, então muda o retorno da função
        ' para True e sai da função
        If rst!Table_Name = Tabela Then
            ExisteTabela = True
            GoTo Sair
        End If
        
        ' Move para a próxima tabela
        rst.MoveNext
    Loop
Sair:
    ' Destrói objeto Recordset
    Set rst = Nothing
End Function
Public Sub Backup()
    
    Dim FSO As Object
    
    On Error GoTo Erro
    
    Set FSO = CreateObject("scripting.filesystemobject")
    
    FSO.CopyFile oConfig.GetCaminhoBD, oConfig.GetCaminhoBackup
    
    MsgBox "Backup realizado com sucesso!", vbInformation
    
    Exit Sub
    
Erro:
    
    MsgBox "Problema no Backup!", vbCritical
    
End Sub
Public Sub IncluiRegistrosTeste()

    If Conecta = True Then
        sSQL = "INSERT INTO tbl_unidades_medida ([nome], [abreviacao]) VALUES ('Saco', 'SC') ": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_unidades_medida ([nome], [abreviacao]) VALUES ('Metro cúbico', 'M3') ": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_unidades_medida ([nome], [abreviacao]) VALUES ('Metro quadrado', 'M2') ": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_unidades_medida ([nome], [abreviacao]) VALUES ('Lata', 'LT') ": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_unidades_medida ([nome], [abreviacao]) VALUES ('Metro linear', 'MT') ": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_unidades_medida ([nome], [abreviacao]) VALUES ('Dia', 'DD') ": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_unidades_medida ([nome], [abreviacao]) VALUES ('Hora', 'HH') ": cnn.Execute sSQL
        
        sSQL = "INSERT INTO tbl_produtos ([nome], [um_id]) VALUES ('Cimento', 1) ": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_produtos ([nome], [um_id]) VALUES ('Cal', 1) ": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_produtos ([nome], [um_id]) VALUES ('Areia média', 2) ": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_produtos ([nome], [um_id]) VALUES ('Pedra', 2) ": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_produtos ([nome], [um_id]) VALUES ('Laje treliça H08', 3) ": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_produtos ([nome], [um_id]) VALUES ('Mão de obra', 3) ": cnn.Execute sSQL
        
        sSQL = "INSERT INTO tbl_categorias ([pag_rec], [categoria], [subcategoria]) VALUES ('R', 'Vendas', 'Obras') ": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_categorias ([pag_rec], [categoria], [subcategoria]) VALUES ('R', 'Vendas', 'Carros') ": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_categorias ([pag_rec], [categoria], [subcategoria]) VALUES ('P', 'Despesas com obras', 'Materiais de construção') ": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_categorias ([pag_rec], [categoria], [subcategoria]) VALUES ('P', 'Despesas com obras', 'Serviços') ": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_categorias ([pag_rec], [categoria], [subcategoria]) VALUES ('P', 'Despesas administrativas', 'Salários') ": cnn.Execute sSQL
        
        sSQL = "INSERT INTO tbl_fornecedores ([nome]) VALUES ('Cardoso Materiais para Construção') ": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_fornecedores ([nome]) VALUES ('Orlando Materiais para Construção') ": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_fornecedores ([nome]) VALUES ('Aparecido (Cidinho)') ": cnn.Execute sSQL
        
        sSQL = "INSERT INTO tbl_tipos_obra ([nome]) VALUES ('Casa') ": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_tipos_obra ([nome]) VALUES ('Sobrado') ": cnn.Execute sSQL
        
        sSQL = "INSERT INTO tbl_pedreiros ([nome], [apelido], [preco_m2]) VALUES ('Aparecido', 'Cidinho', 300.00) ": cnn.Execute sSQL
        
        sSQL = "INSERT INTO tbl_estados ([nome], [uf]) VALUES ('Minas Gerais', 'MG') ": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_estados ([nome], [uf]) VALUES ('São Paulo', 'SP') ": cnn.Execute sSQL
        
        sSQL = "INSERT INTO tbl_clientes ([nome]) VALUES ('Acmo Administração de Bens e Participações Eireli') ": cnn.Execute sSQL
        
        sSQL = "INSERT INTO tbl_obras ([endereco], [tipo_id], [bairro], [cidade], [uf], [cliente_id], [data], [categoria_id]) VALUES ('Alameda Joaquim Marcondes da Silveira, 171', 2, 'Campos Olivotti', 'Extrema', 'MG', 1, " & CLng(CDate("14/12/2019")) & ", 1) ": cnn.Execute sSQL
        
        sSQL = "INSERT INTO tbl_titulos_receber ([obra_id], [cliente_id], [observacao], [vencimento], [valor], [data]) VALUES (1, 1, 'Entrada', " & CLng(CDate("14/12/2019")) & ", 104386.08, " & CLng(CDate("14/12/2019")) & ")": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_titulos_receber ([obra_id], [cliente_id], [observacao], [vencimento], [valor], [data]) VALUES (1, 1, '01/09', " & CLng(CDate("14/01/2020")) & ", 27063.05, " & CLng(CDate("14/12/2019")) & ")": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_titulos_receber ([obra_id], [cliente_id], [observacao], [vencimento], [valor], [data]) VALUES (1, 1, '02/09', " & CLng(CDate("14/02/2020")) & ", 27063.05, " & CLng(CDate("14/12/2019")) & ")": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_titulos_receber ([obra_id], [cliente_id], [observacao], [vencimento], [valor], [data]) VALUES (1, 1, '03/09', " & CLng(CDate("14/03/2020")) & ", 27063.05, " & CLng(CDate("14/12/2019")) & ")": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_titulos_receber ([obra_id], [cliente_id], [observacao], [vencimento], [valor], [data]) VALUES (1, 1, '04/09', " & CLng(CDate("14/04/2020")) & ", 27063.05, " & CLng(CDate("14/12/2019")) & ")": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_titulos_receber ([obra_id], [cliente_id], [observacao], [vencimento], [valor], [data]) VALUES (1, 1, '05/09', " & CLng(CDate("14/05/2020")) & ", 27063.05, " & CLng(CDate("14/12/2019")) & ")": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_titulos_receber ([obra_id], [cliente_id], [observacao], [vencimento], [valor], [data]) VALUES (1, 1, '06/09', " & CLng(CDate("14/06/2020")) & ", 27063.05, " & CLng(CDate("14/12/2019")) & ")": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_titulos_receber ([obra_id], [cliente_id], [observacao], [vencimento], [valor], [data]) VALUES (1, 1, '07/09', " & CLng(CDate("14/07/2020")) & ", 27063.05, " & CLng(CDate("14/12/2019")) & ")": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_titulos_receber ([obra_id], [cliente_id], [observacao], [vencimento], [valor], [data]) VALUES (1, 1, '08/09', " & CLng(CDate("14/08/2020")) & ", 27063.05, " & CLng(CDate("14/12/2019")) & ")": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_titulos_receber ([obra_id], [cliente_id], [observacao], [vencimento], [valor], [data]) VALUES (1, 1, '09/09', " & CLng(CDate("14/09/2020")) & ", 27063.05, " & CLng(CDate("14/12/2019")) & ")": cnn.Execute sSQL
        
        sSQL = "INSERT INTO tbl_compras ([data], [fornecedor_id], [categoria_id]) VALUES (" & CLng(CDate("13/12/2019")) & ", 1, 3)": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_compras ([data], [fornecedor_id], [categoria_id]) VALUES (" & CLng(CDate("14/12/2019")) & ", 2, 3)": cnn.Execute sSQL
        
        sSQL = "INSERT INTO tbl_compras_itens ([compra_id], [produto_id], [quantidade], [um_id], [unitario], [total], [data], [fornecedor_id]) VALUES (1, 1, 2, 1, 23.5, 47, " & CLng(CDate("13/12/2019")) & ", 1)": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_compras_itens ([compra_id], [produto_id], [quantidade], [um_id], [unitario], [total], [data], [fornecedor_id]) VALUES (1, 2, 5, 1, 5, 25, " & CLng(CDate("13/12/2019")) & ", 1)": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_compras_itens ([compra_id], [produto_id], [quantidade], [um_id], [unitario], [total], [data], [fornecedor_id]) VALUES (1, 3, 5, 2, 25, 125, " & CLng(CDate("13/12/2019")) & ", 1)": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_compras_itens ([compra_id], [produto_id], [quantidade], [um_id], [unitario], [total], [data], [fornecedor_id]) VALUES (2, 1, 5, 1, 24.5, 122.5, " & CLng(CDate("14/12/2019")) & ", 2)": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_compras_itens ([compra_id], [produto_id], [quantidade], [um_id], [unitario], [total], [data], [fornecedor_id]) VALUES (2, 2, 5, 1, 4.8, 48, " & CLng(CDate("14/12/2019")) & ", 2)": cnn.Execute sSQL
        
        sSQL = "INSERT INTO tbl_titulos_pagar ([compra_id], [fornecedor_id], [observacao], [vencimento], [valor], [data]) VALUES (1, 1, '01/02', " & CLng(CDate("13/01/2019")) & ", 36, " & CLng(CDate("13/12/2019")) & ")": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_titulos_pagar ([compra_id], [fornecedor_id], [observacao], [vencimento], [valor], [data]) VALUES (1, 1, '02/02', " & CLng(CDate("13/02/2020")) & ", 36, " & CLng(CDate("13/12/2019")) & ")": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_titulos_pagar ([compra_id], [fornecedor_id], [observacao], [vencimento], [valor], [data]) VALUES (2, 2, '01/03', " & CLng(CDate("14/01/2020")) & ", 48.83, " & CLng(CDate("14/12/2019")) & ")": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_titulos_pagar ([compra_id], [fornecedor_id], [observacao], [vencimento], [valor], [data]) VALUES (2, 2, '02/03', " & CLng(CDate("14/02/2020")) & ", 48.83, " & CLng(CDate("14/12/2019")) & ")": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_titulos_pagar ([compra_id], [fornecedor_id], [observacao], [vencimento], [valor], [data]) VALUES (2, 2, '03/03', " & CLng(CDate("14/03/2020")) & ", 48.84, " & CLng(CDate("14/12/2019")) & ")": cnn.Execute sSQL
        
        sSQL = "INSERT INTO tbl_contas ([nome], [saldo_inicial]) VALUES ('Dinheiro em caixa', 0) ": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_contas ([nome], [saldo_inicial]) VALUES ('Santander', 0) ": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_contas ([nome], [saldo_inicial]) VALUES ('Bradesco', 0) ": cnn.Execute sSQL
        
        sSQL = "INSERT INTO tbl_etapas ([nome]) VALUES ('Alvenaria') ": cnn.Execute sSQL
        sSQL = "INSERT INTO tbl_etapas ([nome]) VALUES ('Acabamento') ": cnn.Execute sSQL
        
        Call Desconecta
    End If
    
End Sub
Private Sub PopulaTblDFC()

    sSQL = "INSERT INTO tbl_dfc ([id], [grupo], [subtotal], [sequencia]) VALUES (1, '(+) Receita bruta', 0, 1) ": cnn.Execute sSQL
    sSQL = "INSERT INTO tbl_dfc ([id], [grupo], [subtotal], [sequencia]) VALUES (2, '(-) Caixa dispendido na revenda', 0, 2) ": cnn.Execute sSQL
    sSQL = "INSERT INTO tbl_dfc ([id], [grupo], [subtotal], [sequencia]) VALUES (3, 'A) CAIXA BRUTO OBTIDO NAS OPERAÇÕES', 1, 3) ": cnn.Execute sSQL
    sSQL = "INSERT INTO tbl_dfc ([id], [grupo], [subtotal], [sequencia]) VALUES (4, '(-) Despesas com vendas', 0, 4) ": cnn.Execute sSQL
    sSQL = "INSERT INTO tbl_dfc ([id], [grupo], [subtotal], [sequencia]) VALUES (5, '(-) Despesas administrativas', 0, 5) ": cnn.Execute sSQL
    sSQL = "INSERT INTO tbl_dfc ([id], [grupo], [subtotal], [sequencia]) VALUES (6, 'B) CAIXA GERADO NOS NEGÓCIOS', 1, 6) ": cnn.Execute sSQL
    sSQL = "INSERT INTO tbl_dfc ([id], [grupo], [subtotal], [sequencia]) VALUES (7, '(+) Receitas financeiras', 0, 7) ": cnn.Execute sSQL
    sSQL = "INSERT INTO tbl_dfc ([id], [grupo], [subtotal], [sequencia]) VALUES (8, '(-) Despesas financeiras', 0, 8) ": cnn.Execute sSQL
    sSQL = "INSERT INTO tbl_dfc ([id], [grupo], [subtotal], [sequencia]) VALUES (9, 'C) CAIXA LÍQUIDO APÓS FATOS NÃO OPERACIONAIS', 1, 9) ": cnn.Execute sSQL
    sSQL = "INSERT INTO tbl_dfc ([id], [grupo], [subtotal], [sequencia]) VALUES (10, '(-) Juros sobre empréstimos', 0, 10) ": cnn.Execute sSQL
    sSQL = "INSERT INTO tbl_dfc ([id], [grupo], [subtotal], [sequencia]) VALUES (11, 'D) CAIXA LÍQUIDO APÓS REMUNERAÇÃO DO CAPITAL', 1, 11) ": cnn.Execute sSQL
    sSQL = "INSERT INTO tbl_dfc ([id], [grupo], [subtotal], [sequencia]) VALUES (12, '(-) Capital de empréstimo pago', 0, 12) ": cnn.Execute sSQL
    sSQL = "INSERT INTO tbl_dfc ([id], [grupo], [subtotal], [sequencia]) VALUES (13, 'E) CAIXA APÓS AMORTIZAÇÃO DE EMPRÉSTIMOS', 1, 13) ": cnn.Execute sSQL
    sSQL = "INSERT INTO tbl_dfc ([id], [grupo], [subtotal], [sequencia]) VALUES (14, '(+) Aquisição de empréstimo', 0, 14) ": cnn.Execute sSQL
    sSQL = "INSERT INTO tbl_dfc ([id], [grupo], [subtotal], [sequencia]) VALUES (15, 'F) CAIXA APÓS NOVAS FONTES DE RECURSOS', 1, 15) ": cnn.Execute sSQL
    sSQL = "INSERT INTO tbl_dfc ([id], [grupo], [subtotal], [sequencia]) VALUES (16, '(-) Aquisição de permanentes', 0, 16) ": cnn.Execute sSQL
    sSQL = "INSERT INTO tbl_dfc ([id], [grupo], [subtotal], [sequencia]) VALUES (17, 'G) CAIXA LÍQUIDO', 1, 17) ": cnn.Execute sSQL
 
End Sub

