VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fExtratos 
   Caption         =   ":: Movimentações ::"
   ClientHeight    =   8580
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16980
   OleObjectBlob   =   "fExtratos.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fExtratos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private oConta          As New cConta
Private oCategoria      As New cCategoria
'Private oContaPara      As New cContaPara
Private oDfc                As New cDfc
Private oFornecedor     As New cFornecedor
Private oLoja               As New cLoja

Private oMovFin     As New cMovFin
Private oTransferencia  As New cTransferencia
'Private oAgendamento    As New cAgendamento
Private colControles    As New Collection
'Private sDecisao        As String
'Private iDias           As Integer
Private myRst           As ADODB.Recordset
Private bAtualizaScrool     As Boolean
'Private bSelectConta    As Boolean

Private Sub UserForm_Initialize()
    
    Call Conecta
    
    Call lstContasPopular
    
    Call PopulaCombos
    
    Call Eventos
    
    Call btnCancelar_Click
    
'    Call cbbGrupoPopular
'    Call cbbContaPopular
'    Call cbbFornecedorPopular
'    Call EventosCampos("tbl_mov_fin")
'    Call Campos("Desabilitar")
'
'    btnAlterar.Enabled = False
'    btnExcluir.Enabled = False
'    btnConfirmar.Visible = False
'    btnCancelar.Visible = False
'    lblContaPara.Visible = False
'    cbbContaPara.Visible = False
'
'    btnPaginaInicial.Enabled = False
'    btnPaginaAnterior.Enabled = False
'    btnPaginaSeguinte.Enabled = False
'    btnPaginaFinal.Enabled = False
'    lblPaginas.Enabled = False
'    scrPagina.Enabled = False
'
'    btnIncluir.SetFocus

End Sub
Private Sub UserForm_Terminate()

    Set myRst = Nothing
    
    Call Desconecta
    
End Sub
Private Sub btnValor_Click()
    ccurVisor = IIf(txbValor.Text = "", 0, CCur(txbValor.Text))
    txbValor.Text = Format(GetCalculadora, "#,##0.00")
End Sub
Private Sub lstContas_Change()

'    bSelectConta = True

    Dim lContaID As Long
    
    lContaID = CLng(lstContas.List(lstContas.ListIndex, 1))
    
    If lstContas.ListIndex > -1 Then
    
        Call BuscaRegistros(lContaID)
    
    End If


'
'    Dim lConta  As Long
'    Dim lPagina As Long
'
'    lConta = CLng(lstContas.List(lstContas.ListIndex, 1))
'
'    If lstContas.ListIndex > -1 Then
'
'        lstPrincipal.ListIndex = -1
'        oConta.CRUD eCrud.Read, lConta
'
'        Set myRst = New ADODB.Recordset
'        Set myRst = oMovimentacao.RetornaMovimentacoes(lConta)
'
'        With scrPagina
'            .Min = 1
'            .Max = myRst.PageCount
'        End With
'
'        myRst.AbsolutePage = myRst.PageCount
'
'        scrPagina.Value = myRst.AbsolutePage
'
'        Call lstMovimentacoesPopular(scrPagina.Value)
'
'        Call Campos("Limpar")
'        btnAlterar.Enabled = False
'        btnExcluir.Enabled = False
'    End If
'
'    ' Bloqueia botões de navegação
'    btnPaginaInicial.Enabled = True
'    btnPaginaAnterior.Enabled = True
'    btnPaginaSeguinte.Enabled = True
'    btnPaginaFinal.Enabled = True
'    lblPaginas.Enabled = True
'    scrPagina.Enabled = True
'
'    bSelectConta = False
End Sub
'Private Sub btnIncluir_Click()
    
'    sDecisao = "Inclusão"
'
'    lstContas.Enabled = False
'    lstPrincipal.Enabled = False
'
'    btnConfirmar.Visible = True
'    btnCancelar.Visible = True
'    btnConfirmar.Caption = "Confirmar " & vbNewLine & sDecisao
'    btnCancelar.Caption = "Cancelar " & vbNewLine & sDecisao
'
'    btnIncluir.Visible = False
'    btnAlterar.Visible = False
'    btnExcluir.Visible = False
'
'    Call Campos("Limpar")
'    Call Campos("Habilitar")
'
'    txbLiquidado.Text = Format(Date, "dd/mm/yyyy")
'    txbValor.Text = Format(0, "#,##0.00")
'
'    cbbContaDe.SetFocus
    
'End Sub
'Private Sub btnAlterar_Click()
    
'    sDecisao = "Alteração"
'
'    btnConfirmar.Visible = True
'    btnCancelar.Visible = True
'    btnConfirmar.Caption = "Confirmar " & Chr(13) & sDecisao
'    btnCancelar.Caption = "Cancelar " & Chr(13) & sDecisao
'
'    btnConfirmar.SetFocus
'
'    btnIncluir.Visible = False
'    btnAlterar.Visible = False
'    btnExcluir.Visible = False
'
'    Call Campos("Habilitar")
'
'    cbbCategoria.Enabled = True
'    cbbSubcategoria.Enabled = True
'
'    lstPrincipal.Enabled = False
'    lstContas.Enabled = False
'
'    cbbContaDe.SetFocus
'End Sub
'Private Sub btnExcluir_Click()
    
'    sDecisao = "Exclusão"
'
'    btnConfirmar.Visible = True
'    btnCancelar.Visible = True
'    btnConfirmar.Caption = "Confirmar " & Chr(13) & sDecisao
'    btnCancelar.Caption = "Cancelar " & Chr(13) & sDecisao
'
'    btnConfirmar.SetFocus
'
'    btnIncluir.Visible = False
'    btnAlterar.Visible = False
'    btnExcluir.Visible = False
'
'    lstPrincipal.Enabled = False
'
'    lstContas.Enabled = False

'End Sub
'Private Sub btnConfirmar_Click()
    
'    Dim vbResposta As VbMsgBoxResult
'    Dim sContaSelecionada As String
'    Dim lPaginaAtual As Long
'
'    If Valida = True Then
'
'        If sDecisao = "Inclusão" Then
'            vbResposta = MsgBox("Deseja confirmar a " & sDecisao & " do registro?", vbYesNo, sDecisao & " do registro")
'            If vbResposta = VBA.vbYes Then oMovimentacao.Inclui IsTransferencia:=chbTransferencia.Value, IsAgendamento:=False
'        ElseIf sDecisao = "Alteração" Then
'            vbResposta = MsgBox("Deseja confirmar a " & sDecisao & " do registro?", vbYesNo, sDecisao & " do registro")
'            If vbResposta = VBA.vbYes Then
'                oMovimentacao.Altera chbTransferencia.Value, oMovimentacao.Valor
'                MsgBox sDecisao & " realizada com sucesso!", vbInformation, sDecisao
'            End If
'        ElseIf sDecisao = "Exclusão" Then
'            vbResposta = MsgBox("Deseja confirmar a " & sDecisao & " do registro?", vbYesNo, sDecisao & " do registro")
'            If vbResposta = VBA.vbYes Then
'                oMovimentacao.Exclui IIf(oMovimentacao.Grupo = "T", True, False), oMovimentacao.Valor, IIf(oMovimentacao.Origem = "Agendamento", True, False)
'            End If
'
'        End If
'
'        Call Campos("Limpar")
'        Call Campos("Desabilitar")
'
'        lPaginaAtual = Replace(Mid(lblPaginas.Caption, 1, InStr(1, lblPaginas.Caption, " de")), "Página ", "")
'
'        ' Se houver uma conta selecionada
'        If lstContas.ListIndex > -1 Then
'            sContaSelecionada = lstContas.List(lstContas.ListIndex, 0)
'            Call lstContasPopular
'            lstContas.Text = sContaSelecionada
'            Call lstMovimentacoesPopular(lPaginaAtual)
'        Else
'            Call lstContasPopular
'            Call lstMovimentacoesPopular(lPaginaAtual)
'        End If
'
'        btnConfirmar.Visible = False
'        btnCancelar.Visible = False
'
'        btnIncluir.Visible = True
'        btnAlterar.Visible = True
'        btnExcluir.Visible = True
'        Call Campos("Desabilitar")
'        Call Campos("Limpar")
'
'        btnAlterar.Enabled = False
'        btnExcluir.Enabled = False
'        btnIncluir.SetFocus
'
'        lstPrincipal.Enabled = True
'        lstContas.Enabled = True
'
'    End If

'End Sub
'+-------------------------------------------------------+
'|                                                       |
'| ROTINAS E FUNÇÕES                                     |
'|                                                       |
'+-------------------------------------------------------+

Private Function Valida() As Boolean
    
'    Valida = False
'
'    If cbbContaDe.Text = Empty Then
'        MsgBox "Informe a 'Conta'", vbCritical, "Campo obrigatório"
'    ElseIf txbLiquidado.Text = Empty Then
'        MsgBox "Informe a 'Data'", vbCritical, "Campo obrigatório"
'    Else
'
'        ' Se for uma Receita ou Despesa
'        If chbTransferencia.Value = False Then
'
'            If txbValor.Text = Empty Or txbValor.Text = 0 Then
'                MsgBox "Informe o 'Valor'", vbCritical, "Campo obrigatório"
'            ElseIf cbbFornecedor.Text = Empty Then
'                MsgBox "Informe o 'Fornecedor'", vbCritical, "Campo obrigatório"
'            ElseIf cbbCategoria.Text = Empty Then
'                MsgBox "Informe a 'Categoria'", vbCritical, "Campo obrigatório"
'            ElseIf cbbSubcategoria.Text = Empty Then
'                MsgBox "Informe a 'Subcategoria'", vbCritical, "Campo obrigatório"
'            Else
'                With oMovimentacao
'                    .ContaID = CLng(cbbContaDe.List(cbbContaDe.ListIndex, 1))
'                    .Valor = CCur(txbValor.Text)
'                    .Liquidado = CDate(txbLiquidado.Text)
'                    .Grupo = cbbGrupo.List(cbbGrupo.ListIndex, 1)
'                    .FornecedorID = CLng(cbbFornecedor.List(cbbFornecedor.ListIndex, 1))
'                    .CategoriaID = CLng(cbbCategoria.List(cbbCategoria.ListIndex, 1))
'                    .SubcategoriaID = CLng(cbbSubcategoria.List(cbbSubcategoria.ListIndex, 1))
'                    .Observacao = txbObservacao.Text
'                End With
'
'                Valida = True
'
'            End If
'
'        ' Se for uma transferência
'        Else
'            If txbValor.Text = Empty Or txbValor.Text = 0 Then
'                MsgBox "Informe o 'Valor'", vbCritical, "Campo obrigatório"
'            ElseIf cbbContaPara.Text = Empty Then
'                MsgBox "Informe a 'Conta destino'", vbCritical, "Campo obrigatório":
'            Else
'                With oMovimentacao
'                    .Grupo = "T"
'                    .ContaID = CLng(cbbContaDe.List(cbbContaDe.ListIndex, 1))
'                    .ContaParaID = CLng(cbbContaPara.List(cbbContaPara.ListIndex, 1))
'                    .Liquidado = CDate(txbLiquidado.Text)
'                    '.ContaParaID = CLng(cbbContaPara.List(cbbContaPara.ListIndex, 1))
'                    .Valor = CDbl(txbValor.Text)
'                    .Observacao = txbObservacao.Text
'                End With
'
'                Valida = True
'
'            End If
'
'        End If
'
'    End If
    
    
End Function
Private Sub lstContasPopular()

    Dim rTmp    As New ADODB.Recordset
    
    Set rTmp = oConta.Todos
    
    With lstContas
        .Clear
        .Enabled = True
        .ColumnCount = 2
        .ColumnWidths = "80 pt; 0pt"
        .Font = "Consolas"
        
         While Not rTmp.EOF = True
         
            .AddItem
            .List(.ListCount - 1, 0) = rTmp.Fields("conta").Value
            .List(.ListCount - 1, 1) = rTmp.Fields("id").Value
            
            rTmp.MoveNext
            
         Wend
        
    End With
    
    Set rTmp = Nothing
    
End Sub
Private Sub lstMovimentacoesPopular(Pagina As Long)

'    Dim col         As New Collection
'    Dim curSaldo    As Currency
'    Dim lCount      As Long
'    Dim iItems      As Integer
'    Dim lPosicao    As Long
'
'    oConta.Carrega CLng(lstContas.List(lstContas.ListIndex, 1))
'    curSaldo = oConta.SaldoInicial
'
'    With lstPrincipal
'        .Clear                              ' Limpa ListBox
'        .Enabled = True                     ' Habilita ListBox
'        .ColumnCount = 7   ' Determina número de colunas
'        .ColumnWidths = "0pt; 55pt; 120pt; 220pt; 65pt; 65pt; 90pt"   'Configura largura das colunas
'        .Font = "Consolas"
'
'        lCount = 1
'
'        lPosicao = myRst.AbsolutePosition
'        curSaldo = curSaldo + CalculaAcumulado(lPosicao)
'        myRst.AbsolutePosition = lPosicao
'
'        While Not myRst.EOF = True And lCount <= myRst.PageSize
'
'            .AddItem
'            .List(.ListCount - 1, 0) = myRst.Fields("id").Value
'            .List(.ListCount - 1, 1) = myRst.Fields("liquidado").Value
'
'            If myRst.Fields("grupo").Value = "T" Then
'
'                oTransferencia.Carrega myRst.Fields("transferencia_id").Value
'                oConta.Carrega oMovimentacao.CarregaContaID(oTransferencia.MovimentacaoDeID)
'                oContaPara.Carrega oMovimentacao.CarregaContaID(oTransferencia.MovimentacaoParaID)
'
'                If myRst.Fields("valor").Value < 0 Then
'                    .List(.ListCount - 1, 2) = "---Transferência-->"
'                    .List(.ListCount - 1, 3) = "Foi para a conta: " & oContaPara.Conta
'                Else
'                    .List(.ListCount - 1, 2) = "<--Transferência---"
'                    .List(.ListCount - 1, 3) = "Veio da conta : " & oConta.Conta
'                End If
'            Else
'                oFornecedor.Carrega myRst.Fields("fornecedor_id").Value
'                oSubcategoria.Carrega myRst.Fields("subcategoria_id").Value
'                oCategoria.Carrega oSubcategoria.CategoriaID
'
'                .List(.ListCount - 1, 2) = oFornecedor.NomeFantasia
'                .List(.ListCount - 1, 3) = oCategoria.Categoria & " : " & oSubcategoria.Subcategoria
'            End If
'
'            .List(.ListCount - 1, 4) = Space(10 - Len(Format(myRst.Fields("valor").Value, "#,##0.00"))) & Format(myRst.Fields("valor").Value, "#,##0.00")
'
'            curSaldo = curSaldo + myRst.Fields("valor").Value
'
'            .List(.ListCount - 1, 5) = Space(12 - Len(Format(curSaldo, "#,##0.00"))) & Format(curSaldo, "#,##0.00")
'
'            If Not myRst.Fields("agendamento_id").Value = 0 Then
'
'                oAgendamento.Carrega myRst.Fields("agendamento_id").Value
'
'                If oAgendamento.Recorrente = False Then
'                    .List(.ListCount - 1, 6) = Format(oAgendamento.ID & "", "00000000") & "-ú"
'                Else
'                    If oAgendamento.Infinito = False Then
'                        .List(.ListCount - 1, 6) = Format(oAgendamento.ID & "", "00000000") & "-" & myRst.Fields("parcela").Value
'                    Else
'                        .List(.ListCount - 1, 6) = Format(oAgendamento.ID & "", "00000000") & "-i"
'                    End If
'
'                End If
'            Else
'                .List(.ListCount - 1, 6) = myRst.Fields("observacao").Value
'            End If
'
'            lCount = lCount + 1
'            myRst.MoveNext
'        Wend
'    End With
'
'    lblPaginas.Caption = "Página " & Pagina & " de " & myRst.PageCount
    
End Sub
Private Function CalculaAcumulado(PosicaoRegistro As Long) As Currency
    
    Dim cValor As Currency
    
    Do While myRst.AbsolutePosition < (PosicaoRegistro)
    
        If myRst.Fields("movimento").Value = "E" Then
            CalculaAcumulado = CalculaAcumulado + myRst.Fields("valor").Value
        Else
            CalculaAcumulado = CalculaAcumulado - myRst.Fields("valor").Value
        End If
        
        myRst.MoveNext
        
    Loop

End Function
Private Sub PopulaCombos()

    ' Carrega combo Movimento
    Dim col As Collection
    Dim n   As Variant
    Dim s() As String
    
    Set col = oCategoria.GetMovimentos
    
    With cbbMovimento
        .Clear
        .ColumnCount = 2
        .ColumnWidths = "15pt;30pt;"
        
        For Each n In col
        
            s() = Split(n, ";")
            
                .AddItem
                .List(.ListCount - 1, 0) = s(1)
                .List(.ListCount - 1, 1) = s(0)
        Next n
    
    End With
    
End Sub
Private Sub Eventos()

    ' Declara variáveis
    Dim oControle   As MSForms.control
    Dim oEvento     As c_Evento
    Dim sTag        As String
    Dim sField()    As String
    Dim sCor()   As String
    
    ' Laço para percorrer todos os TextBox e atribuir eventos
    ' de acordo com o tipo de cada campo
    For Each oControle In Me.Controls
    
        If Len(oControle.Tag) > 0 Then
        
            If TypeName(oControle) = "TextBox" Then
            
                Set oEvento = New c_Evento
                
                With oEvento

                    sField() = Split(oControle.Tag, ".")
                    
                    oControle.ControlTipText = cat.Tables(sField(0)).Columns(sField(1)).Properties("Description").Value
                    
                    .FieldType = cat.Tables(sField(0)).Columns(sField(1)).Type
                    .MaxLength = cat.Tables(sField(0)).Columns(sField(1)).DefinedSize
                    .Nullable = cat.Tables(sField(0)).Columns(sField(1)).Properties("Nullable")
                    
                    Set .cTextBox = oControle
                    
                End With
                
                colControles.Add oEvento
                
            ElseIf TypeName(oControle) = "Label" Then
                
                If Mid(oControle.Tag, 1, 4) = "tbl_" Then
                    
                    sField() = Split(oControle.Tag, ".")
                    
                    If cat.Tables(sField(0)).Columns(sField(1)).Properties("Nullable") = False Then
                        oControle.ForeColor = &HFF0000
                        oControle.ControlTipText = "Preenchimento obrigatório"
                    End If
                
                Else
                
                
                    Set oEvento = New c_Evento
                
                    Set oEvento.cLabel = oControle
                
                    colControles.Add oEvento
                
                    If oControle.Tag = "CAB" Then
                
                        sCor() = Split(oConfig.GetCorInfoCab, " ")
                        oControle.ForeColor = RGB(CInt(sCor(0)), CInt(sCor(1)), CInt(sCor(2)))
                
                    End If
                    
                End If
                
            End If
                
        End If
        
    Next

End Sub
Private Sub BuscaRegistros(ContaID As Long)

    Dim n       As Byte
    Dim o       As control
    
    On Error GoTo Erro
    
    Set myRst = oMovFin.Extrato(ContaID)
    
    If myRst.PageCount > 0 Then
        
        bAtualizaScrool = False
        
        With scrPagina
            .Max = myRst.PageCount
            .Value = myRst.PageCount
        End With
        
        Call lstPrincipalPopular(myRst.PageCount)
        
    Else
    
        lstPrincipal.Clear
        
        For n = 1 To myRst.PageSize
            Set o = Controls("l" & Format(n, "00")): o.BackColor = &H8000000F
        Next n
        
    End If
    
Erro:
    Call btnCancelar_Click
    
End Sub
Private Sub lstPrincipalPopular(Pagina As Long)

    Dim n           As Byte
    Dim oControle   As control
    Dim s()         As String
    Dim vLegenda    As Variant
    Dim cSaldo      As Currency
    
    oConta.CRUD eCrud.Read, CLng(lstContas.List(lstContas.ListIndex, 1))
    
    cSaldo = oConta.SaldoInicial
    
    ' Limpa cores da legenda
    For n = 1 To myRst.PageSize
        Set oControle = Controls("l" & Format(n, "00")): oControle.BackColor = &H8000000F
    Next n
    
    ' Carrega coleção de cores da legenda
    Set oLegenda = oMovFin.GetLegendas

    ' Define página que será exibida do Recordset
    myRst.AbsolutePage = Pagina
    
    With lstPrincipal
        .Clear
        .ColumnCount = 7
        .ColumnWidths = "0pt; 55 pt; 120pt; 214pt; 65pt; 65pt; 100pt;"
        .Font = "Consolas"
        ' COLUNAS:
        ' 0 - id
        ' 1 - data
        ' 2 - fornecedor
        ' 3 - categoria
        ' 4 - valor
        ' 5 - saldo
        ' 6 - historico
        
        cSaldo = cSaldo + CalculaAcumulado(myRst.AbsolutePage)
        
        n = 1
        
        While Not myRst.EOF = True And n <= myRst.PageSize
            
            ' Preenche ListBox
            .AddItem
            
            .List(.ListCount - 1, 0) = myRst.Fields("id").Value
            .List(.ListCount - 1, 1) = myRst.Fields("data").Value
            
            If IsNull(myRst.Fields("transferencia_id").Value) Then
                If Not IsNull(myRst.Fields("fornecedor_id").Value) Then
                    oFornecedor.CRUD eCrud.Read, myRst.Fields("fornecedor_id").Value
                    .List(.ListCount - 1, 2) = oFornecedor.Nome
                Else
                    .List(.ListCount - 1, 2) = "<Não-atribuído>"
                End If
                oCategoria.CRUD eCrud.Read, myRst.Fields("categoria_id").Value
                .List(.ListCount - 1, 3) = oCategoria.Categoria & " : " & oCategoria.Subcategoria
            Else
                .List(.ListCount - 1, 2) = "<Transf. entre contas>"
                oTransferencia.CRUD eCrud.Read, myRst.Fields("transferencia_id").Value
                If myRst.Fields("movimento").Value = "S" Then
                    oConta.CRUD eCrud.Read, oTransferencia.CtaDestID
                    .List(.ListCount - 1, 3) = "TRANSF: FOI PARA A CONTA " & oConta.Conta
                Else
                    oConta.CRUD eCrud.Read, oTransferencia.CtaOrigID
                    .List(.ListCount - 1, 3) = "TRANSF: VEIO DA CONTA " & oConta.Conta
                End If
            End If
            
            .List(.ListCount - 1, 4) = Space(ESPACO_ANTES_VALOR - Len(Format(myRst.Fields("valor").Value, "#,##0.00"))) & Format(myRst.Fields("valor").Value, "#,##0.00")
            
            If myRst.Fields("movimento").Value = "E" Then
                cSaldo = cSaldo + myRst.Fields("valor").Value
            Else
                cSaldo = cSaldo - myRst.Fields("valor").Value
            End If
            
            .List(.ListCount - 1, 5) = Space(ESPACO_ANTES_VALOR - Len(Format(cSaldo, "#,##0.00"))) & Format(cSaldo, "#,##0.00")
            .List(.ListCount - 1, 6) = myRst.Fields("historico").Value
'
'            If IsNull(myRst.Fields("transferencia_id").Value) Then
'                oCategoria.CRUD eCrud.Read, myRst.Fields("categoria_id").Value
'                .List(.ListCount - 1, 3) = oCategoria.Categoria & " : " & oCategoria.Subcategoria
'            Else
'                oTransferencia.CRUD eCrud.Read, myRst.Fields("transferencia_id").Value
'
'                If myRst.Fields("movimento").Value = "S" Then
'                    oConta.CRUD eCrud.Read, oTransferencia.CtaDestID
'                    .List(.ListCount - 1, 3) = "TRANSF: FOI PARA A CONTA " & oConta.Conta
'                Else
'                    oConta.CRUD eCrud.Read, oTransferencia.CtaOrigID
'                    .List(.ListCount - 1, 3) = "TRANSF: VEIO DA CONTA " & oConta.Conta
'                End If
'
'            End If
'
'
'            oConta.CRUD eCrud.Read, myRst.Fields("conta_id").Value
'            .List(.ListCount - 1, 4) = oConta.Conta
            

'

            
            ' Colore a legenda
            
            ' Define o rótulo que receberá a cor
            Set oControle = Controls("l" & Format(n, "00"))

            ' Laço para ler cores armazenadas na coleção de legendas da classe
            If oLegenda.Count > 0 Then

                For Each vLegenda In oLegenda

                    s() = Split(vLegenda, ";")

                    If myRst.Fields("movimento").Value = s(0) Then

                        oControle.BackColor = s(2): Exit For

                    End If

                Next

            End If
            
            ' Próximo registro
            myRst.MoveNext: n = n + 1
            
        Wend
        
    End With
    
    ' Posiciona scroll de navegação em páginas
    lblPaginaAtual.Caption = Pagina
    lblNumeroPaginas.Caption = myRst.PageCount
    bAtualizaScrool = False: scrPagina.Value = CLng(lblPaginaAtual.Caption): bAtualizaScrool = True
    lblTotalRegistros.Caption = Format(myRst.RecordCount, "#,##0")
    
    ' Trata os botões de navegação
    Call TrataBotoesNavegacao

End Sub
Private Sub TrataBotoesNavegacao()

    If CLng(lblPaginaAtual.Caption) = myRst.PageCount And CLng(lblPaginaAtual.Caption) > 1 Then
    
        btnPaginaInicial.Enabled = True
        btnPaginaAnterior.Enabled = True
        btnPaginaFinal.Enabled = False
        btnPaginaSeguinte.Enabled = False
        
    ElseIf CLng(lblPaginaAtual.Caption) < myRst.PageCount And CLng(lblPaginaAtual.Caption) = 1 Then
    
        btnPaginaInicial.Enabled = False
        btnPaginaAnterior.Enabled = False
        btnPaginaFinal.Enabled = True
        btnPaginaSeguinte.Enabled = True
        
    ElseIf CLng(lblPaginaAtual.Caption) = myRst.PageCount And CLng(lblPaginaAtual.Caption) = 1 Then
    
        btnPaginaInicial.Enabled = False
        btnPaginaAnterior.Enabled = False
        btnPaginaFinal.Enabled = False
        btnPaginaSeguinte.Enabled = False
    
    Else
    
        btnPaginaInicial.Enabled = True
        btnPaginaAnterior.Enabled = True
        btnPaginaFinal.Enabled = True
        btnPaginaSeguinte.Enabled = True
        
    End If

End Sub
Private Sub btnPaginaInicial_Click()
    
    Call lstPrincipalPopular(1)
    
End Sub
Private Sub btnPaginaAnterior_Click()

    Call lstPrincipalPopular(CLng(lblPaginaAtual.Caption) - 1)
    
End Sub
Private Sub btnPaginaSeguinte_Click()

    Call lstPrincipalPopular(CLng(lblPaginaAtual.Caption) + 1)

End Sub
Private Sub btnPaginaFinal_Click()

    Call lstPrincipalPopular(myRst.PageCount)
    
End Sub
Private Sub btnRegistroAnterior_Click()

        If lstPrincipal.ListIndex > 0 Then
        
            lstPrincipal.ListIndex = lstPrincipal.ListIndex - 1
            
        ElseIf lstPrincipal.ListIndex = 0 And CLng(lblPaginaAtual.Caption) > 1 Then
            
            Call lstPrincipalPopular(CLng(lblPaginaAtual.Caption) - 1)
            
            lstPrincipal.ListIndex = myRst.PageSize - 1
            
        ElseIf CLng(lblPaginaAtual.Caption) = 1 And lstPrincipal.ListIndex = 0 Then
        
            MsgBox "Primeiro registro"
            Exit Sub
            
        Else
        
            lstPrincipal.ListIndex = -1
            
        End If
        
End Sub
Private Sub btnRegistroSeguinte_Click()

    If lstPrincipal.ListIndex = -1 Then
        
        lstPrincipal.ListIndex = 0
    
    ElseIf lstPrincipal.ListIndex = myRst.PageSize - 1 And CLng(lblPaginaAtual.Caption) < myRst.PageCount Then
        
        Call lstPrincipalPopular(CLng(lblPaginaAtual.Caption) + 1)
        
        lstPrincipal.ListIndex = 0
        
    ElseIf CLng(lblPaginaAtual.Caption) = myRst.PageCount And (lstPrincipal.ListIndex + 1) = lstPrincipal.ListCount Then
    
        MsgBox "Último registro"
        Exit Sub
        
    Else
    
        lstPrincipal.ListIndex = lstPrincipal.ListIndex + 1
    
    End If
    
End Sub
Private Sub scrPagina_Change()

    If bAtualizaScrool = True Then
        
        Call lstPrincipalPopular(scrPagina.Value)
        
    End If

End Sub
Private Sub lstPrincipal_Change()

    Dim n As Long
    
    If lstPrincipal.ListIndex >= 0 Then
    
        btnAlterar.Enabled = True
        btnExcluir.Enabled = True
    
        With oMovFin
    
            .CRUD Acao:=eCrud.Read, _
                  Transferencia:=chbTransferencia.Value, _
                  ID:=(CLng(lstPrincipal.List(lstPrincipal.ListIndex, 0)))
            
            If IsNull(.TransferenciaID) Then
                chbTransferencia.Value = False
                oCategoria.CRUD eCrud.Read, .CategoriaID
                txbCategoriaID.Text = oCategoria.ID: txbCategoriaID.TextAlign = fmTextAlignRight
                txbCategoriaInfo.Text = oCategoria.Categoria & " : " & oCategoria.Subcategoria
                cbbMovimento.Value = .Movimento
                
                If IsNull(.DfcID) Then
                    txbDfcID.Text = Empty: txbDfcInfo.Text = ""
                Else
                    oDfc.CRUD eCrud.Read, .DfcID
                    txbDfcID.Text = oDfc.ID: txbDfcID.TextAlign = fmTextAlignRight
                    txbContaInfo.Text = " " & oConta.Conta
                End If
                
                If IsNull(.FornecedorID) Then
                    txbFornecedorID.Text = Empty: txbFornecedorInfo.Text = ""
                Else
                    oFornecedor.CRUD eCrud.Read, .FornecedorID
                    txbFornecedorID.Text = oFornecedor.ID: txbFornecedorID.TextAlign = fmTextAlignRight
                    txbFornecedorInfo.Text = oFornecedor.Nome
                End If
                
                If IsNull(.LojaID) Then
                    txbLojaID.Text = Empty: txbLojaInfo.Text = ""
                Else
                    oLoja.CRUD eCrud.Read, .LojaID
                    txbLojaID.Text = oLoja.ID: txbLojaID.TextAlign = fmTextAlignRight
                    txbFornecedorInfo.Text = oLoja.Nome
                End If
                
                oConta.CRUD eCrud.Read, .ContaID
                txbContaID.Text = oConta.ID: txbContaID.TextAlign = fmTextAlignRight
                txbContaInfo.Text = oConta.Conta

                txbDataCompra.Text = IIf(IsNull(.DataCompra), "", .DataCompra)
                txbHistorico.Text = .Historico
            Else
                chbTransferencia.Value = True
                oTransferencia.CRUD eCrud.Read, .TransferenciaID
                
                oConta.CRUD eCrud.Read, oTransferencia.CtaOrigID
                txbContaID.Text = oConta.ID: txbContaID.TextAlign = fmTextAlignRight: txbContaInfo.Text = oConta.Conta
                
                oConta.CRUD eCrud.Read, oTransferencia.CtaDestID
                txbCtaDestID.Text = oConta.ID: txbCtaDestID.TextAlign = fmTextAlignRight: txbCtaDestInfo.Text = oConta.Conta
                
            End If
    
            lblCabID.Caption = IIf(.ID = 0, "", .ID)
            txbData.Text = .Data
            txbValor.Text = Format(.Valor, "#,##0.00")
            
        End With
        
    End If

End Sub
Private Sub chbTransferencia_Click()

    Dim b As Boolean
    
    If chbTransferencia.Value = True Then
        b = True
        lblConta.Caption = "Conta origem"
        lblHistorico.Left = 6: lblHistorico.Top = 354
        txbHistorico.Left = 6: txbHistorico.Top = 366
        lblCtaDest.Left = 174: lblCtaDest.Top = 318
        txbCtaDestID.Left = 174: txbCtaDestID.Top = 330
        btnCtaDestID.Left = 210: btnCtaDestID.Top = 330
        txbCtaDestInfo.Left = 228: txbCtaDestInfo.Top = 330
    Else
        b = False
        lblConta.Caption = "Conta"
        lblHistorico.Left = 90: lblHistorico.Top = 390
        txbHistorico.Left = 90: txbHistorico.Top = 402
    End If

    lblCategoria.Visible = Not b: txbCategoriaID.Visible = Not b: btnCategoriaID.Visible = Not b: txbCategoriaInfo.Visible = Not b
    lblMovimento.Visible = Not b: cbbMovimento.Visible = Not b
    lblDFC.Visible = Not b: txbDfcID.Visible = Not b: btnDfcID.Visible = Not b: txbDfcInfo.Visible = Not b
    lblFornecedor.Visible = Not b: txbFornecedorID.Visible = Not b: btnFornecedorID.Visible = Not b: txbFornecedorInfo.Visible = Not b
    lblLoja.Visible = Not b: txbLojaID.Visible = Not b: btnLojaID.Visible = Not b: txbLojaInfo.Visible = Not b
    lblDataCompra.Visible = Not b: txbDataCompra.Visible = Not b: btnDataCompra.Visible = Not b
    lblCtaDest.Visible = b: txbCtaDestID.Visible = b: btnCtaDestID.Visible = b: txbCtaDestInfo.Visible = b

End Sub
Private Sub btnCancelar_Click()
    
    btnIncluir.Visible = True: btnAlterar.Visible = True: btnExcluir.Visible = True
    btnConfirmar.Visible = False: btnCancelar.Visible = False
    
    Call Campos("Limpar")
    Call Campos("Desabilitar")
    
    btnAlterar.Enabled = False
    btnExcluir.Enabled = False
    btnIncluir.SetFocus
    
    If oGlobal.ModoAbrir = eModoAbrirForm.Cadastro Then
        lstPrincipal.ListIndex = -1 ' Tira a seleção
    Else
        lstContas.ListIndex = 0
        lstContas.SetFocus
    End If
    
End Sub
Private Sub Campos(Acao As String)
    
    Dim sDecisao    As String
    Dim b           As Boolean
    
    sDecisao = Replace(btnConfirmar.Caption, "Confirmar ", "")
    
    If Acao <> "Limpar" Then
    
        If Acao = "Desabilitar" Then
            b = False
        ElseIf Acao = "Habilitar" Then
            b = True
        End If
        
        chbTransferencia.Enabled = b
        txbData.Enabled = b: lblData.Enabled = b: btnData.Enabled = b
        txbValor.Enabled = b: lblValor.Enabled = b: btnValor.Enabled = b
        txbCategoriaID.Enabled = b: lblCategoria.Enabled = b: btnCategoriaID.Enabled = b
        cbbMovimento.Enabled = b: lblMovimento.Enabled = b
        txbContaID.Enabled = b: lblConta.Enabled = b: btnContaID.Enabled = b
        txbDfcID.Enabled = b: lblDFC.Enabled = b: btnDfcID.Enabled = b
        txbFornecedorID.Enabled = b: lblFornecedor.Enabled = b: btnFornecedorID.Enabled = b
        txbLojaID.Enabled = b: lblLoja.Enabled = b: btnLojaID.Enabled = b
        txbDataCompra.Enabled = b: lblDataCompra.Enabled = b: btnDataCompra.Enabled = b
        txbHistorico.Enabled = b: lblHistorico.Enabled = b
        txbCtaDestID.Enabled = b: lblCtaDest.Enabled = b: btnCtaDestID.Enabled = b
        
    Else
        
        chbTransferencia.Value = False
        lblCabID.Caption = ""
        txbData.Text = Empty
        txbValor.Text = Format(0, "#,##0.00")
        txbCategoriaID.Text = Empty: txbCategoriaInfo.Text = Empty
        cbbMovimento.ListIndex = -1
        txbContaID.Text = Empty: txbContaInfo.Text = Empty
        txbDfcID.Text = Empty: txbDfcInfo.Text = Empty
        txbFornecedorID.Text = Empty: txbFornecedorInfo.Text = Empty
        txbLojaID.Text = Empty: txbLojaInfo.Text = Empty
        txbDataCompra.Text = Empty
        txbHistorico.Text = Empty
        txbCtaDestID.Text = Empty: txbCtaDestInfo.Text = Empty
             
    End If

End Sub
Private Sub btnIncluir_Click()
    
    Call PosDecisaoTomada("Inclusão")
    
End Sub
Private Sub btnAlterar_Click()
    
    Call PosDecisaoTomada("Alteração")

End Sub
Private Sub btnExcluir_Click()

    Call PosDecisaoTomada("Exclusão")
    
End Sub
Private Sub PosDecisaoTomada(Decisao As String)

    btnCancelar.Visible = True: btnConfirmar.Visible = True
    btnConfirmar.Caption = "Confirmar " & Decisao
    btnCancelar.Caption = "Cancelar " & Decisao
    
    btnIncluir.Visible = False: btnAlterar.Visible = False: btnExcluir.Visible = False
    
    If Decisao <> "Exclusão" Then
        
        Call Campos("Habilitar")
        
        If Decisao = "Inclusão" Then
            
            Call Campos("Limpar")
            chbTransferencia.Value = False: Call chbTransferencia_Click
            txbData.SetFocus: txbData.Text = Date
            lstPrincipal.ListIndex = -1: lstPrincipal.Clear
                
        End If
        
    End If
    
End Sub
Private Sub btnData_Click()
    dtDate = IIf(txbData.Text = Empty, Date, txbData.Text)
    txbData.Text = GetCalendario
End Sub
Private Sub btnDfcID_Click()
    
    If txbDfcID.Text = Empty Then
        oGlobal.PesquisaID = Null
    Else
        oGlobal.PesquisaID = CInt(txbDfcID.Text)
    End If

    f_Dfc.Show

    If Not IsNull(oGlobal.PesquisaID) Then
        oDfc.CRUD eCrud.Read, oGlobal.PesquisaID
        txbDfcID.Text = oDfc.ID: txbDfcID.TextAlign = fmTextAlignRight
        txbDfcInfo.Text = oDfc.Grupo
    Else
        txbDfcID.Text = Empty
        txbDfcInfo.Text = Empty
    End If
    
End Sub
Private Sub txbDfcID_AfterUpdate()
  
    If IsNumeric(txbDfcID.Text) Then
        
        oDfc.CRUD eCrud.Read, CLng(txbDfcID.Text)
        
        If oDfc.ID = 0 Then
        
            txbDfcID.Text = Empty
            txbDfcInfo.Text = "<DFC não existe ou é subtotal!>"
        
        Else
        
            txbDfcInfo.Text = oDfc.Grupo
            
        End If

    ElseIf txbDfcID.Text = Empty Then

        txbDfcID.Text = Empty
        txbDfcInfo.Text = Empty
        
    End If

End Sub
Private Sub btnFornecedorID_Click()

    oGlobal.ModoAbrir = eModoAbrirForm.Pesquisa: fFornecedores.Show
    
    Call PesquisaBtn(oFornecedor, Controls("txbFornecedorID"), Controls("lblFornecedor"), Controls("txbFornecedorInfo"))

End Sub
Private Sub txbFornecedorID_AfterUpdate()

    Call PesquisaTxt(Controls("txbFornecedorID"), Controls("lblFornecedor"), Controls("txbFornecedorInfo"), oFornecedor)
        
End Sub
Private Sub btnCategoriaID_Click()

    oGlobal.ModoAbrir = eModoAbrirForm.Pesquisa: fCategorias.Show
    
    Call PesquisaBtn(oCategoria, Controls("txbCategoriaID"), Controls("lblCategoria"), Controls("txbCategoriaInfo"))

End Sub
Private Sub txbCategoriaID_AfterUpdate()

    Call PesquisaTxt(Controls("txbCategoriaID"), Controls("lblCategoria"), Controls("txbCategoriaInfo"), oCategoria)
        
End Sub
Private Sub btnContaID_Click()

    oGlobal.ModoAbrir = eModoAbrirForm.Pesquisa: fContas.Show
    
    Call PesquisaBtn(oConta, Controls("txbContaID"), Controls("lblConta"), Controls("txbContaInfo"))

End Sub
Private Sub txbContaID_AfterUpdate()

    Call PesquisaTxt(Controls("txbContaID"), Controls("lblConta"), Controls("txbContaInfo"), oConta)
        
End Sub
Private Sub btnLojaID_Click()

    oGlobal.ModoAbrir = eModoAbrirForm.Pesquisa: fLojas.Show
    
    Call PesquisaBtn(oLoja, Controls("txbLojaID"), Controls("lblLoja"), Controls("txbLojaInfo"))

End Sub
Private Sub txbLojaID_AfterUpdate()

    Call PesquisaTxt(Controls("txbLojaID"), Controls("lblLoja"), Controls("txbLojaInfo"), oLoja)
        
End Sub
Private Sub txbCtaDestID_AfterUpdate()

    Call PesquisaTxt(Controls("txbCtaDestID"), Controls("lblCtaDest"), Controls("txbCtaDestInfo"), oConta)

End Sub
Private Sub btnCtaDestID_Click()

    oGlobal.ModoAbrir = eModoAbrirForm.Pesquisa: fContas.Show
    
    Call PesquisaBtn(oConta, Controls("txbCtaDestID"), Controls("lblCtaDest"), Controls("txbCtaDestInfo"))

End Sub
Private Sub PesquisaTxt(TextBoxID As control, LabelTitulo As control, TextBoxInfo As control, Classe As Object)
    
    If IsNumeric(TextBoxID.Text) Then
        
        Classe.CRUD eCrud.Read, CLng(TextBoxID.Text)
        
        If Classe.ID = 0 Then
        
            TextBoxID.Text = Empty
            TextBoxInfo.Text = "<" & LabelTitulo & " não existe!>"
        
        Else
        
            TextBoxInfo.Text = GetTextBoxInfo(Classe)
            
        End If
        
    ElseIf TextBoxID.Text = Empty Then

        TextBoxID.Text = Empty
        TextBoxInfo.Text = Empty
        
    End If
    
End Sub
Private Sub PesquisaBtn(Classe As Object, TextBoxID As control, LabelTitulo As control, TextBoxInfo As control)
    
    If Not IsNull(oGlobal.PesquisaID) Then

        Classe.CRUD eCrud.Read, oGlobal.PesquisaID
        TextBoxID.Text = Classe.ID: TextBoxID.TextAlign = fmTextAlignRight
        TextBoxInfo.Text = GetTextBoxInfo(Classe)

    Else

        TextBoxID.Text = Empty
        TextBoxInfo.Text = Empty

    End If

End Sub
Private Function GetTextBoxInfo(Classe As Object) As String

    Select Case TypeName(Classe)
        Case "cLoja": GetTextBoxInfo = Classe.Nome
        Case "cFornecedor": GetTextBoxInfo = Classe.Nome
        Case "cCategoria"
            GetTextBoxInfo = " " & Classe.Categoria & " : " & Classe.Subcategoria
            cbbMovimento.Value = Classe.Movimento
            If Not IsNull(Classe.DfcID) Then
                With txbDfcID
                    .Text = Classe.DfcID
                    .TextAlign = fmTextAlignRight
                End With
                oDfc.CRUD eCrud.Read, Classe.DfcID
                txbDfcInfo.Text = oDfc.Grupo
            End If
        Case "cConta": GetTextBoxInfo = Classe.Conta
    End Select

End Function
Private Sub txbCategoriaID_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 115 Then Call btnCategoriaID_Click
End Sub
Private Sub txbContaID_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 115 Then Call btnContaID_Click
End Sub
Private Sub txbDfcID_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 115 Then Call btnDfcID_Click
End Sub
Private Sub txbFornecedorID_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 115 Then Call btnFornecedorID_Click
End Sub
Private Sub txbLojaID_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 115 Then Call btnLojaID_Click
End Sub
Private Sub btnConfirmar_Click()
    
    Call Gravar(Replace(btnConfirmar.Caption, "Confirmar ", ""))
    
End Sub
Private Sub Gravar(Decisao As String)

    Dim vbResposta  As VbMsgBoxResult
    
    On Error GoTo Erro
    
    vbResposta = MsgBox("Deseja realmente fazer a " & Decisao & "?", vbYesNo + vbQuestion, "Pergunta")
    
    If vbResposta = vbYes Then
    
        If Decisao <> "Exclusão" Then
        
            If txbData.Text = Empty Then
                MsgBox "Campo 'Data' é obrigatório", vbCritical: txbData.SetFocus
            ElseIf txbValor.Text = Empty Or CCur(txbValor.Text) = 0 Then
                MsgBox "Campo 'Valor' não preenchido ou inválido", vbCritical: txbValor.SetFocus
            ElseIf txbContaID.Text = Empty Then
                MsgBox "Campo 'Conta' é obrigatório", vbCritical: txbContaID.SetFocus
            Else
                If chbTransferencia.Value = False Then
                    If txbCategoriaID.Text = Empty Then
                        MsgBox "Campo 'Categoria' é obrigatório", vbCritical: txbCategoriaID.SetFocus
                    ElseIf cbbMovimento.ListIndex = -1 Then
                        MsgBox "Campo 'Movimento' é obrigatório", vbCritical: cbbMovimento.SetFocus
                    Else
                        With oMovFin
                            .Data = CDate(txbData.Text)
                            .Valor = CCur(txbValor.Text)
                            .Movimento = cbbMovimento.List(cbbMovimento.ListIndex, 0)
                            .ContaID = CLng(txbContaID.Text)
                            .CategoriaID = CLng(txbCategoriaID.Text)
                            If RTrim(txbDfcID.Text) = "" Then .DfcID = Null Else .DfcID = CLng(txbDfcID.Text)
                            If RTrim(txbFornecedorID.Text) = "" Then .FornecedorID = Null Else .FornecedorID = CLng(txbFornecedorID.Text)
                            If RTrim(txbLojaID.Text) = "" Then .LojaID = Null Else .LojaID = CLng(txbLojaID.Text)
                            .Historico = txbHistorico.Text
                            If RTrim(txbDataCompra.Text) = "" Then .DataCompra = Null Else .DataCompra = CDate(txbDataCompra.Text)
                            
                            If Decisao = "Inclusão" Then
                                .CRUD Acao:=eCrud.Create, _
                                      Transferencia:=chbTransferencia.Value, _
                                      Decisao:=Decisao
                                
                            Else
                                .CRUD Acao:=eCrud.Update, _
                                      Transferencia:=chbTransferencia.Value, _
                                      ID:=.ID, _
                                      Decisao:=Decisao
                            End If
                            
                            Call BuscaRegistros(CLng(lstContas.List(lstContas.ListIndex, 1)))
                            
                        End With
            
                    End If
                    
                ElseIf chbTransferencia.Value = True Then
                    If txbCtaDestID.Text = Empty Then
                        MsgBox "Campo 'Conta destino' é obrigatório", vbCritical: txbCtaDestID.SetFocus
                    ElseIf txbCtaDestID.Text = txbContaID.Text Then
                        MsgBox "Campo 'Conta destino' não pode ser igual a 'Conta origem'", vbCritical: txbCtaDestID.SetFocus
                    Else
                        With oMovFin
                            .Data = CDate(txbData.Text)
                            .Valor = CCur(txbValor.Text)
                            .ContaID = CLng(txbContaID.Text)
                            .CategoriaID = Null
                            .DfcID = Null
                            .FornecedorID = Null
                            .LojaID = Null
                            .Historico = txbHistorico.Text
                            .DataCompra = Null
                            .CtaDestID = CLng(txbCtaDestID.Text)
                            
                            If Decisao = "Inclusão" Then
                                .CRUD Acao:=eCrud.Create, _
                                      Transferencia:=chbTransferencia.Value, _
                                      Decisao:=Decisao
                                
                            Else
                                .CRUD Acao:=eCrud.Update, _
                                      Transferencia:=chbTransferencia.Value, _
                                      ID:=.ID, _
                                      Decisao:=Decisao
                            End If
                            
                            Call BuscaRegistros(CLng(lstContas.List(lstContas.ListIndex, 1)))
    
                        End With
                        
                    End If
                
                End If
                
            End If
        
        Else ' Se for exclusão
        
            oMovFin.CRUD Acao:=eCrud.Delete, _
                         Transferencia:=chbTransferencia.Value, _
                         ID:=oMovFin.ID, _
                         Decisao:=Decisao
            
            Call BuscaRegistros(CLng(lstContas.List(lstContas.ListIndex, 1)))
            
        End If
               
    ElseIf vbResposta = vbNo Then
    
Erro:
        If Decisao = "Exclusão" Then
            
            Call btnCancelar_Click
            
        End If
        
    End If
    
End Sub
