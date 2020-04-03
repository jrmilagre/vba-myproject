VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fMovFin 
   Caption         =   ":: Movimentações Financeiras ::"
   ClientHeight    =   9105
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10020
   OleObjectBlob   =   "fMovFin.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fMovFin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private oMovFin             As New cMovFin
Private oCategoria          As New cCategoria
Private oConta              As New cConta
Private oDfc                As New cDfc
Private oFornecedor         As New cFornecedor
Private oLoja               As New cLoja
Private colControles        As New Collection               ' Para eventos de campos
Private myRst               As New ADODB.Recordset
Private bAtualizaScrool     As Boolean

Private Sub UserForm_Initialize()
    
    Call PopulaCombos
    
    Call Eventos
    
    Call BuscaRegistros

End Sub

Private Sub UserForm_Terminate()
    
    Set oMovFin = Nothing
    Set myRst = Nothing
    
    Call Desconecta
    
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
    
    MultiPage1.Value = 1
    
    If Decisao <> "Exclusão" Then
    
        If Decisao = "Inclusão" Then
        
            Call Campos("Limpar")
            
        End If
        
        Call Campos("Habilitar")
        
        txbData.SetFocus: txbData.Text = Date
        
    End If
    
    MultiPage1.Pages(0).Enabled = False
    
End Sub
Private Sub btnConfirmar_Click()
    
    Call Gravar(Replace(btnConfirmar.Caption, "Confirmar ", ""))
    
End Sub
Private Sub btnCancelar_Click()
    
    btnIncluir.Visible = True: btnAlterar.Visible = True: btnExcluir.Visible = True
    btnConfirmar.Visible = False: btnCancelar.Visible = False
    
    Call Campos("Limpar")
    Call Campos("Desabilitar")
    
    btnAlterar.Enabled = False
    btnExcluir.Enabled = False
    btnIncluir.SetFocus
   
    MultiPage1.Value = 0
    
    lstPrincipal.ListIndex = -1 ' Tira a seleção
    
End Sub
Private Sub lstPrincipal_Change()

    Dim n As Long
    
    If lstPrincipal.ListIndex >= 0 Then
    
        btnAlterar.Enabled = True
        btnExcluir.Enabled = True
    
        With oMovFin
    
            .CRUD eCrud.Read, (CLng(lstPrincipal.List(lstPrincipal.ListIndex, 0)))
    
            lblCabID.Caption = IIf(.ID = 0, "", .ID)
            lblCabData.Caption = .Data
            txbData.Text = .Data
            txbValor.Text = Format(.Valor, "#,##0.00")
            
            oCategoria.CRUD eCrud.Read, .CategoriaID
            txbCategoriaID.Text = oCategoria.ID: txbCategoriaID.TextAlign = fmTextAlignRight
            lblCategoriaInfo.Caption = oCategoria.Categoria & " : " & oCategoria.Subcategoria
            
            cbbMovimento.Value = .Movimento
            
            oConta.CRUD eCrud.Read, .ContaID
            txbContaID.Text = oConta.ID: txbContaID.TextAlign = fmTextAlignRight
            lblContaInfo.Caption = " " & oConta.Conta
            
            If IsNull(.DfcID) Then
                txbDfcID.Text = Empty: lblDfcInfo.Caption = ""
            Else
                oDfc.CRUD eCrud.Read, .DfcID
                txbDfcID.Text = oDfc.ID: txbDfcID.TextAlign = fmTextAlignRight
                lblContaInfo.Caption = " " & oConta.Conta
            End If

            If IsNull(.FornecedorID) Then
                txbFornecedorID.Text = Empty: lblFornecedorInfo.Caption = ""
            Else
                oFornecedor.CRUD eCrud.Read, .FornecedorID
                txbFornecedorID.Text = oFornecedor.ID: txbFornecedorID.TextAlign = fmTextAlignRight
                lblFornecedorInfo.Caption = " " & oFornecedor.Nome
            End If
            
            If IsNull(.LojaID) Then
                txbLojaID.Text = Empty: lblLojaInfo.Caption = ""
            Else
                oLoja.CRUD eCrud.Read, .LojaID
                txbLojaID.Text = oLoja.ID: txbLojaID.TextAlign = fmTextAlignRight
                lblLojaInfo.Caption = " " & oLoja.Nome
            End If
            
            txbDataCompra.Text = IIf(IsNull(.DataCompra), "", .DataCompra)
            txbHistorico.Text = .Historico
            
        End With
        
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
        
        MultiPage1.Pages(0).Enabled = Not b
        
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
        
    Else
    
        lblCabID.Caption = ""
        lblCabData.Caption = ""
        txbData.Text = Empty
        txbValor.Text = Format(0, "#,##0.00")
        txbCategoriaID.Text = Empty
        cbbMovimento.ListIndex = -1
        txbContaID.Text = Empty
        txbDfcID.Text = Empty
        txbFornecedorID.Text = Empty
        txbLojaID.Text = Empty
        txbDataCompra.Text = Empty
        txbHistorico.Text = Empty
             
    End If

End Sub
Private Sub lstPrincipalPopular(Pagina As Long)

    Dim n           As Byte
    Dim oControle   As control
    Dim s()         As String
    Dim vLegenda    As Variant
    
    ' Limpa cores da legenda
    For n = 1 To myRst.PageSize
        Set oControle = Controls("l" & Format(n, "00")): oControle.BackColor = &H8000000F
    Next n
    
    ' Carrega coleção de cores da legenda
    Set oLegenda = oMovFin.GetLegendas

    ' Define página que será exibida do Recordset
    myRst.AbsolutePage = Pagina
    
    With lstPrincipal
        .Clear                                      ' Limpa conteúdo
        .ColumnCount = 4                            ' Define número de colunas
        .ColumnWidths = "40pt; 180 pt; 55pt; 60pt;" ' Configura largura das colunas
        .Font = "Consolas"                          ' Configura fonte
        
        n = 1
        
        While Not myRst.EOF = True And n <= myRst.PageSize
            
            ' Preenche ListBox
            .AddItem
            
            .List(.ListCount - 1, 0) = myRst.Fields("id").Value
            .List(.ListCount - 1, 1) = myRst.Fields("data").Value
            
            ' Colore a legenda
            
            ' Define o rótulo que receberá a cor
            Set oControle = Controls("l" & Format(n, "00"))
            
            ' Laço para ler cores armazenadas na coleção de legendas da classe
            If oLegenda.Count > 0 Then
            
                For Each vLegenda In oLegenda
                    
                    s() = Split(vLegenda, ";")
                    
'                    If myRst.Fields("genero").Value = s(0) Then
'
'                        oControle.BackColor = s(2): Exit For
'
'                    End If
                    
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
Private Sub Gravar(Decisao As String)

    Dim vbResposta  As VbMsgBoxResult
    
    On Error GoTo err
    
    vbResposta = MsgBox("Deseja realmente fazer a " & Decisao & "?", vbYesNo + vbQuestion, "Pergunta")
    
    If vbResposta = vbYes Then
    
        If Decisao <> "Exclusão" Then
        
            If txbData.Text = Empty Then
                MsgBox "Campo 'Data' é obrigatório", vbCritical: MultiPage1.Value = 1: txbData.SetFocus
            ElseIf txbValor.Text = Empty Or CCur(txbValor.Text) = 0 Then
                MsgBox "Campo 'Valor' não preenchido ou inválido", vbCritical: MultiPage1.Value = 1: txbValor.SetFocus
            ElseIf txbCategoriaID.Text = Empty Then
                MsgBox "Campo 'Categoria' é obrigatório", vbCritical: MultiPage1.Value = 1: txbCategoriaID.SetFocus
            ElseIf cbbMovimento.ListIndex = -1 Then
                MsgBox "Campo 'Movimento' é obrigatório", vbCritical: MultiPage1.Value = 1: cbbMovimento.SetFocus
            ElseIf txbContaID.Text = Empty Then
                MsgBox "Campo 'Conta' é obrigatório", vbCritical: MultiPage1.Value = 1: txbContaID.SetFocus
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
                        .CRUD eCrud.Create
                    Else
                        .CRUD eCrud.Update, .ID
                    End If
                    
                End With
                
                MsgBox Decisao & " realizada com sucesso.", vbInformation, Decisao & " de registro"
                
                Call BuscaRegistros
                                    
            End If
        
        Else ' Se for exclusão
        
            oMovFin.CRUD eCrud.Delete, oMovFin.ID
                
            MsgBox Decisao & " realizada com sucesso.", vbInformation, Decisao & " de registro"
            
            Call BuscaRegistros
            
        End If
               
    ElseIf vbResposta = vbNo Then
    
err:
        If Decisao = "Exclusão" Then
            
            Call btnCancelar_Click
            
        End If
        
    End If
    
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
Private Sub BuscaRegistros(Optional Ordem As String)

    Dim n       As Byte
    Dim o       As control
    Dim sOrdem  As String
    Dim a()     As String

    On Error GoTo err
    
    If Ordem <> "" Then
    
        If oGlobal.Ordem <> "" Then
    
            a() = Split(oGlobal.Ordem, " ")
            
            sOrdem = oGlobal.Ordem
            
            If Ordem = a(0) Then
                
                If a(1) = "ASC" Then
                    Ordem = Ordem & " DESC"
                    oGlobal.Ordem = Ordem
                Else
                    Ordem = Ordem & " ASC"
                    oGlobal.Ordem = Ordem
                End If
            Else
                
                Ordem = Ordem & " ASC"
                oGlobal.Ordem = Ordem
            
            End If
            
        Else
        
            Ordem = Ordem & " ASC"
            oGlobal.Ordem = Ordem
        
        End If
    
    End If
    
    Set myRst = oMovFin.Todos(Ordem, txbFiltro.Text)
    
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
    
err:
    Call btnCancelar_Click
    
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
Private Sub lstPrincipal_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    MultiPage1.Value = 1
    
End Sub

Private Sub lblHdCodigo_Click(): Call BuscaRegistros("id"): End Sub
Private Sub lblHdData_Click(): Call BuscaRegistros("data"): End Sub

Private Sub lblFiltrar_Click()

    oGlobal.Tabela = "tbl_mov_fin" ' Pode ser uma tabela ou consulta
    oGlobal.Filtro = txbFiltro.Text

    f_Filtro.Show

    txbFiltro.Text = oGlobal.Filtro

    Call BuscaRegistros

End Sub
Private Sub MultiPage1_MouseMove(ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Dim oControle As MSForms.control
    
    For Each oControle In Me.Controls
    
        If TypeName(oControle) = "Label" Then
        
            If oControle.Tag = "Header" Then
            
                oControle.ForeColor = &H80000012
            
            ElseIf oControle.Tag = "Filtro" Then
            
                oControle.Font.Bold = False
                oControle.Font.Underline = False
            
            End If
        
        End If
    
    Next

End Sub
Private Sub lblLimpar_Click()
    
    txbFiltro.Text = ""
    
    Call BuscaRegistros

End Sub
Private Sub lblLegenda_Click()
    
    Set oLegenda = New Collection
    
    Set oLegenda = oMovFin.GetLegendas
    
    f_Legenda.Show

End Sub
Private Sub btnData_Click()
    dtDate = IIf(txbData.Text = Empty, Date, txbData.Text)
    txbData.Text = GetCalendario
End Sub
Private Sub btnDataCompra_Click()
    dtDate = IIf(txbDataCompra.Text = Empty, Date, txbDataCompra.Text)
    txbDataCompra.Text = GetCalendario
End Sub
Private Sub btnValor_Click()
    ccurVisor = IIf(txbValor.Text = "", 0, CCur(txbValor.Text))
    txbValor.Text = Format(GetCalculadora, "#,##0.00")
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
        lblDfcGrupo.Caption = oDfc.Grupo
    Else
        txbDfcID.Text = Empty
        lblDfcGrupo.Caption = Empty
    End If
    
End Sub
Private Sub txbDfcID_AfterUpdate()
  
    If IsNumeric(txbDfcID.Text) Then
        
        oDfc.CRUD eCrud.Read, CLng(txbDfcID.Text)
        
        If oDfc.ID = 0 Then
        
            txbDfcID.Text = ""
            lblDfcInfo.Caption = " <DFC não existe ou é subtotal!>"
        
        Else
        
            lblDfcInfo.Caption = " " & oDfc.Grupo
            
        End If

    ElseIf txbDfcID.Text = Empty Then

        txbDfcID.Text = Empty
        lblDfcInfo.Caption = ""
        
    End If

End Sub
Private Sub btnFornecedorID_Click()

    oGlobal.ModoAbrir = eModoAbrirForm.Pesquisa: fFornecedores.Show
    
    Call PesquisaBtn(oFornecedor, Controls("txbFornecedorID"), Controls("lblFornecedor"), Controls("lblFornecedorInfo"))

End Sub
Private Sub txbFornecedorID_AfterUpdate()

    Call PesquisaTxt(Controls("txbFornecedorID"), Controls("lblFornecedor"), Controls("lblFornecedorInfo"), oFornecedor)
        
End Sub
Private Sub btnCategoriaID_Click()

    oGlobal.ModoAbrir = eModoAbrirForm.Pesquisa: fCategorias.Show
    
    Call PesquisaBtn(oCategoria, Controls("txbCategoriaID"), Controls("lblCategoria"), Controls("lblCategoriaInfo"))

End Sub
Private Sub txbCategoriaID_AfterUpdate()

    Call PesquisaTxt(Controls("txbCategoriaID"), Controls("lblCategoria"), Controls("lblCategoriaInfo"), oCategoria)
        
End Sub
Private Sub btnContaID_Click()

    oGlobal.ModoAbrir = eModoAbrirForm.Pesquisa: fContas.Show
    
    Call PesquisaBtn(oConta, Controls("txbContaID"), Controls("lblConta"), Controls("lblContaInfo"))

End Sub
Private Sub txbContaID_AfterUpdate()

    Call PesquisaTxt(Controls("txbContaID"), Controls("lblConta"), Controls("lblContaInfo"), oConta)
        
End Sub
Private Sub btnLojaID_Click()

    oGlobal.ModoAbrir = eModoAbrirForm.Pesquisa: fLojas.Show
    
    Call PesquisaBtn(oLoja, Controls("txbLojaID"), Controls("lblLoja"), Controls("lblLojaInfo"))

End Sub
Private Sub txbLojaID_AfterUpdate()

    Call PesquisaTxt(Controls("txbLojaID"), Controls("lblLoja"), Controls("lblLojaInfo"), oLoja)
        
End Sub
Private Sub PesquisaTxt(TextBoxID As control, LabelTitulo As control, LabelInfo As control, Classe As Object)
    
    If IsNumeric(TextBoxID.Text) Then
        
        Classe.CRUD eCrud.Read, CLng(TextBoxID.Text)
        
        If Classe.ID = 0 Then
        
            TextBoxID.Text = ""
            LabelInfo.Caption = " <" & LabelTitulo & " não existe!>"
        
        Else
        
            LabelInfo.Caption = " " & GetLabelInfo(Classe)
            
        End If
        
    ElseIf TextBoxID.Text = Empty Then

        TextBoxID.Text = Empty
        LabelInfo.Caption = ""
        
    End If
    
End Sub
Private Sub PesquisaBtn(Classe As Object, TextBoxID As control, LabelTitulo As control, LabelInfo As control)
    
    If Not IsNull(oGlobal.PesquisaID) Then

        Classe.CRUD eCrud.Read, oGlobal.PesquisaID
        TextBoxID.Text = Classe.ID: TextBoxID.TextAlign = fmTextAlignRight
        LabelInfo.Caption = " " & GetLabelInfo(Classe)

    Else

        TextBoxID.Text = Empty
        LabelInfo.Caption = Empty

    End If

End Sub
Private Function GetLabelInfo(Classe As Object) As String

    Select Case TypeName(Classe)
        Case "cLoja": GetLabelInfo = Classe.Nome
        Case "cFornecedor": GetLabelInfo = Classe.Nome
        Case "cCategoria"
            GetLabelInfo = " " & Classe.Categoria & " : " & Classe.Subcategoria
            cbbMovimento.Value = Classe.Movimento
            If Not IsNull(Classe.DfcID) Then
                With txbDfcID
                    .Text = Classe.DfcID
                    .TextAlign = fmTextAlignRight
                End With
                oDfc.CRUD eCrud.Read, Classe.DfcID
                lblDfcInfo.Caption = " " & oDfc.Grupo
            End If
        Case "cConta": GetLabelInfo = Classe.Conta
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
