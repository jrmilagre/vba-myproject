VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fCategorias 
   Caption         =   ":: Cadastro de Categorias ::"
   ClientHeight    =   9105
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11655
   OleObjectBlob   =   "fCategorias.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fCategorias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private oCategoria          As New cCategoria
Private colControles        As New Collection           ' Para atribuir eventos aos campos
Private myRst               As New ADODB.Recordset
Private bAtualizaScrool     As Boolean
Private Sub UserForm_Initialize()

    Call PopulaCombos
    
    Call Eventos
    
    Call BuscaRegistros

End Sub
Private Sub UserForm_Terminate()
    
    Set oCategoria = Nothing
    Set myRst = Nothing
    
    Call Desconecta
    
End Sub
Private Sub btnSalario_Click()
    ccurVisor = IIf(txbSalario.Text = "", 0, CCur(txbSalario.Text))
    txbSalario.Text = Format(GetCalculadora, "#,##0.00")
End Sub
Private Sub btnNascimento_Click()
    dtDate = IIf(txbNascimento.Text = Empty, Date, txbNascimento.Text)
    txbNascimento.Text = GetCalendario
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
        
        txbCategoria.SetFocus
        
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

    Dim n           As Long
    Dim oControl    As control
    
    If lstPrincipal.ListIndex >= 0 Then
    
        btnAlterar.Enabled = True
        btnExcluir.Enabled = True
    
        With oCategoria
    
            .CRUD eCrud.Read, (CLng(lstPrincipal.List(lstPrincipal.ListIndex, 0)))
    
            lblCabID.Caption = IIf(.ID = 0, "", Format(.ID, "000000"))
            lblCabCategoria.Caption = .Categoria
            lblCabSubcategoria.Caption = .Subcategoria
            lblCabGrupo.Caption = IIf(.Grupo = "R", "RECEITA", "DESPESA")
            txbCategoria.Text = .Categoria
            txbSubcategoria.Text = .Subcategoria
            
            For n = 0 To cbbGrupo.ListCount - 1
                If .Grupo = cbbGrupo.List(n, 0) Then
                    cbbGrupo.ListIndex = n
                    Exit For
                End If
            Next n
            
            For n = 1 To 10
            
                Set oControl = Controls("opt" & Format(n, "00"))
                
                If oControl.Tag = .DfcID Then oControl.Value = True
            
            Next n
            
        End With
        
    End If

End Sub
Private Sub Campos(Acao As String)
    
    Dim sDecisao    As String
    Dim b           As Boolean
    Dim oControl    As control
    Dim n           As Integer
    
    sDecisao = Replace(btnConfirmar.Caption, "Confirmar ", "")
    
    If Acao <> "Limpar" Then
    
        If Acao = "Desabilitar" Then
            b = False
        ElseIf Acao = "Habilitar" Then
            b = True
        End If
        
        MultiPage1.Pages(0).Enabled = Not b
        
        txbCategoria.Enabled = b: lblCategoria.Enabled = b
        txbSubcategoria.Enabled = b: lblSubcategoria.Enabled = b
        cbbGrupo.Enabled = b: lblGrupo.Enabled = b
        btnLimparSelecao.Enabled = b
        lblOperacional.Enabled = b
        lblTatico.Enabled = b
        lblEstrategico.Enabled = b
        
        For n = 1 To 10
            
            Set oControl = Controls("opt" & Format(n, "00")): oControl.Enabled = b
            
        Next n
        
        frmDFC.Enabled = b
        
    Else
    
        lblCabID.Caption = ""
        lblCabCategoria.Caption = ""
        txbCategoria.Text = Empty
        txbSubcategoria.Text = Empty
        cbbGrupo.ListIndex = -1
        
        For n = 1 To 10
            
            Set oControl = Controls("opt" & Format(n, "00")): oControl.Value = False
            
        Next n
             
    End If

End Sub
Private Sub lstPrincipalPopular(Pagina As Long)

    Dim n           As Byte
    Dim oLegenda     As control
    Dim vDFC        As Variant
    
    ' Limpa cores da legenda
    For n = 1 To myRst.PageSize
        Set oLegenda = Controls("l" & Format(n, "00")): oLegenda.BackColor = &H8000000F
    Next n

    ' Define página que será exibida do Recordset
    myRst.AbsolutePage = Pagina
    
    With lstPrincipal
        .Clear                                              ' Limpa conteúdo
        .ColumnCount = 5                                    ' Define número de colunas
        .ColumnWidths = "40 pt; 152pt; 152pt; 55pt; 60pt;"  ' Configura largura das colunas
        .Font = "Consolas"                                  ' Configura fonte
        
        n = 1
        
        While Not myRst.EOF = True And n <= myRst.PageSize
            
            ' Preenche ListBox
            .AddItem
            
            .List(.ListCount - 1, 0) = myRst.Fields("id").Value
            .List(.ListCount - 1, 1) = myRst.Fields("categoria").Value
            .List(.ListCount - 1, 2) = myRst.Fields("subcategoria").Value
            .List(.ListCount - 1, 3) = IIf(myRst.Fields("grupo").Value = "R", "RECEITA", "DESPESA")
            
            If IsNull(myRst.Fields("dfc_id").Value) Then
                vDFC = "<não-atribuído>"
            Else
                vDFC = oCategoria.GetDFC(myRst.Fields("dfc_id").Value)
            End If
            
            .List(.ListCount - 1, 4) = vDFC
            
            'If IsNull(myRst.Fields("nascimento").Value) Then vNascimento = "--/--/----" Else vNascimento = myRst.Fields("nascimento").Value
            'If IsNull(myRst.Fields("salario").Value) Then vSalario = 0 Else vSalario = myRst.Fields("salario").Value
            
            '.List(.ListCount - 1, 2) = vNascimento
            '.List(.ListCount - 1, 3) = Space(12 - Len(Format(vSalario, "#,##0.00"))) & Format(vSalario, "#,##0.00")
            
            ' Colore a legenda
            Set oLegenda = Controls("l" & Format(n, "00"))
            
            If myRst.Fields("grupo").Value = "D" Then
                oLegenda.BackColor = &HFF& 'Vermelho
            ElseIf myRst.Fields("grupo").Value = "R" Then
                oLegenda.BackColor = &HFF0000 ' Azul
            Else
                oLegenda.BackColor = &H8000000F 'Cinza (sem cor)
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
    Dim e           As eCrud
    Dim n           As Integer
    Dim oControl    As control
    Dim optButton   As Boolean
    
    vbResposta = MsgBox("Deseja realmente fazer a " & Decisao & "?", vbYesNo + vbQuestion, "Pergunta")
    
    If vbResposta = vbYes Then
    
        If Decisao <> "Exclusão" Then
        
            If txbCategoria.Text = Empty Then
                MsgBox "Campo 'Categoria' é obrigatório", vbCritical: MultiPage1.Value = 1: txbCategoria.SetFocus
            ElseIf txbSubcategoria.Text = Empty Then
                MsgBox "Campo 'Subcategoria' é obrigatório", vbCritical: MultiPage1.Value = 1: txbSubcategoria.SetFocus
            ElseIf cbbGrupo.ListIndex = -1 Then
                MsgBox "Campo 'Grupo' é obrigatório", vbCritical: MultiPage1.Value = 1: txbSubcategoria.SetFocus
            Else
            
                optButton = False
                
                For n = 1 To 10
                
                    Set oControl = Controls("opt" & Format(n, "00"))
                    
                    If oControl.Value = True Then
                        
                        optButton = True
                        
                        oCategoria.DfcID = oControl.Tag
                        
                        Exit For
                        
                    End If
                        
                Next n
                
                If optButton = False Then
                
                    oCategoria.DfcID = Null
                    
                End If
                
                With oCategoria
                
                    .Categoria = txbCategoria.Text
                    .Subcategoria = txbSubcategoria.Text
                    .Grupo = cbbGrupo.List(cbbGrupo.ListIndex, 0)
                
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
        
            oCategoria.CRUD eCrud.Delete, oCategoria.ID
                
            MsgBox Decisao & " realizada com sucesso.", vbInformation, Decisao & " de registro"
            
            Call BuscaRegistros
            
        End If
               
    ElseIf vbResposta = vbNo Then
        
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
                
                Set oEvento = New c_Evento
                
                Set oEvento.cLabel = oControle
                
                colControles.Add oEvento
                
                If oControle.Tag = "CAB" Then
                
                    sCor() = Split(oConfig.GetCorInfoCab, " ")
                    oControle.ForeColor = RGB(CInt(sCor(0)), CInt(sCor(1)), CInt(sCor(2)))
                
                End If
                
            End If
                
        End If
        
    Next

End Sub
Private Sub BuscaRegistros(Optional Ordem As String)

    Dim n As Byte
    Dim o As control
    Dim sOrdem  As String
    Dim a()     As String

    On Error GoTo Erro
    
    If Ordem <> "" Then
    
        If oFiltro.Ordem <> "" Then
    
            a() = Split(oFiltro.Ordem, " ")
            
            sOrdem = oFiltro.Ordem
            
            If Ordem = a(0) Then
                
                If a(1) = "ASC" Then
                    Ordem = Ordem & " DESC"
                    oFiltro.Ordem = Ordem
                Else
                    Ordem = Ordem & " ASC"
                    oFiltro.Ordem = Ordem
                End If
            Else
                
                Ordem = Ordem & " ASC"
                oFiltro.Ordem = Ordem
            
            End If
            
        Else
        
            Ordem = Ordem & " ASC"
            oFiltro.Ordem = Ordem
        
        End If
    
    End If

    Set myRst = oCategoria.Todos(Ordem, txbFiltro.Text)
    
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

    ' Carrega combo Grupos
    Dim col As Collection
    Dim n   As Variant
    Dim s() As String
    
    Set col = oCategoria.GetGrupos
    
    With cbbGrupo
        .Clear
        .ColumnCount = 2
        .ColumnWidths = "15pt;30pt;"
        
        For Each n In col
        
            s() = Split(n, ",")
            
                .AddItem
                .List(.ListCount - 1, 0) = s(1)
                .List(.ListCount - 1, 1) = s(0)
        Next n
    
    End With
    
End Sub
Private Sub lstPrincipal_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    MultiPage1.Value = 1
    
End Sub
Private Sub lblHdNome_Click()

    Call BuscaRegistros("nome")
    
End Sub
Private Sub lblFiltrar_Click()

    oFiltro.Tabela = "tbl_categorias" ' Pode ser uma tabela ou consulta
    oFiltro.Filtro = txbFiltro.Text

    f_Filtro.Show

    txbFiltro.Text = oFiltro.Filtro

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
Private Sub btnLimparSelecao_Click()

    Dim n           As Integer
    Dim oControl    As control

    For n = 1 To 10
    
        Set oControl = Controls("opt" & Format(n, "00"))
        
        oControl.Value = False
            
    Next n

End Sub
Private Sub lblHdCodigo_Click(): Call BuscaRegistros("id"): End Sub
Private Sub lblHdCategoria_Click(): Call BuscaRegistros("categoria"): End Sub
Private Sub lblHdSubcategoria_Click(): Call BuscaRegistros("subcategoria"): End Sub
Private Sub lblHdGrupo_Click(): Call BuscaRegistros("grupo"): End Sub
