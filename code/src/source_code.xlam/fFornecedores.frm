VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fFornecedores 
   Caption         =   ":: Cadastro de Fornecedores ::"
   ClientHeight    =   9105
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9960
   OleObjectBlob   =   "fFornecedores.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fFornecedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private oFornecedor         As New cFornecedor
Private colControles        As New Collection       ' Para eventos de campos
Private myRst               As New ADODB.Recordset
Private bAtualizaScrool     As Boolean

Private Sub UserForm_Initialize()

    Call PopulaCombos
    
    Call Eventos
    
    Call BuscaRegistros

End Sub

Private Sub UserForm_Terminate()
    
    Set oFornecedor = Nothing
    Set myRst = Nothing
    
    If oGlobal.ModoAbrir = Cadastro Then
        
        oGlobal.Find = Null
        Call Desconecta
        
    Else
    
        If lstPrincipal.ListIndex = -1 Then
            oGlobal.PesquisaID = Null
        Else
            oGlobal.PesquisaID = CLng(lstPrincipal.List(lstPrincipal.ListIndex, 0))
        End If
        
        oGlobal.ModoAbrir = eModoAbrirForm.Cadastro
        
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
    
    MultiPage1.Value = 1
    
    If Decisao <> "Exclusão" Then
    
        If Decisao = "Inclusão" Then
        
            Call Campos("Limpar")
            
        End If
        
        Call Campos("Habilitar")
        
        txbNome.SetFocus
        
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
    
    If oGlobal.ModoAbrir = eModoAbrirForm.Cadastro Then
        lstPrincipal.ListIndex = -1 ' Tira a seleção
    Else
        lstPrincipal.SetFocus
    End If
    
End Sub
Private Sub lstPrincipal_Change()

    Dim n As Long
    
    If lstPrincipal.ListIndex >= 0 Then
    
        btnAlterar.Enabled = True
        btnExcluir.Enabled = True
    
        With oFornecedor
    
            .CRUD eCrud.Read, (CLng(lstPrincipal.List(lstPrincipal.ListIndex, 0)))
    
            lblCabID.Caption = IIf(.ID = 0, "", .ID)
            lblCabNome.Caption = .Nome
            txbNome.Text = .Nome
            
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
        
        txbNome.Enabled = b: lblNome.Enabled = b
        
    Else
    
        lblCabID.Caption = ""
        lblCabNome.Caption = ""
        txbNome.Text = Empty
             
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
    Set oLegenda = oFornecedor.GetLegendas

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
            .List(.ListCount - 1, 1) = myRst.Fields("nome").Value
            
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
    
    ' Posiciona no registro invocado, caso exista
    If Not IsNull(oGlobal.Find) Then
        lstPrincipal.ListIndex = oGlobal.AbsolutePosition - (((Pagina - 1) * myRst.PageSize) + 1)
        oGlobal.Find = Null
    End If
    
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
    
    On Error GoTo err
    
    vbResposta = MsgBox("Deseja realmente fazer a " & Decisao & "?", vbYesNo + vbQuestion, "Pergunta")
    
    If vbResposta = vbYes Then
    
        If Decisao <> "Exclusão" Then
        
            If txbNome.Text = Empty Then
                MsgBox "Campo 'Nome' é obrigatório", vbCritical: MultiPage1.Value = 1: txbNome.SetFocus
            Else
                
                With oFornecedor
                    
                    .Nome = txbNome.Text
                    
                    If Decisao = "Inclusão" Then
                        .CRUD eCrud.Create, , Decisao
                    Else
                        .CRUD eCrud.Update, .ID, Decisao
                    End If
                    
                End With
                
                Call BuscaRegistros
                                    
            End If
        
        Else ' Se for exclusão
        
            oFornecedor.CRUD eCrud.Delete, oFornecedor.ID, Decisao
            
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
    
    Set myRst = oFornecedor.Todos(Ordem, txbFiltro.Text)
    
    If Not myRst.EOF = True Then
        
        bAtualizaScrool = False
        
        If Not IsNull(oGlobal.Find) Then
        
            myRst.MoveFirst
            myRst.Find "id= " & oGlobal.Find, , adSearchForward
            
            oGlobal.AbsolutePosition = CLng(myRst.AbsolutePosition)
            
        Else
        
            myRst.MoveLast
            
        End If
        
        scrPagina.Max = myRst.PageCount
        scrPagina.Value = myRst.AbsolutePage
            
        Call lstPrincipalPopular(myRst.AbsolutePage)
        
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

End Sub
Private Sub lstPrincipal_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    If oGlobal.ModoAbrir = eModoAbrirForm.Cadastro Then
        
        MultiPage1.Value = 1
        
    Else
    
        If lstPrincipal.ListIndex = -1 Then
            oGlobal.PesquisaID = Null
        Else
            oGlobal.PesquisaID = CLng(lstPrincipal.List(lstPrincipal.ListIndex, 0))
        End If
    
        Unload Me
    
    End If
    
End Sub

Private Sub lblHdCodigo_Click(): Call BuscaRegistros("id"): End Sub
Private Sub lblHdNome_Click(): Call BuscaRegistros("nome"): End Sub

Private Sub lblFiltrar_Click()

    oGlobal.Tabela = "tbl_fornecedores" ' Pode ser uma tabela ou consulta
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
    
    Set oLegenda = oFornecedor.GetLegendas
    
    f_Legenda.Show

End Sub
Private Sub lstPrincipal_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    If KeyCode = 13 Then Call lstPrincipal_DblClick(Nothing)
    
End Sub
