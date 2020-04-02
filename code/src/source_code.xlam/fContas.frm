VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fContas 
   Caption         =   ":: Cadastro de Contas ::"
   ClientHeight    =   9195
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9960
   OleObjectBlob   =   "fContas.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fContas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private oConta              As New cConta
Private colControles        As New Collection       ' Para eventos de campos
Private myRst               As New ADODB.Recordset
Private bAtualizaScrool     As Boolean

Private Sub UserForm_Initialize()
    
    'Call PopulaCombos
    
    Call Eventos
    
    Call BuscaRegistros

End Sub

Private Sub UserForm_Terminate()
    
    Set oConta = Nothing
    Set myRst = Nothing
    
    Call Desconecta
    
End Sub
Private Sub btnSaldoInicial_Click()
    ccurVisor = IIf(txbSaldoInicial.Text = "", 0, CCur(txbSaldoInicial.Text))
    txbSaldoInicial.Text = Format(GetCalculadora, "#,##0.00")
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
        
        txbConta.SetFocus
        
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
    
        With oConta
    
            .CRUD eCrud.Read, (CLng(lstPrincipal.List(lstPrincipal.ListIndex, 0)))
    
            lblCabID.Caption = IIf(.ID = 0, "", .ID)
            lblCabConta.Caption = .Conta
            lblCabSaldoInicial.Caption = Format(.SaldoInicial, "#,##0.00")
            txbConta.Text = .Conta
            txbSaldoInicial.Text = Format(.SaldoInicial, "#,##0.00")
            
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
        
        txbConta.Enabled = b: lblConta.Enabled = b
        txbSaldoInicial.Enabled = b: lblSaldoInicial.Enabled = b: btnSaldoInicial.Enabled = b
        
    Else
    
        lblCabID.Caption = ""
        lblCabConta.Caption = ""
        lblCabSaldoInicial.Caption = ""
        
        txbConta.Text = Empty
        txbSaldoInicial.Text = IIf(sDecisao = "Inclusão", Format(0, "#,##0.00"), "")
             
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

    ' Define página que será exibida do Recordset
    myRst.AbsolutePage = Pagina
    
    With lstPrincipal
        .Clear                                      ' Limpa conteúdo
        .ColumnCount = 3                            ' Define número de colunas
        .ColumnWidths = "40pt; 180 pt; 75pt;"       ' Configura largura das colunas
        .Font = "Consolas"                          ' Configura fonte
        
        n = 1
        
        While Not myRst.EOF = True And n <= myRst.PageSize
            
            ' Preenche ListBox
            .AddItem
            
            .List(.ListCount - 1, 0) = myRst.Fields("id").Value
            .List(.ListCount - 1, 1) = myRst.Fields("conta").Value
            .List(.ListCount - 1, 2) = Space(15 - Len(Format(myRst.Fields("saldo_inicial").Value, "#,##0.00"))) & Format(myRst.Fields("saldo_inicial").Value, "#,##0.00")
            
            ' Colore a legenda
            
            ' Define o rótulo que receberá a cor
            Set oControle = Controls("l" & Format(n, "00"))
            
            ' Laço para ler cores armazenadas na coleção de legendas da classe
            If oLegenda.Count > 0 Then
            
                For Each vLegenda In oLegenda
                    
'                    s() = Split(vLegenda, ";")
'
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
    Dim e           As eCrud
    
    On Error GoTo err
    
    vbResposta = MsgBox("Deseja realmente fazer a " & Decisao & "?", vbYesNo + vbQuestion, "Pergunta")
    
    If vbResposta = vbYes Then
    
        If Decisao <> "Exclusão" Then
        
            If txbConta.Text = Empty Then
                MsgBox "Campo 'Conta' é obrigatório", vbCritical: MultiPage1.Value = 1: txbConta.SetFocus
            ElseIf txbSaldoInicial.Text = "" Then
                MsgBox "Campo 'Saldo inicial' é obrigatório", vbCritical: MultiPage1.Value = 1: txbSaldoInicial.SetFocus
            Else
                
                With oConta
                    
                    .Conta = txbConta.Text
                    .SaldoInicial = CCur(txbSaldoInicial.Text)
                    
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
        
            oConta.CRUD eCrud.Delete, oConta.ID
                
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

    Dim n       As Byte
    Dim o       As control
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
    
    Set myRst = oConta.Todos(Ordem, txbFiltro.Text)
    
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

End Sub
Private Sub lstPrincipal_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    MultiPage1.Value = 1
    
End Sub

Private Sub lblHdCodigo_Click(): Call BuscaRegistros("id"): End Sub
Private Sub lblHdConta_Click(): Call BuscaRegistros("conta"): End Sub
Private Sub lblHdSaldoInicial_Click(): Call BuscaRegistros("saldo_inicial"): End Sub

Private Sub lblFiltrar_Click()

    oFiltro.Tabela = "vw_contas" ' Pode ser uma tabela ou consulta
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
