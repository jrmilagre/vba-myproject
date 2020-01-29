VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fContatos 
   Caption         =   ":: Cadastro de Contatos ::"
   ClientHeight    =   8310
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10050
   OleObjectBlob   =   "fContatos.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fContatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private oContato            As New cContato
Private colControles        As New Collection       ' Para eventos de campos
Private myRst               As New ADODB.Recordset
Private bAtualizaScrool     As Boolean

Private Sub UserForm_Initialize()
    
    Call EventosCampos
    
    Call btnFiltrar_Click

End Sub
Private Sub UserForm_Terminate()
    
    Set oContato = Nothing
    
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
    
    lstPrincipal.ListIndex = -1 ' Tira a seleção
    
End Sub
Private Sub lstPrincipal_Change()
    
    If lstPrincipal.ListIndex >= 0 Then
    
        btnAlterar.Enabled = True
        btnExcluir.Enabled = True
    
        With oContato
    
            .CRUD CRUD.Read, (CLng(lstPrincipal.List(lstPrincipal.ListIndex, 1)))
    
            lblCabID.Caption = IIf(.ID = 0, "", Format(.ID, "000000"))
            lblCabNome.Caption = .Nome
            txbNome.Text = .Nome
            txbNascimento.Text = IIf(IsNull(.Nascimento), "", .Nascimento)
            txbSalario.Text = Format(.Salario, "#,##0.00")
            
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
        txbNascimento.Enabled = b: lblNascimento.Enabled = b: btnNascimento.Enabled = b
        txbSalario.Enabled = b: lblSalario.Enabled = b: btnSalario.Enabled = b
        
    Else
    
        lblCabID.Caption = ""
        lblCabNome.Caption = ""
        txbNome.Text = Empty
        txbNascimento.Text = IIf(sDecisao = "Inclusão", Date, Empty)
        txbSalario.Text = IIf(sDecisao = "Inclusão", Format(0, "#,##0.00"), "")
             
    End If

End Sub
Private Sub lstPrincipalPopular(Pagina As Long)

    Dim n           As Long
    Dim vNascimento As Variant
    Dim vSalario    As Variant

    myRst.AbsolutePage = Pagina
    
    With lstPrincipal
        .Clear                                      ' Limpa conteúdo
        .ColumnCount = 4                            ' Define número de colunas
        .ColumnWidths = "180 pt; 0pt; 55pt; 60pt;"  ' Configura largura das colunas
        .Font = "Consolas"                          ' Configura fonte
        
        n = 1
        
        While Not myRst.EOF = True And n <= myRst.PageSize
            
            .AddItem
            
            .List(.ListCount - 1, 0) = myRst.Fields("nome").Value
            .List(.ListCount - 1, 1) = myRst.Fields("id").Value
            
            If IsNull(myRst.Fields("nascimento").Value) Then vNascimento = "--/--/----" Else vNascimento = myRst.Fields("nascimento").Value
            If IsNull(myRst.Fields("salario").Value) Then vSalario = 0 Else vSalario = myRst.Fields("salario").Value
            
            .List(.ListCount - 1, 2) = vNascimento
            .List(.ListCount - 1, 3) = Space(12 - Len(Format(vSalario, "#,##0.00"))) & Format(vSalario, "#,##0.00")
        
            myRst.MoveNext: n = n + 1
            
        Wend
        
    End With
    
    ' Colore legenda
'    Call ColoreLegenda
    
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
    
    vbResposta = MsgBox("Deseja realmente fazer a " & Decisao & "?", vbYesNo + vbQuestion, "Pergunta")
    
    If vbResposta = vbYes Then
    
        If Decisao <> "Exclusão" Then
        
            If txbNome.Text = Empty Then
                MsgBox "Campo 'Nome' é obrigatório", vbCritical: MultiPage1.Value = 1: txbNome.SetFocus
            Else
                
                With oContato
                    
                    .Nome = txbNome.Text
                    If RTrim(txbNascimento.Text) = "" Then .Nascimento = Null Else .Nascimento = CDate(txbNascimento.Text)
                    If RTrim(txbSalario.Text) = "" Then .Salario = Null Else .Salario = CCur(txbSalario.Text)
                    
                    If Decisao = "Inclusão" Then
                        .CRUD CRUD.Create
                    Else
                        .CRUD CRUD.Update, .ID
                    End If
                    
                End With
                
                MsgBox Decisao & " realizada com sucesso.", vbInformation, Decisao & " de registro"
                
                Call btnFiltrar_Click
                                    
            End If
        
        Else
        
            oContato.CRUD Delete, oContato.ID
                
            MsgBox Decisao & " realizada com sucesso.", vbInformation, Decisao & " de registro"
            
            Call btnFiltrar_Click
            
        End If
               
    ElseIf vbResposta = vbNo Then
        
        If Decisao = "Exclusão" Then
            
            Call btnCancelar_Click
            
        End If
        
    End If
    
End Sub
Private Sub EventosCampos()

    ' Declara variáveis
    Dim oControle   As MSForms.control
    Dim oEvento     As c_EventoCampo
    Dim sTag        As String
    Dim sField()    As String
    
    ' Laço para percorrer todos os TextBox e atribuir eventos
    ' de acordo com o tipo de cada campo
    For Each oControle In Me.Controls
    
        If Len(oControle.Tag) > 0 Then
        
            If TypeName(oControle) = "TextBox" Then
            
                Set oEvento = New c_EventoCampo
                
                With oEvento
                
                    sField() = Split(oControle.Tag, ".")
                    
                    oControle.ControlTipText = cat.Tables(sField(0)).Columns(sField(1)).Properties("Description").Value
                    
                    .FieldType = cat.Tables(sField(0)).Columns(sField(1)).Type
                    .MaxLength = cat.Tables(sField(0)).Columns(sField(1)).DefinedSize
                    .Nullable = cat.Tables(sField(0)).Columns(sField(1)).Properties("Nullable")
                    
                    Set .cGeneric = oControle
                    
                End With
                
                colControles.Add oEvento
                
            End If
            
        End If
    Next

End Sub
Private Sub btnFiltrar_Click()

    Set myRst = oContato.Todos()
    
    If myRst.PageCount > 0 Then
        
        bAtualizaScrool = False
        
        With scrPagina
            .Max = myRst.PageCount
            .Value = myRst.PageCount
        End With
        
        Call lstPrincipalPopular(myRst.PageCount)
        
    Else
    
        lstPrincipal.Clear
        
'        Call ColoreLegenda
        
    End If
    
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
Private Sub ColoreLegenda()

'    Dim idx         As Integer
'    Dim c           As control
'
'    For Each c In fTitulosReceber.Controls
'
'        If TypeName(c) = "Label" And c.Tag = "status" Then
'
'            idx = CInt(Mid(c.name, 2, 2))
'
'            If idx <= (lstPrincipal.ListCount - 1) Then
'                If CDate(lstPrincipal.List(idx, 2)) > (Date + 3) Then
'                    c.BackColor = &HC000& ' Verde
'                ElseIf CDate(lstPrincipal.List(idx, 2)) < Date Then
'                    c.BackColor = &HC0& ' Vermelho
'                Else
'                    c.BackColor = &HFFFF&         ' Amarelo
'                End If
'            Else
'                c.BackColor = &H8000000F
'            End If
'
'        End If
'
'    Next c
    
End Sub
