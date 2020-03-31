VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} f_Filtro 
   Caption         =   ":: Filtro ::"
   ClientHeight    =   3360
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9255
   OleObjectBlob   =   "f_Filtro.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "f_Filtro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private colEventos        As New Collection       ' Para eventos de campos
Private adxTbl            As New ADOX.Table
Private sTbl              As String

Private Sub UserForm_Initialize()
    
    Call cbbCamposPopular

    txbFiltro.Text = oFiltro.Filtro

End Sub
Private Sub cbbCamposPopular()

    Dim col     As ADOX.Column
    Dim c()     As String
    
    sTbl = oFiltro.Tabela
    
    Set adxTbl = cat.Tables(sTbl)

    With cbbCampo
        .Clear
        .ColumnCount = 4
        .ColumnWidths = "120pt; 0pt; 0pt; 0pt;"
        
        ' COLUNAS
        ' 1) - Nome da coluna no banco de dados
        ' 2) - Apelido da coluna
        ' 3) - Tipo da coluna
        ' 4) - Tabela de origem
        
        If adxTbl.Type = "TABLE" Then
            
            For Each col In adxTbl.Columns
        
                .AddItem
            
                .List(.ListCount - 1, 0) = col.Properties.Item(2)
                .List(.ListCount - 1, 1) = col.Name
                .List(.ListCount - 1, 2) = col.Type
                .List(.ListCount - 1, 3) = adxTbl.Name
                
            Next
            
        ElseIf adxTbl.Type = "VIEW" Then
            
            For Each col In adxTbl.Columns

                .AddItem

                c() = Split(col.Name, "-")

                .List(.ListCount - 1, 0) = cat.Tables(c(0)).Columns(c(1)).Properties.Item(2)
                .List(.ListCount - 1, 1) = cat.Tables(c(0)).Columns(c(1)).Name
                .List(.ListCount - 1, 2) = cat.Tables(c(0)).Columns(c(1)).Type
                .List(.ListCount - 1, 3) = c(0)

            Next
            
        End If
        
        cbbCampo.ListIndex = 0
        
    End With

End Sub
Private Sub cbbOperadorPopular()

    With cbbOperador
        .Clear
        .ColumnCount = 2
        .ColumnWidths = "100pt; 0pt;"
        .AddItem: .List(.ListCount - 1, 0) = "Igual a": .List(.ListCount - 1, 1) = "="
        '.AddItem: .List(.ListCount - 1, 0) = "Diferente de": .List(.ListCount - 1, 1) = "<>"
        .AddItem: .List(.ListCount - 1, 0) = "Menor que": .List(.ListCount - 1, 1) = "<"
        .AddItem: .List(.ListCount - 1, 0) = "Menor ou igual a": .List(.ListCount - 1, 1) = "<="
        .AddItem: .List(.ListCount - 1, 0) = "Maior que": .List(.ListCount - 1, 1) = ">"
        .AddItem: .List(.ListCount - 1, 0) = "Maior ou igual a": .List(.ListCount - 1, 1) = ">="
        
        If cbbCampo.List(cbbCampo.ListIndex, 2) = "202" Then
            .AddItem: .List(.ListCount - 1, 0) = "Começa com": .List(.ListCount - 1, 1) = ""
            .AddItem: .List(.ListCount - 1, 0) = "Contém a expressão": .List(.ListCount - 1, 1) = ""
            .AddItem: .List(.ListCount - 1, 0) = "Não contém a expressão": .List(.ListCount - 1, 1) = ""
        End If
        
        cbbOperador.ListIndex = 0
        
    End With
    
'- Que contém a expressão
'- Que não contém
'- Diferente de vazio
'- Vazio
'- Começa com
'- Termina com

End Sub
Private Sub btnAplicar_Click()
    oFiltro.Filtro = txbFiltro.Text
    Unload Me
End Sub
Private Sub btnAdicionar_Click()

    Dim sCampo              As String
    Dim sTipoCampo          As String
    Dim sOperador           As String
    Dim sExpressao          As String
    Dim sExpressaoAnterior  As String
    
    sExpressaoAnterior = txbFiltro.Text
    sCampo = cbbCampo.List(cbbCampo.ListIndex, 1)
    sTipoCampo = cbbCampo.List(cbbCampo.ListIndex, 2)
    sOperador = cbbOperador.List(cbbOperador.ListIndex, 0)
    sExpressao = txbExpressao.Text
    
    If sExpressaoAnterior = "" Then
        sExpressao = TrataExpressao(sCampo, sTipoCampo, sOperador, sExpressao) 'sInstrucao '& " " & oFiltro.TrataExpressao(sCampo, cbbCampo.List(cbbCampo.ListIndex, 2), sOperador, txbExpressao.Text)
    Else
        sExpressao = sExpressaoAnterior & " " & TrataExpressao(sCampo, sTipoCampo, sOperador, sExpressao) 'sInstrucao '& " " & oFiltro.TrataExpressao(sCampo, cbbCampo.List(cbbCampo.ListIndex, 2), sOperador, txbExpressao.Text)
    End If
    
    txbFiltro.Text = sExpressao

End Sub
Private Sub btnLimpar_Click()

    txbFiltro.Text = Empty

End Sub
Private Sub btnParentesesAbre_Click()
    
    txbFiltro.Text = txbFiltro.Text & " ("

End Sub
Private Sub btnParentesesFecha_Click()
    
    txbFiltro.Text = txbFiltro.Text & " )"

End Sub
Private Sub btnAND_Click()
    
    txbFiltro.Text = txbFiltro.Text & " AND"

End Sub
Private Sub btnOR_Click()
    
    txbFiltro.Text = txbFiltro.Text & " OR"

End Sub
Private Sub cbbCampo_AfterUpdate()
    
    Call cbbOperadorPopular
    
    Call ConfiguraCampoExpressao

End Sub
Private Sub ConfiguraCampoExpressao()

    Dim oControle   As MSForms.control
    Dim oEvento     As c_Evento
    Dim sCampo      As String
    Dim c(0 To 1)   As String
    
    sCampo = cbbCampo.List(cbbCampo.ListIndex, 1)
    
    For Each oControle In Me.Controls
    
        If TypeName(oControle) = "TextBox" And oControle.Name = "txbExpressao" Then
        
            Set oEvento = New c_Evento
            
            With oEvento
            
                If adxTbl.Type = "TABLE" Then
                    
                    .FieldType = cat.Tables(sTbl).Columns(sCampo).Type
                    .MaxLength = cat.Tables(sTbl).Columns(sCampo).DefinedSize
                    .Nullable = cat.Tables(sTbl).Columns(sCampo).Properties("Nullable")
                    
                ElseIf adxTbl.Type = "VIEW" Then
                
                    c(0) = cbbCampo.List(cbbCampo.ListIndex, 1)
                    c(1) = cbbCampo.List(cbbCampo.ListIndex, 3)
                
                    .FieldType = cat.Tables(c(1)).Columns(c(0)).Type
                    .MaxLength = cat.Tables(c(1)).Columns(c(0)).DefinedSize
                    .Nullable = cat.Tables(c(1)).Columns(c(0)).Properties("Nullable")
                
                End If
                    
                Set .cTextBox = oControle
                
                If colEventos.Count > 0 Then: colEventos.Remove 1
                
                colEventos.Add oEvento
                    
            End With
            
        End If
        
    Next
    
End Sub
Private Function TrataExpressao(Campo As String, TipoCampo As String, Operador As String, Expressao As String) As String

    ' Trata expressão
    If Operador = "Igual a" Then
        If TipoCampo = "202" Then
            Expressao = "= '" & Expressao & "'"
        ElseIf TipoCampo = "7" Then
            Expressao = "= #" & Expressao & "#"
        ElseIf TipoCampo = "6" Then
            Expressao = "= " & Replace(Replace(Expressao, ".", ""), ",", ".")
        Else
            Expressao = "= " & Expressao
        End If
    ElseIf Operador = "Menor ou igual a" Then
        If TipoCampo = "7" Then
            Expressao = "<= #" & Expressao & "#"
        ElseIf TipoCampo = "6" Then
            Expressao = "<= " & Replace(Replace(Expressao, ".", ""), ",", ".")
        Else
            Expressao = "<= " & Expressao
        End If
    ElseIf Operador = "Menor que" Then
        If TipoCampo = "7" Then
            Expressao = "< #" & Expressao & "#"
        ElseIf TipoCampo = "6" Then
            Expressao = "< " & Replace(Replace(Expressao, ".", ""), ",", ".")
        Else
            Expressao = "< " & Expressao
        End If
    ElseIf Operador = "Maior ou igual a" Then
        If TipoCampo = "7" Then
            Expressao = ">= #" & Expressao & "#"
        ElseIf TipoCampo = "6" Then
            Expressao = ">= " & Replace(Replace(Expressao, ".", ""), ",", ".")
        Else
            Expressao = ">= " & Expressao
        End If
    ElseIf Operador = "Maior que" Then
        If TipoCampo = "7" Then
            Expressao = "> #" & Expressao & "#"
        ElseIf TipoCampo = "6" Then
            Expressao = "> " & Replace(Replace(Expressao, ".", ""), ",", ".")
        Else
            Expressao = "> " & Expressao
        End If
    ElseIf Operador = "Começa com" Then
        Expressao = "LIKE '" & Expressao & "%'"
    ElseIf Operador = "Contém a expressão" Then
        Expressao = "LIKE '%" & Expressao & "%'"
    ElseIf Operador = "Não contém a expressão" Then
        Expressao = "NOT LIKE '%" & Expressao & "%'"
    End If
    
    TrataExpressao = Campo & " " & Expressao

End Function
Private Sub btnCalendario_Click()
    dtDate = IIf(txbExpressao.Text = Empty, Date, txbExpressao.Text)
    txbExpressao.Text = GetCalendario
End Sub
Private Sub btnCalculadora_Click()
    txbExpressao.Text = Format(GetCalculadora, "#,##0.00")
End Sub
Private Sub cbbCampo_Change()

    If cbbCampo.List(cbbCampo.ListIndex, 2) = "6" Then
        btnCalculadora.Visible = True
    ElseIf cbbCampo.List(cbbCampo.ListIndex, 2) = "7" Then
        btnCalendario.Visible = True
    Else
        btnCalculadora.Visible = False
        btnCalendario.Visible = False
    End If
    
End Sub
