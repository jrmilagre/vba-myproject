VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_Evento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents cTextBox      As MSForms.TextBox
Attribute cTextBox.VB_VarHelpID = -1
Public WithEvents cLabel        As MSForms.Label
Attribute cLabel.VB_VarHelpID = -1

Private Type TProperty
    FieldType   As Integer
    Nullable    As Boolean
    MaxLength   As Integer
    Tag         As String
    Name        As String
End Type

Private This As TProperty

Public Property Get FieldType() As Integer: FieldType = This.FieldType: End Property
Public Property Let FieldType(ByVal Value As Integer): This.FieldType = Value: End Property

Public Property Get Nullable() As Boolean: Nullable = This.Nullable: End Property
Public Property Let Nullable(ByVal Value As Boolean): This.Nullable = Value: End Property

Public Property Get MaxLength() As Integer: MaxLength = This.MaxLength: End Property
Public Property Let MaxLength(ByVal Value As Integer): This.MaxLength = Value: End Property

Public Property Get Tag() As String: Tag = This.Tag: End Property
Public Property Let Tag(ByVal Value As String): This.Tag = Value: End Property

Public Property Get Name() As String: Name = This.Name: End Property
Public Property Let Name(ByVal Value As String): This.Name = Value: End Property
'Public Function Evento(Controle As MSForms.control, Tabela As String) As Variant
'
'    If TypeName(Controle) = "TextBox" Then
'
'        With Me
'
'            Controle.ControlTipText = cat.Tables(Tabela).Columns(Controle.Tag).Properties("Description").Value
'
'            .FieldType = cat.Tables(Tabela).Columns(Controle.Tag).Type
'
'            If .FieldType = 6 Then
'                Controle.TextAlign = fmTextAlignRight
'            End If
'
'            .MaxLength = cat.Tables(Tabela).Columns(Controle.Tag).DefinedSize
'            .Nullable = cat.Tables(Tabela).Columns(Controle.Tag).Properties("Nullable")
'
'            Set .cTextBox = Controle
'
'        End With
'
'    End If
'
'    Set Evento = Me
'
'End Function

Private Sub cTextBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
       
        
    If KeyAscii = 13 Then: SendKeys "{TAB}"   ' Emula o TAB
    
    If This.FieldType = 3 Then ' Long
    
        Select Case KeyAscii
        
            Case 48 To 57               ' N�meros de 0 a 9
            Case Else
                KeyAscii = 0     ' Ignora os outros caracteres
                MsgBox "S� � permitido n�mero.", vbInformation
        End Select
        
        cTextBox.TextAlign = fmTextAlignRight
    
    ElseIf This.FieldType = 6 Then ' Currency
    
        Select Case KeyAscii
            Case 8                      ' Backspace (seta de apagar)
            Case 48 To 57               ' N�meros de 0 a 9
            Case 44                     ' V�rgula
            
            If InStr(cTextBox, ",") Then 'Se o campo j� tiver v�rgula ent�o ele n�o adiciona
                KeyAscii = 0 'N�o adiciona a v�rgula caso ja tenha
            Else
                KeyAscii = 44 'Adiciona uma v�rgula
            End If
            
            ' Para aceitar n�mero negativo
            Case 45

            If cTextBox.SelStart > 1 Then
                KeyAscii = 0
            End If
            
            Case Else
                KeyAscii = 0 'N�o deixa nenhuma outra caractere ser escrito
                MsgBox "S� � permitido n�mero.", vbInformation
        End Select
    
    ElseIf This.FieldType = 7 Then ' Date
    
        With cTextBox
            .MaxLength = 10
            
            Select Case KeyAscii
                Case 8                      ' Aceita o BACK SPACE
                Case 48 To 57
                    If .SelStart = 2 Then .SelText = "/"
                    If .SelStart = 5 Then .SelText = "/"
                Case Else: KeyAscii = 0     ' Ignora os outros caracteres
            End Select
        End With
    ElseIf This.FieldType = 202 Then
    
        cTextBox.TextAlign = fmTextAlignLeft
        
    End If
    
End Sub
Private Sub cTextBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    If This.FieldType = 6 Then
        Select Case KeyCode
            Case 9
                cTextBox.Text = Format(cTextBox.Text, "#,##0.00")
                cTextBox.TextAlign = fmTextAlignRight
            Case 13
                cTextBox.Text = Format(cTextBox.Text, "#,##0.00")
                cTextBox.TextAlign = fmTextAlignRight
                SendKeys "{TAB}"   ' Emula o TAB
            Case 115
                ccurVisor = IIf(cTextBox.Text = "", 0, CCur(cTextBox.Text))
                cTextBox.Text = Format(GetCalculadora, "#,##0.00")
        End Select
    ElseIf This.FieldType = 7 Then
        ' Se a tecla {F4} for pressionada, abre o formul�rio calend�rio
        Select Case KeyCode
            Case 115
                dtDate = IIf(cTextBox.Text = "", Date, cTextBox.Text)
                cTextBox.Text = GetCalendario
        End Select
    End If
End Sub

Private Sub cTextBox_Enter()
    If This.FieldType = 6 Then
        ' Seleciona todos os caracteres do campo
        cTextBox.SelStart = 0
        cTextBox.SelLength = Len(cTextBox.Text)
    End If
End Sub
Private Sub cTextBox_Exit()
    If This.FieldType = 6 Then
        cTextBox.TextAlign = fmTextAlignLeft
        cTextBox.Text = Format(cTextBox.Text, "#,##0.00")
    End If
End Sub
'Private Sub cCpf_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'
'    cCpf.MaxLength = 14
'
'    With cCpf
'
'    Select Case KeyAscii
'            Case 8                      ' Aceita o BACK SPACE
'            Case 13: SendKeys "{TAB}"   ' Emula o TAB
'            Case 48 To 57
'                If .SelStart = 3 Then .SelText = "."
'                If .SelStart = 7 Then .SelText = "."
'                If .SelStart = 11 Then .SelText = "-"
'            Case Else: KeyAscii = 0     ' Ignora os outros caracteres
'        End Select
'    End With
'
'End Sub
Private Sub cLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    If cLabel.Tag = "Filtro" Then
    
        With cLabel.Font
            .Underline = True
            .Bold = True
        End With
        
    ElseIf cLabel.Tag = "Header" Then
    
        cLabel.ForeColor = &H80000011
    
    End If

End Sub

