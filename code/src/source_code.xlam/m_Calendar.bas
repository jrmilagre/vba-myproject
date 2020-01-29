Attribute VB_Name = "m_Calendar"
Option Private Module
Option Explicit

Public Const sMascaraData   As String = "DD/MM/YYYY"   ' Máscara de formatação de datas
Public dtDate               As Date
Dim aBotoes()               As New c_Calendario  ' Vetor que armazena todos os botões de dia do Calendário

Public Function GetCalendario() As Date
    ' Função GetCalendario
    
    ' Declara variáveis
    Dim lTotalBotoes As Long   ' Total de rótulos
    Dim Ctrl As control
    Dim frm As New f_Calendario      ' Formulário
    
    'Set frm = New fCalendario ' Cria novo objeto setando formulário nele
    
    ' Atribui cada um dos botões em um elemento do vetor da classe
    For Each Ctrl In frm.Controls
        If Ctrl.name Like "l?c?" Then
            lTotalBotoes = lTotalBotoes + 1
            ReDim Preserve aBotoes(1 To lTotalBotoes)
            Set aBotoes(lTotalBotoes).btnGrupo = Ctrl
        End If
    Next Ctrl
    
    frm.Show
    
    ' Se a data escolhida for nula ou inválida, retorna-se a data atual:
    If IsDate(frm.Tag) Then
        GetCalendario = frm.Tag
    Else
        GetCalendario = dtDate
    End If
    
    Unload frm
End Function
    

