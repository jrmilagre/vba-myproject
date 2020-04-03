Attribute VB_Name = "m_Calendario"
Option Private Module
Option Explicit

Public Const sMascaraData   As String = "DD/MM/YYYY"   ' Formata��o de datas
Public dtDate               As Date
Dim Botoes()                As New c_Calendario  ' Vetor que armazena todos os bot�es de dia do Calend�rio

Function GetCalendario() As Date
    
    Dim iTotalBotoes As Integer ' Total de bot�es
    Dim Ctrl As control
    Dim frm As f_Calendario     ' Formul�rio
    
    Set frm = New f_Calendario  ' Cria novo objeto setando formul�rio nele
    
    ' Atribui cada um dos Label num elemento do vetor da classe
    For Each Ctrl In frm.Controls
        If Ctrl.Name Like "l?c?" Then
            iTotalBotoes = iTotalBotoes + 1
            ReDim Preserve Botoes(1 To iTotalBotoes)
            Set Botoes(iTotalBotoes).btnGrupo = Ctrl
        End If
    Next Ctrl
    
    frm.Show
    
    ' Se a data escolhida for nula ou inv�lida, retorna-se a data atual:
    If IsDate(frm.Tag) Then
        GetCalendario = frm.Tag
    Else
        GetCalendario = dtDate
    End If
    
    Unload frm
    
End Function
    

