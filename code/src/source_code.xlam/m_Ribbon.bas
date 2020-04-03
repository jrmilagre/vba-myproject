Attribute VB_Name = "m_Ribbon"
Option Private Module

Private MyRibbon    As IRibbonUI
Private sPath       As String
Private sXML        As String
Private oXML        As Object

Public Sub naAcaoBotao(control As IRibbonControl)

    Dim frm As Object

    If Conecta() = True Then
    
        On Error GoTo Erro
        
        If Mid(control.Tag, 1, 1) = "f" Then
        
            oGlobal.ModoAbrir = eModoAbrirForm.Cadastro
            
            Set frm = UserForms.Add(control.Tag)
        
            frm.Show
            
        Else
        
            Application.Run control.Tag
            
        End If
        
        Exit Sub
        
    End If
    
Erro:
    
    MsgBox "Botão ainda não implementado" & err.Description, vbInformation

End Sub
Sub ribbonLoaded(ribbon As IRibbonUI)

    Set MyRibbon = ribbon
    
End Sub
Sub GetModulos(control As IRibbonControl, ByRef returnedVal)
    
    sPath = Mid(wbCode.Path, 1, Len(wbCode.Path) - 5) & _
        Application.PathSeparator & "menus" & _
        Application.PathSeparator & "modulos" & _
        Application.PathSeparator & Environ("username") & ".xml"
    
    Set oXML = CreateObject("Microsoft.XMLDOM")
    
    oXML.Load (sPath)
    
    sXML = oXML.XML
    
    returnedVal = sXML
    
End Sub
Sub GetConfiguracoes(control As IRibbonControl, ByRef returnedVal)
    
    sPath = Mid(wbCode.Path, 1, Len(wbCode.Path) - 5) & _
        Application.PathSeparator & "menus" & _
        Application.PathSeparator & "configuracoes" & _
        Application.PathSeparator & Environ("username") & ".xml"
    
    Set oXML = CreateObject("Microsoft.XMLDOM")
    
    oXML.Load (sPath)
    
    sXML = oXML.XML
    
    returnedVal = sXML
    
End Sub
