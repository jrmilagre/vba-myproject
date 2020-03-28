Attribute VB_Name = "m_Ribbon"
Option Private Module

Private MyRibbon    As IRibbonUI
Private sPath       As String
Private sXML        As String
Private oXML        As Object

Public Sub naAcaoBotao(control As IRibbonControl)

    Dim frm As Object

    If Conecta() = True Then
    
        On Error GoTo err
        
        Set frm = UserForms.Add(control.Tag)
        frm.Show
        Exit Sub
        
    End If
    
err:
    
    MsgBox "Botão ainda não implementado", vbInformation

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
