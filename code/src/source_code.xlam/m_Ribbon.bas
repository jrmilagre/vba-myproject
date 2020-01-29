Attribute VB_Name = "m_Ribbon"
Option Private Module

Private MyRibbon    As IRibbonUI
Private sPath       As String
Private sXML        As String
Private oXml        As Object

Public Sub naAcaoBotao(control As IRibbonControl)

    If Conecta() = True Then
    
        Select Case control.ID
        
            Case "btnContatos": fContatos.Show
            Case Else: MsgBox "Botão ainda não implementado", vbInformation
            
        End Select
        
    End If

End Sub
Sub ribbonLoaded(ribbon As IRibbonUI)

    Set MyRibbon = ribbon
    
End Sub
Sub GetCadastros(control As IRibbonControl, ByRef returnedVal)
    
    sPath = Mid(wbCode.Path, 1, Len(wbCode.Path) - 5) & _
        Application.PathSeparator & "menus" & _
        Application.PathSeparator & "cadastros" & _
        Application.PathSeparator & Environ("username") & ".xml"
    
    Set oXml = CreateObject("Microsoft.XMLDOM")
    
    oXml.Load (sPath)
    
    sXML = oXml.XML
    
    returnedVal = sXML
    
End Sub
