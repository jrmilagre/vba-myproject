Attribute VB_Name = "m_Testes"
Sub LerXML()
    
    Dim sPath As String
    Dim sXML As String
    Dim oXML As Variant
    Dim n As Variant
    
    sPath = "C:\Users\jfonseca\Desktop\teste.xml"
    
    Set oXML = CreateObject("Microsoft.XMLDOM")
    
    oXML.Load (sPath)
    
For Each n In oXML.ChildNodes
    If n.BaseName = "ingrediente" Then
        Debug.Print n.Text
    End If
Next
    
    sXML = oXML.XML
    
    Debug.Print sXML
    
End Sub
Sub LerXML2()

    Dim xmlObj As Object
    Dim sPath As String
    
    sPath = wbCode.Path & _
        Application.PathSeparator & "app_config.xml"
    
    Set xmlObj = CreateObject("MSXML2.DOMDocument")
 
    xmlObj.async = False
    xmlObj.validateOnParse = False
    xmlObj.Load (sPath)
 
    Dim nodesThatMatter As Object
    Dim node            As Object
    Set nodesThatMatter = xmlObj.SelectNodes("//connectionStrings")
    
    Dim level1 As Object
    Dim level2 As Object
    Dim level3 As Object
    
    For Each level1 In nodesThatMatter
        For Each level2 In level1.ChildNodes
            
            If level2.BaseName = "connectionStrings" Then
                
                'Debug.Print level2.ChildNodes.Item(3).Attributes.getNamedItem("name").Text
                
                For Each level3 In level2.ChildNodes '.Item(3).ChildNodes
                
                    If level3.BaseName = "add" Then
                        Debug.Print level3.Attributes.Item(0).Name & "=" & level3.Attributes.Item(0).Value
                        Debug.Print level3.Attributes.Item(1).Name & "=" & level3.Attributes.Item(1).Value
                        Debug.Print level3.Attributes.Item(2).Name & "=" & level3.Attributes.Item(2).Value
                    End If
                  
                Next level3
                
            End If
            

        Next
    Next
    
End Sub
Sub UpdateXML()
    
    Dim sPath As String
    Dim o As Object
    Dim oXML As Object

    sPath = "C:\Users\jfonseca\Desktop\teste.xml"
    
    Set oXML = CreateObject("MSXML2.DOMDocument")
    ' Set oXML = CreateObject("Microsoft.XMLDOM")
    
    With oXML
        .async = False
        .validateOnParse = False
        .Load (sPath)
    End With

    ' Selecionando um único nó
    Set o = oXML.SelectSingleNode("//configuration/connectionString/@caminho")
    
    o.Value = "C:\temp\ws-vba\_project_model\data\banco.mdb"
    
    oXML.Save (sPath)

End Sub

