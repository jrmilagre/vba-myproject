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
' BeginGroupVB
Sub Main()
    
    On Error GoTo GroupXError
  
    If m_Database.Conecta = True Then
    
        Dim usrNew As ADOX.User
        Dim usrLoop As ADOX.User
        Dim grpLoop As ADOX.Group
  
        With cat
        
            ' Cria e acrescenda novo grupo com uma string.
            .Groups.Append "Accounting"
      
            ' Cria e acrescenta um novo usuário com um objeto.
            Set usrNew = New ADOX.User
            usrNew.Name = "Pat Smith"
            usrNew.ChangePassword "", "Password1"
            .Users.Append usrNew
      
            ' Make the user Pat Smith a member of the
            ' Accounting group by creating and adding the
            ' appropriate Group object to the user's Groups
            ' collection. The same is accomplished if a User
            ' object representing Pat Smith is created and
            ' appended to the Accounting group Users collection
            usrNew.Groups.Append "Accounting"
      
            ' Enumerate all User objects in the
            ' catalog's Users collection.
            For Each usrLoop In .Users
                Debug.Print "  " & usrLoop.Name
                Debug.Print "    Belongs to these groups:"
                ' Enumerate all Group objects in each User
                ' object's Groups collection.
                If usrLoop.Groups.Count <> 0 Then
                    For Each grpLoop In usrLoop.Groups
                        Debug.Print "    " & grpLoop.Name
                    Next grpLoop
                Else
                    Debug.Print "    [None]"
                End If
            Next usrLoop
      
            ' Enumerate all Group objects in the default
            ' workspace's Groups collection.
            For Each grpLoop In .Groups
                Debug.Print "  " & grpLoop.Name
                Debug.Print "    Has as its members:"
                ' Enumerate all User objects in each Group
                ' object's Users collection.
                If grpLoop.Users.Count <> 0 Then
                    For Each usrLoop In grpLoop.Users
                        Debug.Print "    " & usrLoop.Name
                    Next usrLoop
                Else
                    Debug.Print "    [None]"
                End If
            Next grpLoop
      
            ' Delete new User and Group objects because this
            ' is only a demonstration.
            ' These two line are commented out because the sub "OwnersX" uses
            ' the group "Accounting".
    '        .Users.Delete "Pat Smith"
    '        .Groups.Delete "Accounting"
      
        End With
        
    End If
  
    'Clean up
    Call Desconecta
    Set usrNew = Nothing
    Exit Sub
  
GroupXError:
  
    Call Desconecta
    Set usrNew = Nothing
  
    If err <> 0 Then
        MsgBox err.Source & "-->" & err.Description, , "Error"
    End If
End Sub
' EndGroupVB
