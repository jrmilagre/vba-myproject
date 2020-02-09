VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_Database"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Classe
Private Type TClass
    Caminho    As String
    Provedor   As String
End Type

Private This As TClass

' Propriedades
Public Property Get Caminho() As String: Caminho = This.Caminho: End Property
Public Property Let Caminho(ByVal Value As String): This.Caminho = Value: End Property

Public Property Get Provedor() As String: Provedor = This.Provedor: End Property
Public Property Let Provedor(ByVal Value As String): This.Provedor = Value: End Property

' Métodos
Public Sub Informacoes()

    Dim oXML    As Object
    Dim sPath   As String
    Dim o       As Object
    
    sPath = wbCode.Path & Application.PathSeparator & "app_config.xml"
    
    Set oXML = CreateObject("MSXML2.DOMDocument")
 
    With oXML
        .async = False
        .validateOnParse = False
        .Load (sPath)
    End With
    
    Set o = oXML.SelectNodes("//connectionString")
    
    Me.Caminho = o.Item(0).Attributes.Item(2).Value
    Me.Provedor = o.Item(0).Attributes.Item(1).Value
    
End Sub
