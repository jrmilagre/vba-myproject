VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_Config"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Classe
Private Type TClass
    CaminhoXML      As String
End Type

Private This As TClass
Private oXML As Object

Public Property Get CaminhoXML() As String
    CaminhoXML = wbCode.Path & Application.PathSeparator & "app_config.xml"
End Property

Public Function GetCaminhoBD() As String
    
    Set oXML = CreateObject("MSXML2.DOMDocument")
 
    With oXML
        .async = False
        .validateOnParse = False
        .Load (Me.CaminhoXML)
    End With

    GetCaminhoBD = oXML.SelectSingleNode("//configuration/connectionString/@caminho").Value

End Function
Public Function GetProvedorDB() As String
    
    Set oXML = CreateObject("MSXML2.DOMDocument")
 
    With oXML
        .async = False
        .validateOnParse = False
        .Load (Me.CaminhoXML)
    End With

    GetProvedorDB = oXML.SelectSingleNode("//configuration/connectionString/@provedor").Value

End Function
Public Function GetCorInfoCab() As String

    Set oXML = CreateObject("MSXML2.DOMDocument")
 
    With oXML
        .async = False
        .validateOnParse = False
        .Load (Me.CaminhoXML)
    End With

    GetCorInfoCab = oXML.SelectSingleNode("//configuration/colors/@fonteCab").Value

End Function


