VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_Filtro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Classe
Private Type TClass
    Tabela     As String
    Filtro     As String
    Ordem      As String
End Type

Private This As TClass

' Propriedades

Public Property Get Tabela() As String: Tabela = This.Tabela: End Property
Public Property Let Tabela(ByVal Value As String): This.Tabela = Value: End Property

Public Property Get Filtro() As String: Filtro = This.Filtro: End Property
Public Property Let Filtro(ByVal Value As String): This.Filtro = Value: End Property

Public Property Get Ordem() As String: Ordem = This.Ordem: End Property
Public Property Let Ordem(ByVal Value As String): This.Ordem = Value: End Property



