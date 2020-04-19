VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_Global"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Classe
Private Type TClass
    Tabela              As String
    Filtro              As String
    Ordem               As String
    ModoAbrir           As eModoAbrirForm
    PesquisaID          As Variant
    Find                As Variant
    AbsolutePosition    As Variant
    AbsolutePage        As Variant
End Type

Private This As TClass

' Propriedades

Public Property Get Tabela() As String: Tabela = This.Tabela: End Property
Public Property Let Tabela(ByVal Value As String): This.Tabela = Value: End Property

Public Property Get Filtro() As String: Filtro = This.Filtro: End Property
Public Property Let Filtro(ByVal Value As String): This.Filtro = Value: End Property

Public Property Get Ordem() As String: Ordem = This.Ordem: End Property
Public Property Let Ordem(ByVal Value As String): This.Ordem = Value: End Property

Public Property Get ModoAbrir() As eModoAbrirForm: ModoAbrir = This.ModoAbrir: End Property
Public Property Let ModoAbrir(ByVal Value As eModoAbrirForm): This.ModoAbrir = Value: End Property

Public Property Get PesquisaID() As Variant: PesquisaID = This.PesquisaID: End Property
Public Property Let PesquisaID(ByVal Value As Variant): This.PesquisaID = Value: End Property

Public Property Get Find() As Variant: Find = This.Find: End Property
Public Property Let Find(ByVal Value As Variant): This.Find = Value: End Property

Public Property Get AbsolutePosition() As Variant: AbsolutePosition = This.AbsolutePosition: End Property
Public Property Let AbsolutePosition(ByVal Value As Variant): This.AbsolutePosition = Value: End Property

Public Property Get AbsolutePage() As Variant: AbsolutePage = This.AbsolutePage: End Property
Public Property Let AbsolutePage(ByVal Value As Variant): This.AbsolutePage = Value: End Property




