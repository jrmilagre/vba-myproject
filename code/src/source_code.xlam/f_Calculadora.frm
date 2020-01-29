VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} f_Calculadora 
   Caption         =   ":: Calculadora"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3570
   OleObjectBlob   =   "f_Calculadora.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "f_Calculadora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim oCalculadora As New c_Calculadora

Dim dMemoria    As Double
Dim iOperacao   As Integer
Enum Operacao
    Adicao = 1
    Subtracao = 2
    Multiplicacao = 3
    Divisao = 4
End Enum

Private Sub btn0_Click()
    txbVisor.Text = IIf(txbVisor.Text = "0", "0", txbVisor.Text + "0")
End Sub
Private Sub btn1_Click()
    txbVisor.Text = IIf(txbVisor.Text = "0", "1", txbVisor.Text + "1")
End Sub
Private Sub btn2_Click()
    txbVisor.Text = IIf(txbVisor.Text = "0", "2", txbVisor.Text + "2")
End Sub
Private Sub btn3_Click()
    txbVisor.Text = IIf(txbVisor.Text = "0", "3", txbVisor.Text + "3")
End Sub
Private Sub btn4_Click()
    txbVisor.Text = IIf(txbVisor.Text = "0", "4", txbVisor.Text + "4")
End Sub
Private Sub btn5_Click()
    txbVisor.Text = IIf(txbVisor.Text = "0", "5", txbVisor.Text + "5")
End Sub
Private Sub btn6_Click()
    txbVisor.Text = IIf(txbVisor.Text = "0", "6", txbVisor.Text + "6")
End Sub
Private Sub btn7_Click()
    txbVisor.Text = IIf(txbVisor.Text = "0", "7", txbVisor.Text + "7")
End Sub
Private Sub btn8_Click()
    txbVisor.Text = IIf(txbVisor.Text = "0", "8", txbVisor.Text + "8")
End Sub
Private Sub btn9_Click()
    txbVisor.Text = IIf(txbVisor.Text = "0", "9", txbVisor.Text + "9")
End Sub
Private Sub btnSomar_Click()
    dMemoria = CDbl(txbVisor.Text)
    txbVisor.Text = 0
    iOperacao = Operacao.Adicao
End Sub
Private Sub btnSubtrair_Click()
    dMemoria = CDbl(txbVisor.Text)
    txbVisor.Text = 0
    iOperacao = Operacao.Subtracao
End Sub
Private Sub btnMultiplicar_Click()
    dMemoria = CDbl(txbVisor.Text)
    txbVisor.Text = 0
    iOperacao = Operacao.Multiplicacao
End Sub
Private Sub btnDividir_Click()
    dMemoria = CDbl(txbVisor.Text)
    txbVisor.Text = 0
    iOperacao = Operacao.Divisao
End Sub

Private Sub btnResultado_Click()
    Select Case iOperacao
        Case Operacao.Adicao
            txbVisor.Text = dMemoria + txbVisor.Text
        Case Operacao.Subtracao
            txbVisor.Text = dMemoria - txbVisor.Text
        Case Operacao.Multiplicacao
            txbVisor.Text = dMemoria * txbVisor.Text
        Case Operacao.Divisao
            txbVisor.Text = dMemoria / txbVisor.Text
        Case Else
            MsgBox "Dígito inválido!"
        End Select
End Sub

Private Sub btnUsar_Click()
    oCalculadora.Resultado = CDbl(txbVisor.Text)
End Sub

Private Sub btnVirgula_Click()
    txbVisor.Text = txbVisor.Text + ","
End Sub

Private Sub txbVisor_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' Permite apenas números
    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

Private Sub UserForm_Initialize()
    txbVisor.Text = IIf(ccurVisor > 0, Format(ccurVisor, "#,##0.00"), 0)
End Sub
