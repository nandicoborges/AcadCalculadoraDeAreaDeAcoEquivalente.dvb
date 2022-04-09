VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TelaInicial 
   Caption         =   "Calcular área de aço equivalente"
   ClientHeight    =   4632
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6780
   OleObjectBlob   =   "TelaInicial.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TelaInicial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CaixaTextoQuantAtual_Change()

        Call CaixaCombinavelDiametroAtual_Change

End Sub

Private Sub CaixaCombinavelDiametroAtual_Change()
    Dim area As Double
    
    If IsNumeric(CaixaTextoQuantAtual.Value) Then
        area = CalcAreaAtual.CalcularAreaAtual
        CalcAreaEquiv.CalcularAreaEquivalente (area)
    Else
        MsgBox "Digite um valor válido para a quantidade de barras atuais."
        CaixaTextoQuantAtual.Value = 0
    End If
End Sub

Private Sub UserForm_Initialize()
    CaixaCombinavelDiametroAtual.AddItem "Ø5"
    CaixaCombinavelDiametroAtual.AddItem "Ø6,3"
    CaixaCombinavelDiametroAtual.AddItem "Ø8"
    CaixaCombinavelDiametroAtual.AddItem "Ø10"
    CaixaCombinavelDiametroAtual.AddItem "Ø12,5"
    CaixaCombinavelDiametroAtual.AddItem "Ø16"
    CaixaCombinavelDiametroAtual.AddItem "Ø20"
    CaixaCombinavelDiametroAtual.AddItem "Ø25"
    CaixaCombinavelDiametroAtual.AddItem "Ø32"
    
    TelaInicial.RotuloAreaAtual.Caption = "0"
End Sub

Private Sub BotaoCancelar_Click()
        Unload Me
End Sub

