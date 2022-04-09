Attribute VB_Name = "CalcAreaAtual"
'Copyright 2022 FERNANDO A. BORGES

'Permission is hereby granted, free of charge, to any person obtaining a copy of
'this software and associated documentation files (the "Software"), to deal in
'the Software without restriction, including without limitation the rights to
'use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies
'of the Software, and to permit persons to whom the Software is furnished to do
'so, subject to the following conditions:

'The above copyright notice and this permission notice shall be included in all
'copies or substantial portions of the Software.

'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
'INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A
'PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT
'HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF
'CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE
'OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

Option Explicit
Const pi As Double = 3.1415926535
Public Function CalcularAreaAtual() As Double
    Dim area As Double

    
    Select Case TelaInicial.CaixaCombinavelDiametroAtual
        Case "Ø5"
            area = (((0.5 ^ 2 * pi) / 4) * CDbl(TelaInicial.CaixaTextoQuantAtual.Text))
            TelaInicial.RotuloAreaAtual.Caption = CStr(VBA.Round(area, 3))
            CalcularAreaAtual = area
        Case "Ø6,3"
            area = (((0.635 ^ 2 * pi) / 4) * CDbl(TelaInicial.CaixaTextoQuantAtual.Text))
            TelaInicial.RotuloAreaAtual.Caption = CStr(VBA.Round(area, 3))
            CalcularAreaAtual = area
        Case "Ø8"
            area = (((0.8 ^ 2 * pi) / 4) * CDbl(TelaInicial.CaixaTextoQuantAtual.Text))
            TelaInicial.RotuloAreaAtual.Caption = CStr(VBA.Round(area, 3))
            CalcularAreaAtual = area
        Case "Ø10"
            area = (((1 ^ 2 * pi) / 4) * CDbl(TelaInicial.CaixaTextoQuantAtual.Text))
            TelaInicial.RotuloAreaAtual.Caption = CStr(VBA.Round(area, 3))
            CalcularAreaAtual = area
        Case "Ø12,5"
            area = (((1.25 ^ 2 * pi) / 4) * CDbl(TelaInicial.CaixaTextoQuantAtual.Text))
            TelaInicial.RotuloAreaAtual.Caption = CStr(VBA.Round(area, 3))
            CalcularAreaAtual = area
        Case "Ø16"
            area = (((1.6 ^ 2 * pi) / 4) * CDbl(TelaInicial.CaixaTextoQuantAtual.Text))
            TelaInicial.RotuloAreaAtual.Caption = CStr(VBA.Round(area, 3))
            CalcularAreaAtual = area
        Case "Ø20"
            area = (((2# ^ 2 * pi) / 4) * CDbl(TelaInicial.CaixaTextoQuantAtual.Text))
            TelaInicial.RotuloAreaAtual.Caption = CStr(VBA.Round(area, 3))
            CalcularAreaAtual = area
        Case "Ø25"
            area = (((2.5 ^ 2 * pi) / 4) * CDbl(TelaInicial.CaixaTextoQuantAtual.Text))
            TelaInicial.RotuloAreaAtual.Caption = CStr(VBA.Round(area, 3))
            CalcularAreaAtual = area
        Case "Ø32"
            area = (((3.2 ^ 2 * pi) / 4) * CDbl(TelaInicial.CaixaTextoQuantAtual.Text))
            TelaInicial.RotuloAreaAtual.Caption = CStr(VBA.Round(area, 3))
            CalcularAreaAtual = area
    End Select
End Function

