Attribute VB_Name = "CalcAreaEquiv"
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

Public Function CalcularAreaEquivalente(ByVal areaAtual As Double)

    Dim area5, area6, area8, area10, area12, area16, area20, area25, area32 As Double
    Dim quant5, quant6, quant8, quant10, quant12, quant16, quant20, quant25, quant32 As Double
      
    'Ø5
    quant5 = ArredParaCima.ArredondarParaCima(areaAtual / ((0.5 ^ 2 * pi) / 4))
    area5 = (((0.5 ^ 2 * pi) / 4) * quant5)
    TelaInicial.Rotulo5QuantEquivalente.Caption = CStr(quant5)
    TelaInicial.Rotulo5AreaEquivalente.Caption = CStr(VBA.Round(area5, 3))
    TelaInicial.Rotulo5AreaEquivalente.BackColor = ClassificarAreaEquivalente(areaAtual, area5)
    
    'Ø6,3
    quant6 = ArredParaCima.ArredondarParaCima(areaAtual / ((0.635 ^ 2 * pi) / 4))
    area6 = (((0.635 ^ 2 * pi) / 4) * quant6)
    TelaInicial.Rotulo6QuantEquivalente.Caption = CStr(quant6)
    TelaInicial.Rotulo6AreaEquivalente.Caption = CStr(VBA.Round(area6, 3))
    TelaInicial.Rotulo6AreaEquivalente.BackColor = ClassificarAreaEquivalente(areaAtual, area6)
    
    'Ø8Teste
    quant8 = ArredParaCima.ArredondarParaCima(areaAtual / ((0.8 ^ 2 * pi) / 4))
    area8 = (((0.8 ^ 2 * pi) / 4) * quant8)
    TelaInicial.Rotulo8QuantEquivalente.Caption = CStr(quant8)
    TelaInicial.Rotulo8AreaEquivalente.Caption = CStr(VBA.Round(area8, 3))
    TelaInicial.Rotulo8AreaEquivalente.BackColor = ClassificarAreaEquivalente(areaAtual, area8)
    
    'Ø10
    quant10 = ArredParaCima.ArredondarParaCima(areaAtual / ((1 ^ 2 * pi) / 4))
    area10 = (((1 ^ 2 * pi) / 4) * quant10)
    TelaInicial.Rotulo10QuantEquivalente.Caption = CStr(quant10)
    TelaInicial.Rotulo10AreaEquivalente.Caption = CStr(VBA.Round(area10, 3))
    TelaInicial.Rotulo10AreaEquivalente.BackColor = ClassificarAreaEquivalente(areaAtual, area10)
    
    'Ø12,5
    quant12 = ArredParaCima.ArredondarParaCima(areaAtual / ((1.25 ^ 2 * pi) / 4))
    area12 = (((1.25 ^ 2 * pi) / 4) * quant12)
    TelaInicial.Rotulo12QuantEquivalente.Caption = CStr(quant12)
    TelaInicial.Rotulo12AreaEquivalente.Caption = CStr(VBA.Round(area12, 3))
    TelaInicial.Rotulo12AreaEquivalente.BackColor = ClassificarAreaEquivalente(areaAtual, area12)
        
    'Ø16
    quant16 = ArredParaCima.ArredondarParaCima(areaAtual / ((1.6 ^ 2 * pi) / 4))
    area16 = (((1.6 ^ 2 * pi) / 4) * quant16)
    TelaInicial.Rotulo16QuantEquivalente.Caption = CStr(quant16)
    TelaInicial.Rotulo16AreaEquivalente.Caption = CStr(VBA.Round(area16, 3))
    TelaInicial.Rotulo16AreaEquivalente.BackColor = ClassificarAreaEquivalente(areaAtual, area16)
    
    'Ø20
    quant20 = ArredParaCima.ArredondarParaCima(areaAtual / ((2 ^ 2 * pi) / 4))
    area20 = (((2 ^ 2 * pi) / 4) * quant20)
    TelaInicial.Rotulo20QuantEquivalente.Caption = CStr(quant20)
    TelaInicial.Rotulo20AreaEquivalente.Caption = CStr(VBA.Round(area20, 3))
    TelaInicial.Rotulo20AreaEquivalente.BackColor = ClassificarAreaEquivalente(areaAtual, area20)
        
    'Ø25
    quant25 = ArredParaCima.ArredondarParaCima(areaAtual / ((2.5 ^ 2 * pi) / 4))
    area25 = (((2.5 ^ 2 * pi) / 4) * quant25)
    TelaInicial.Rotulo25QuantEquivalente.Caption = CStr(quant25)
    TelaInicial.Rotulo25AreaEquivalente.Caption = CStr(VBA.Round(area25, 3))
    TelaInicial.Rotulo25AreaEquivalente.BackColor = ClassificarAreaEquivalente(areaAtual, area25)
    
    'Ø32
    quant32 = ArredParaCima.ArredondarParaCima(areaAtual / ((3.2 ^ 2 * pi) / 4))
    area32 = (((3.2 ^ 2 * pi) / 4) * quant32)
    TelaInicial.Rotulo32QuantEquivalente.Caption = CStr(quant32)
    TelaInicial.Rotulo32AreaEquivalente.Caption = CStr(VBA.Round(area32, 3))
    TelaInicial.Rotulo32AreaEquivalente.BackColor = ClassificarAreaEquivalente(areaAtual, area32)
End Function
