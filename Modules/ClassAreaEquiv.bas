Attribute VB_Name = "ClassAreaEquiv"
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

Public Function ClassificarAreaEquivalente(ByVal areaAtual, ByVal areaEquivalente As Double)
    Dim verdeClaro, amareloClaro, vermelhoClaro As Variant
            
    verdeClaro = RGB(192, 255, 192)
    amareloClaro = RGB(255, 255, 192)
    vermelhoClaro = RGB(255, 128, 128)
    
    If areaEquivalente <= (1.1 * areaAtual) Then
        ClassificarAreaEquivalente = verdeClaro
    ElseIf areaEquivalente > (1.1 * areaAtual) And areaEquivalente <= (1.5 * areaAtual) Then
        ClassificarAreaEquivalente = amareloClaro
    ElseIf areaEquivalente > (1.5 * areaAtual) Then
        ClassificarAreaEquivalente = vermelhoClaro
    End If
    
End Function
