Attribute VB_Name = "DICOTOMICO"
'funciones del programa
Public Sub PasoInicial()
'icializar variables
k = 1
eps = 0.01
L = 0.02
a(1) = 0
b(1) = 1

'llama al paso 1
Paso1
End Sub

Public Sub Paso1()
'halla la longitud de la incertidumbre
'LongitudIncertidumbre
If LongitudIncertidumbre <= L Then
    'termina el proceso y muestra el resultado
    MostrarRta
    Exit Sub
Else
    'evaluamos landa y miu
    landa(k) = ((a(k) + b(k)) / 2) - eps
    miu(k) = ((a(k) + b(k)) / 2) + eps
    'llama al paso 2
    Paso2
End If
End Sub
Public Sub Paso2()
'evalua landa y miu en la funcion

flanda(k) = Evaluar(landa(k))
fmiu(k) = Evaluar(miu(k))
If flanda(k) < fmiu(k) Then
    a(k + 1) = a(k)
    b(k + 1) = miu(k)
Else
    a(k + 1) = landa(k)
    b(k + 1) = b(k)
    
End If
k = k + 1
'llama al paso 1
Paso1
End Sub

Public Function LongitudIncertidumbre() As Single
LongitudIncertidumbre = ((1 / (2 ^ k)) * (b(k) - a(k))) + ((2 * eps) * (1 - (1 / (2 ^ k))))
End Function
