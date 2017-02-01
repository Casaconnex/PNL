Attribute VB_Name = "DORADO"
Public ALPHA As Single
Public Sub PASOINDORADO()
    L = 0.02
    a(1) = 0
    b(1) = 1
    eps = 0.01
    k = 1
    ALPHA = 0.618
    landa(k) = a(k) + ((1 - ALPHA) * (b(k) - a(k)))
    miu(k) = a(k) + (ALPHA * (b(k) - a(k)))
    'EVALUAR LANDA Y MIU
    flanda(k) = Evaluar(landa(k))
    fmiu(k) = Evaluar(miu(k))
    'IR AL PASO PRINCIPAL
    PASOPRINCIPALDORADO
End Sub

Public Sub PASOPRINCIPALDORADO()
    PASO1DOR
    
End Sub

Public Sub PASO1DOR()
    If (b(k) - a(k)) < L Then
        MostrarRta
        Exit Sub
    End If
    If flanda(k) > fmiu(k) Then
        PASO2DOR
    ElseIf flanda(k) <= fmiu(k) Then
        PASO3DOR
    End If
End Sub

Public Sub PASO2DOR()
    a(k + 1) = landa(k)
    b(k + 1) = b(k)
    landa(k + 1) = miu(k)
    miu(k + 1) = a(k + 1) + (ALPHA * (b(k + 1) - a(k + 1)))
    flanda(k + 1) = fmiu(k)
    fmiu(k + 1) = Evaluar(miu(k + 1))
    PASO4DOR
End Sub

Public Sub PASO3DOR()
    a(k + 1) = a(k)
    b(k + 1) = miu(k)
    landa(k + 1) = a(k + 1) + ((1 - ALPHA) * (b(k + 1) - a(k + 1)))
    miu(k + 1) = landa(k)
    fmiu(k + 1) = flanda(k)
    flanda(k + 1) = Evaluar(landa(k + 1))
    PASO4DOR
End Sub

Public Sub PASO4DOR()
    k = k + 1
    PASO1DOR
End Sub
