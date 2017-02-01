Attribute VB_Name = "GENERAL"
'declaración de variables
Public k As Integer
Public a(1 To 100) As Single
Public b(1 To 100) As Single
Public L As Single
Public eps As Single
Public landa(1 To 100) As Single
Public miu(1 To 100) As Single
Public flanda(1 To 100) As Double
Public fmiu(1 To 100) As Double
Public Longitud As Single
Public i As Integer
Public rta As Variant
Public n As Integer

Public Sub MostrarRta()
'mostramos la tabla resultante
For i = 1 To k
    Form1.List1.AddItem i
    Form1.List2.AddItem a(i)
    Form1.List3.AddItem b(i)
    Form1.List4.AddItem landa(i)
    Form1.List5.AddItem miu(i)
    Form1.List6.AddItem flanda(i)
    Form1.List7.AddItem fmiu(i)
Next i
'mostramos el resultado
Form1.Label3(0).Caption = "b(" & k & ")" & "-a(" & k & ")"
Form1.Label6.Caption = Format((b(k) - a(k)) / 2, "0.####")

Form1.Label7.Caption = "a(" & k & ") + "
Form1.Label3(1).Caption = "b(" & k & ")" & "-a(" & k & ")"
Form1.Label8.Caption = Format(a(k) + ((b(k) - a(k)) / 2), "0.####")

Form1.Label10.Caption = "b(" & k & ") - "
Form1.Label11.Caption = "b(" & k & ")" & "-a(" & k & ")"
Form1.Label13.Caption = Format(b(k) - ((b(k) - a(k)) / 2), "0.####")

MsgBox "a= " & a(k) & vbCrLf & "b= " & b(k) & vbCrLf & "Iteración= " & k, vbInformation, "Solución"
rta = MsgBox("Deseas intentarlo de nuevo?", vbYesNo + vbQuestion, "Método Dicotómico")
If rta = vbYes Then
    Form1.List1.Clear
    Form1.List2.Clear
    Form1.List3.Clear
    Form1.List4.Clear
    Form1.List5.Clear
    Form1.List6.Clear
    Form1.List7.Clear
    Form1.Label6.Caption = ""
    Form1.Label8.Caption = ""
    Form1.Label13.Caption = ""
End If

End Sub


Public Function Evaluar(X As Single) As Single
    Evaluar = 1 - (4 * X) + (3 * (X ^ 2))
End Function

