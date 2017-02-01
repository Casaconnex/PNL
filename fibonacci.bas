Attribute VB_Name = "Fibonacci"
Public Sub pasoinfibona()
   L = 0.02
   eps = 0.01
   a(1) = 0
   b(1) = 1
   k = 1
   n = calcularN
   landa(k) = a(k) + ((fibonacci(n - 2)) / fibonacci(n)) * (b(k) - a(k))
   miu(k) = a(k) + ((fibonacci(n - 1)) / fibonacci(n)) * (b(k) - a(k))
   flanda(k) = Evaluar(landa(k))
   fmiu(k) = Evaluar(miu(k))
   pasoprincipaF
End Sub
Public Sub pasoprincipaF()
   paso1FI
End Sub
   
Public Sub paso1FI()
   flanda(k) = Evaluar(landa(k))
   fmiu(k) = Evaluar(miu(k))
   If flanda(k) > fmiu(k) Then
      paso2FI
   Else
      paso3FI
   End If
   
End Sub
Public Sub paso2FI()
   a(k + 1) = landa(k)
   b(k + 1) = b(k)
   landa(k + 1) = miu(k)
   miu(k + 1) = a(k + 1) + (fibonacci(n - k - 1) / fibonacci(n - k)) * (b(k + 1) - a(k + 1))
   If k = n - 2 Then
      paso5FI
   Else
      
      fmiu(k + 1) = Evaluar(miu(k + 1))
       paso4FI
   End If
End Sub
Public Sub paso3FI()

 a(k + 1) = a(k)
   b(k + 1) = miu(k)
   miu(k + 1) = landa(k)
   landa(k + 1) = a(k + 1) + (fibonacci(n - k - 2) / fibonacci(n - k)) * (b(k + 1) - a(k + 1))
   If n - k = 2 Then
      paso5FI
   Else
      flanda(k + 1) = Evaluar(landa(k + 1))
      paso4FI
   End If

End Sub
Public Sub paso4FI()
   k = k + 1
   paso1FI
End Sub
Public Sub paso5FI()
   landa(n) = landa(n - 1)
   miu(n) = landa(n - 1) + eps
   flanda(n) = Evaluar(landa(n))
   fmiu(n) = Evaluar(miu(n))
   If Evaluar(landa(n)) > Evaluar(miu(n)) Then
      a(n) = landa(n)
      b(n) = b(n - 1)
   Else
      a(n) = a(n - 1)
      b(n) = landa(n)
   End If
   MostrarRta
   Exit Sub
End Sub
Public Function fibonacci(n As Integer) As Integer
   If n = 0 Then
      fibonacci = 1
   End If
   If n > 0 Then
      fibonacci = fibonacci(n - 1) + fibonacci(n - 2)
   End If
End Function

Public Function calcularN() As Integer
   Dim i As Integer
   i = 1
   
   While fibonacci(i) < ((b(1) - a(1)) / 0.02)
      i = i + 1
   Wend
   calcularN = i
End Function
