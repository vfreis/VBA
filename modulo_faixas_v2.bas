Attribute VB_Name = "Módulo1"
Function faixa_atraso(entrada As Integer)

If (entrada < 180) Then
faixa_atraso = "61 e 180"
ElseIf (entrada <= 360) Then
faixa_atraso = "181 a 360"
ElseIf (entrada > 360) Then
faixa_atraso = "Acima 360"
End If

End Function
