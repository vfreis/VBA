Attribute VB_Name = "Módulo1"
Function textoMoeda(entrada As String)

Dim charac As Integer
Dim virgula As String
Dim restante As Integer
Dim dps As String


virgula = ","
charac = Len(entrada)
restante = charac - 3
dps = Left(entrada, restante)
 
moedita = dps & virgula & Right(entrada, 2)

End Function

