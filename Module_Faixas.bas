Attribute VB_Name = "Módulo1"
' Attribute VB_Name = "Módulo11"
'           Função Faixas Valores/Atrasos
'
'           DATA 14/09/2020
'           '=faixa(tipo de faixa; valor de atraso)'
'           tipo = 'maquina de estado'[1 = faixa de tempo em atraso; 2 = faixa de valor em atraso]
'


Function faixas(tipo As Integer, faixa As Integer)


If (tipo = 1) Then

    'Faixa de Tempo

    If (faixa <= 5) Then
    faixas = "Não Tem"
    ElseIf (faixa >= 5 And faixa <= 30) Then
    faixas = "5 a 30"
    ElseIf (faixa >= 31 And faixa <= 60) Then
    faixas = "31 a 60"
    ElseIf (faixa >= 61 And faixa <= 90) Then
    faixas = "61 a 90"
    ElseIf (faixa >= 91 And faixa <= 120) Then
    faixas = "91 a 120"
    ElseIf (faixa >= 121 And faixa <= 150) Then
    faixas = "121 a 150"
    ElseIf (faixa >= 151 And faixa <= 180) Then
    faixas = "151 a 180"
    ElseIf (faixa >= 181 And faixa <= 210) Then
    faixas = "181 a 210"
    ElseIf (faixa >= 211 And faixa <= 240) Then
    faixas = "211 a 240"
    ElseIf (faixa >= 241 And faixa <= 270) Then
    faixas = "241 a 270"
    ElseIf (faixa >= 271 And faixa <= 300) Then
    faixas = "271 a 300"
    ElseIf (faixa >= 301 And faixa <= 330) Then
    faixas = "301 a 330"
    ElseIf (faixa >= 331 And faixa <= 360) Then
    faixas = "331 a 360"
    ElseIf (faixa >= 361 And faixa <= 540) Then
    faixas = "361 a 540"
    ElseIf (faixa >= 541 And faixa <= 720) Then
    faixas = "541 a 720"
    ElseIf (faixa >= 721 And faixa <= 1080) Then
    faixas = "721 a 1080"
    ElseIf (faixa >= 1081 And faixa <= 1440) Then
    faixas = "1081 a 1440"
    ElseIf (faixa >= 1441 And faixa <= 1800) Then
    faixas = "1441 a 1800"
    ElseIf (faixa >= 1800) Then
    faixas = "Acima de 1800"
    End If
End If

    'Faixa de valor
    
If (tipo = 2) Then


    If (faixa >= 25 And faixa <= 100) Then
    faixas = "Entre R$25 e R$100"
    ElseIf (faixa >= 101 And faixa <= 250) Then
    faixas = "Entre R$101 e R$250"
    ElseIf (faixa >= 251 And faixa <= 500) Then
    faixas = "Entre R$251 e R$500"
    ElseIf (faixa >= 501 And faixa <= 750) Then
    faixas = "Entre R$501 e R$750"
    ElseIf (faixa >= 751 And faixa <= 1000) Then
    faixas = "Entre R$751 e R$1.000"
    ElseIf (faixa >= 1001 And faixa <= 2500) Then
    faixas = "Entre R$1.001 e R$2.500"
    ElseIf (faixa >= 2501 And faixa <= 5000) Then
    faixas = "Entre R$2.501 e R$5.000"
    ElseIf (faixa >= 5001 And faixa <= 10000) Then
    faixas = "Entre R$5.001 e R$10.000"
    ElseIf (faixa >= 10001 And faixa <= 25000) Then
    faixas = "Entre R$10.001 e R$25.000"
    ElseIf (faixa >= 25001 And faixa <= 50000) Then
    faixas = "Entre R$25.001 e R$50.000"
    ElseIf (faixa > 50000) Then faixas = faixas = "Acima  de R$ 50.000 "
End If
    
    If (tipo < 1) Or (tipo > 2) Then
    faixas = "sem função definida"
    End If
End If

 
    
    
    
    
    
End Function



