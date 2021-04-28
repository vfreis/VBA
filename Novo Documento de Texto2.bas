' Attribute VB_Name = "Módulo Charge Cluster"
'           Função Faixas Valores/Atrasos
'
'           DATA 14/09/2020
'           '=faixa(tipo de faixa; valor de atraso)'
'           tipo = 'maquina de estado'[1 = faixa de tempo em atraso; 2 = faixa de valor em atraso]
Function charge_cluster(entrada As Integer)

    If (Left(entrada, 3) <= 18) Then
    charge_cluster = "Menor de Idade"
    ElseIf (Left(entrada, 3) <= 25) Then
    charge_cluster = "Entre 18 e 25 Anos"
    ElseIf (Left(entrada, 3) <= 32) Then
    charge_cluster = "Entre 26 e 32 Anos"
    ElseIf (Left(entrada, 3) > 39) Then
    charge_cluster = "Entre 26 e 32 Anos"
    End If


End Function


