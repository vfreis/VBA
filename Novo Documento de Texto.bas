' Attribute VB_Name = "Módulo Charge Cluster"
'           Função Faixas Valores/Atrasos
'
'           DATA 14/09/2020
'           '=faixa(tipo de faixa; valor de atraso)'
'           tipo = 'maquina de estado'[1 = faixa de tempo em atraso; 2 = faixa de valor em atraso]
Function charge_cluster(entrada As Integer)

    If (Left(entrada, 3) <= 180) Then
    charge_cluster = "PDD 1"
    ElseIf (Left(entrada, 3) <= 360) Then
    charge_cluster = "PDD 2"
    ElseIf (Left(entrada, 3) <= 720) Then
    charge_cluster = "WO 1"
    ElseIf (Left(entrada, 3) > 720) Then
    charge_cluster = "WO 2"
    End If


End Function


