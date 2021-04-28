Attribute VB_Name = "Módulo2"
Sub auto_update()
Attribute auto_update.VB_ProcData.VB_Invoke_Func = " \n14"

Dim data As String
Dim agora As String
Dim inicio As Double

'Agendamento de Atualizações'

Application.OnTime "08:05:00", "auto_update"
Application.OnTime "09:05:00", "auto_update"
Application.OnTime "10:05:00", "auto_update"
Application.OnTime "11:05:00", "auto_update"
Application.OnTime "12:05:00", "auto_update"
Application.OnTime "13:05:00", "auto_update"
Application.OnTime "14:05:00", "auto_update"
Application.OnTime "15:05:00", "auto_update"
Application.OnTime "16:05:00", "auto_update"
Application.OnTime "17:05:00", "auto_update"
Application.OnTime "18:05:00", "auto_update"
Application.OnTime "19:05:00", "auto_update"
Application.OnTime "20:05:00", "auto_update"
Application.OnTime "21:05:00", "auto_update"



'Desliga gráfico
Application.ScreenUpdating = True
'Alertas
Application.DisplayAlerts = False

'
' Limpa sheet
'
    Rows("2:2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("A1").Select

agora = Right(Int(Now), 4) & "-" & Mid(Now, 4, 2) & "-" & Left(Now, 2)

'
' Atualiza discagens VoxAge
'
'
    
    Workbooks.Open Filename:= _
        "\\10.230.215.148\Relatorios\VOXAGE_Export_Discagem_Hora__" & agora & ".csv"
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Windows("AUTO_UPDATE.xlsm").Activate
    Range("A1").Select
    ActiveSheet.Paste
    Workbooks("VOXAGE_Export_Discagem_Hora__" & agora & ".csv").Close SaveChanges:=False
    ActiveWorkbook.Save
    
End Sub


    
