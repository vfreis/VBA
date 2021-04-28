Attribute VB_Name = "Módulo1"
Sub EnviarEmail()

    'Dim OutApp As Outlook.Application
    'Dim OutMail As Outlook.MailItem
    'Dim wrdEdit
    'Dim Assinatura

    'Ctrl+Q Enviar
    
    With Application
        .EnableEvents = False
        .ScreenUpdating = False
    End With

    Select Case Hour(Now)
        Case 1 To 11: saudaçao = "Manhã"
        Case 12 To 17: saudaçao = "Tarde"
        Case 18 To 23: saudaçao = "Noite"
    End Select
    
    'tempo = Format(Now(), "hh")
    

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    OutMail.Display
    Assinatura = OutMail.HTMLBody
    sCorpoMail = "<p>Caros, <p> Segue acompanhamento diário da operação de Marisa e Portocred:"
 With OutMail
        .To = "planejamento.poa@zanc.com.br; dainara.eisenmann@zanc.com.br; luthiane.machado@zanc.com.br; gisele.garcia@zanc.com.br;"
        .CC = "'estrategia.digital@zanc.com.br'"
        .Subject = "Nucleo Digital | Acompanhamento Diário | Marisa & Portocred | " & saudaçao & " - " & Format(Now, "dd/mm")
        .HTMLBody = "<HTML><BODY><FONT FACE=Verdana SIZE=2" & sCorpoMail & "</FONT></BODY></HTML>" & Assinatura
        .Display
    End With
    On Error GoTo 0


    With Application
        .EnableEvents = True
        .ScreenUpdating = True
    End With
    
    On Error GoTo tentarDenovo
tentarDenovo:
    
    Set wrdEdit = OutApp.ActiveInspector.WordEditor
    OutApp.Application.ActiveInspector.Activate
    
    wrdEdit.Application.Selection.Paragraphs(1).Range.Select
    wrdEdit.Application.Selection.EndOf
    wrdEdit.Application.Selection.Paragraphs(1).Range.Select
    wrdEdit.Application.Selection.EndOf
    
    'Pula linha
    wrdEdit.Application.Selection.InsertBreak (2)
    
    Sheets("Produção Diária").Activate

    Calculate
        Range("B5:J47").CopyPicture xlScreen, xlBitmap
        wrdEdit.Application.Selection.Paste

    
    'Range("A1:Q37").Copy
    'wrdEdit.Application.Selection.PasteSpecial
    
    End Sub
