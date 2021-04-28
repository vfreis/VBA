Attribute VB_Name = "AtualizarFormulas"
' Attribute VB_Name = "Módulo11"
Sub atualizar()

Sheets("BASE").Select

Range("J1").Value = ("CARTEIRA")
Range("K1").Value = ("STATUA NEGÓCIO")
Range("L1").Value = ("STATUS")
Range("M1").Value = ("DATA")
Range("N1").Value = ("HORA")

'----------------------------------------------------------COLAR FORMULAS-----------------------------------------------------------------------

Range("J2").FormulaLocal = ("=$H2")
Range("K2").FormulaLocal = ("=PROCV($E2;$O:$P;2;0)")
Range("L2").FormulaLocal = ("=$e2")
Range("M2").FormulaLocal = ("=TEXTO($B2;""dd/mm/aaaa"")")
Range("N2").FormulaLocal = ("=TEXTO($c2;""hh"")")

Sheets("BASE").Select
    Range("A1").Select
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False

Sheets("BASE").Select
LIN = Range("A1048576").End(xlUp).Row
[J2:N2].Copy
Range("J2:N" & LIN & "").PasteSpecial xlPasteFormulas

Range("Tabela__10.130.115.47_OCSH_Historic[valor]").Select
    Selection.TextToColumns Destination:=Range("G2"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
        
Sheets("HxH PROMESSAS").Select
Calculate

End Sub