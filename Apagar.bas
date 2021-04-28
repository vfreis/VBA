Attribute VB_Name = "Apagar"

Sub Apagar()

    Sheets("base").Select
            Range("J1:n1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    
    'Sheets("HxH PROMESSAS").Select
    
End Sub



    
