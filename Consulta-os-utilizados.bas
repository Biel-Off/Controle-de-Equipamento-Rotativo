Attribute VB_Name = "Consulta"
Sub consulta()
'
' Consulta dos utilizados
'
    Range("usados").AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Range("H2:L2"), Unique:=False
'
End Sub
