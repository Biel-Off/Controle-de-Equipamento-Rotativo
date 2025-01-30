Attribute VB_Name = "Adicionar"
Sub Adicionar()
'
' Adicionar Macro
'
    'Seleciona os dados
    Range("B3").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy

    Dim ws As Worksheet

    ' Define a planilha de destino
    Set ws = Sheets("UTILIZADOS")

    ' Seleciona o cabecalho da coluna ID na tabela Tabela2
    ws.Select
    Range("Tabela2[[#Headers],[ID]]").Select

    ' Move para a ultima celula preenchida na coluna ID
    Selection.End(xlDown).Select

    ' Posiciona na proxima celula disponivel (primeira vazia abaixo da ultima preenchida)
    ActiveCell.Offset(1, 0).Select

    'Cola os valores sem formatação
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    'Volta para a Planilha "Home"
    Sheets("HOME").Select
    Range("B3").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = ""

    'Limpa os dados modificados
    Range("E3").Select
    ActiveCell.FormulaR1C1 = ""
    Range("H3").Select

End Sub
