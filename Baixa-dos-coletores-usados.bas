Attribute VB_Name = "Baixa"
Sub Baixa()
'
' Macro para copiar dados da planilha "HOME" e "UTILIZADOS" e colá-los na planilha "HISTORICO",
' além de excluir uma linha correspondente na planilha "UTILIZADOS" com base em um valor de "HOME".
'
    Dim wsOrigem As Worksheet, wsDestino As Worksheet
    Dim celula As Range, primeiroAchado As String
    Dim valorProcurado As Variant
    Dim rngAchados As Range, rngLinha As Range
    Dim i As Integer

    ' Define as planilhas
    Set wsOrigem = ThisWorkbook.Sheets("HOME")  ' Contem os valores de referencia
    Set wsDestino = ThisWorkbook.Sheets("UTILIZADOS") ' Onde sera feita a busca

    ' Armazena os valores de referencia das celulas A1, B1 e C1 da Planilha1
    Sheets("HOME").Select
    valorProcurado = Array(wsOrigem.Range("B7:F7").Value)

    ' Percorre cada valor de referencia
    For i = LBound(valorProcurado) To UBound(valorProcurado)
        ' Inicia a busca na Planilha2 (coluna A)
        With wsDestino.Range("A:A")
            Set celula = .Find(What:=valorProcurado(i), LookAt:=xlWhole, MatchCase:=False)

            ' Se encontrou pelo menos uma ocorrencia
            If Not celula Is Nothing Then
                primeiroAchado = celula.Address ' Guarda o primeiro endereÃƒÂ§o encontrado

                ' Adiciona A, B e C da linha encontrada
                Set rngAchados = wsDestino.Range(celula, celula.Offset(0, 4))

                ' Continua a busca para encontrar outras ocorrencias do mesmo valor
                Do
                    Set celula = .FindNext(celula)
                    If celula Is Nothing Then Exit Do
                    If celula.Address = primeiroAchado Then Exit Do

                    ' Adiciona A, B e C da nova ocorrencia encontrada
                    Set rngLinha = wsDestino.Range(celula, celula.Offset(0, 4))
                    Set rngAchados = Union(rngAchados, rngLinha)
                Loop
            End If
        End With
    Next i

    ' Se encontrou celulas, seleciona todas de uma vez
    If Not rngAchados Is Nothing Then
        Sheets("UTILIZADOS").Select
        rngAchados.Select
        ' MsgBox "Foram encontradas " & rngAchados.Areas.Count & " ocorrencias dos valores pesquisados.", vbInformation, "Busca Concluida"
    Else
        MsgBox "Nenhuma celula correspondente foi encontrada.", vbExclamation, "Busca Falhou"
    End If
    
    
    
    ' Copia a célula atualmente selecionada
    Selection.Copy
    
    ' Muda para a planilha "HISTORICO"
    Sheets("HISTORICO").Select
    
    ' Seleciona a última célula preenchida na coluna ID da tabela "Tabela4" e desce para a próxima linha vazia
    Range("Tabela4[[#Headers],[ID]]").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    
    ' Cola os valores copiados na próxima linha vazia, sem formatação
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    ' Retorna para a planilha "HOME"
    Sheets("HOME").Select
    
    ' Seleciona a célula F3 (Data-Hora)
    Range("F3").Select
    
    ' Cancela o modo de cópia (para evitar erro ao colar posteriormente)
    ' Application.CutCopyMode = False
    Selection.Copy
    
    ' Volta para a planilha "HISTORICO"
    Sheets("HISTORICO").Select

    ' Seleciona a última célula preenchida na coluna "Data_Entrega" da tabela "Tabela4" e desce para a próxima linha vazia
    Range("Tabela4[[#Headers],[Data_Entrega]]").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    
    ' Cola os dados copiados de F3 como **valores**, sem formatação
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        
    ' Muda para a planilha "UTILIZADOS"
    Sheets("UTILIZADOS").Select
    
    ' Cancela o modo de cópia novamente
    Application.CutCopyMode = False
    
    ' --------------------------------------
    '  EXCLUSÃO DE LINHA COM BASE EM UM VALOR DA PLANILHA "HOME"
    ' --------------------------------------

    Dim wsHome As Worksheet, wsUtilizados As Worksheet
    Dim tabela As ListObject
    Dim celula1 As Range
    Dim valorProcurado1 As String

    ' Define as planilhas
    Set wsHome = ThisWorkbook.Sheets("HOME")         ' Planilha de onde vem o valor
    Set wsUtilizados = ThisWorkbook.Sheets("UTILIZADOS") ' Planilha onde será feita a exclusão

    ' Define a tabela na planilha "UTILIZADOS"
    Set tabela = wsUtilizados.ListObjects("Tabela2") ' Substitua pelo nome real da sua tabela
    
    ' Obtém o valor da célula B7 da planilha "HOME"
    valorProcurado1 = wsHome.Range("B7").Value
    
    ' Verifica se a célula B7 não está vazia antes de continuar
    If valorProcurado1 = "" Then
        MsgBox "A célula B7 está vazia. Nenhuma linha será excluída.", vbExclamation, "Erro"
        Exit Sub ' Sai da macro se não houver um valor para buscar
    End If
    
    ' Procura o valor na primeira coluna da tabela "Tabela2" na planilha "UTILIZADOS"
    Set celula1 = tabela.DataBodyRange.Columns(1).Find(What:=valorProcurado1, LookAt:=xlWhole, MatchCase:=False)
    
        'Volta para a Planilha "Home"
    Sheets("HOME").Select
    Range("B3").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = ""

    'Limpa os dados modificados
    Range("E3").Select
    ActiveCell.FormulaR1C1 = ""
    Range("B3").Select
    

    'Limpa os dados modificados
    Range("B7").Select
    ActiveCell.FormulaR1C1 = ""
    Range("B3").Select
    
    ' Se encontrou o valor, exclui a linha correspondente
    If Not celula1 Is Nothing Then
        tabela.ListRows(celula1.Row - tabela.DataBodyRange.Row + 1).Delete
        MsgBox "Linha com o valor '" & valorProcurado1 & "' excluída com sucesso!", vbInformation, "Sucesso"
    Else
        MsgBox "O valor '" & valorProcurado1 & "' não foi encontrado na tabela.", vbExclamation, "Erro"
    End If

End Sub