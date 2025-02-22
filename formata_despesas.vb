Sub InicializaConfiguracoes()
    Dim i As Integer
    Dim colunas As Variant

    ' Lista de colunas de início
    colunas = Array(6, 7, 1, 1, 1, 1, 8) ' Exemplo: F, G, A, etc.

    ' Redimensiona o array dinâmico COL_INICIO
    ReDim COL_INICIO(LBound(colunas) To UBound(colunas))

    ' Atribui os valores ao array dinâmico
    For i = LBound(colunas) To UBound(colunas)
        COL_INICIO(i) = colunas(i)
    Next i

    ' Inicializa o array PLAN_EXEC
    Dim anos As Variant
    anos = Array("exec_2021", "exec_2022", "exec_2023", "exec_2024")
    ReDim PLAN_EXEC(LBound(anos) To UBound(anos))
    For i = LBound(anos) To UBound(anos)
        PLAN_EXEC(i) = anos(i)
    Next i

    ' Configurações de execução
    EXIBE_MENSAGENS = True
    EXIBE_BARRA_PROGRESSO = True
    GERA_RELATORIO = True
End Sub

Sub SaneiaDadosPlanilhaTroca()
    On Error GoTo Erro

    Dim ws As Worksheet
    Dim arrDados As Variant
    Dim i As Long, lastRow As Long

    ' Define a planilha
    Set ws = ThisWorkbook.Sheets(PLAN_PLAN_TROCA)

    ' Determina a última linha preenchida
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Carrega os dados na memória
    arrDados = ws.Range("A1:E" & lastRow).Value

    ' Percorre o array em memória
    For i = LBound(arrDados, 1) To UBound(arrDados, 1)
        ' Verifica se há erros na coluna 4
        If IsError(arrDados(i, 4)) Then
            ' Limpa a linha inteira se houver erro na coluna 4
            LimparLinha arrDados, i
        ElseIf IsEmpty(arrDados(i, 4)) Then
            ' Limpa a linha inteira se a coluna 4 estiver vazia
            LimparLinha arrDados, i
        ElseIf arrDados(i, 4) = "outros" Then
            ' Limpa a linha inteira se a coluna 4 for "outros"
            LimparLinha arrDados, i
        ElseIf IsError(arrDados(i, 3)) Or IsEmpty(arrDados(i, 3)) Or arrDados(i, 3) = 0 Then
            ' Verifica a coluna 3 (dados inválidos ou zero) e limpa a linha
            LimparLinha arrDados, i
        End If
    Next i

    ' Retorna os dados limpos para a tabela
    ws.Range("A1:E" & lastRow).Value = arrDados

    ' Log de sucesso
    RegistrarLog "SaneiaDadosPlanilhaTroca concluída com sucesso."
    Exit Sub

Erro:
    RegistrarLog "Erro em SaneiaDadosPlanilhaTroca: " & Err.Description
    MsgBox "Erro em SaneiaDadosPlanilhaTroca: " & Err.Description, vbCritical
End Sub

' ----------------------------------------------------------------------
' Sub-rotina auxiliar para limpar uma linha no array
' ----------------------------------------------------------------------
Sub LimparLinha(ByRef arr As Variant, linha As Long)
    Dim col As Long
    For col = LBound(arr, 2) To UBound(arr, 2)
        arr(linha, col) = "" ' Limpa cada célula da linha
    Next col
End Sub

