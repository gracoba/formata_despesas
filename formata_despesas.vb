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