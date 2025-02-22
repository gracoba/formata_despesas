Sub InicializaConfiguracoes()
    Dim i As Integer
    Dim anos As Variant

    ' Lista de anos para execução
    anos = Array("exec_2021", "exec_2022", "exec_2023", "exec_2024")

    ' Redimensiona o array dinâmico PLAN_EXEC
    ReDim PLAN_EXEC(LBound(anos) To UBound(anos))

    ' Atribui os valores ao array dinâmico
    For i = LBound(anos) To UBound(anos)
        PLAN_EXEC(i) = anos(i)
    Next i

    ' Colunas de início para consolidação
    COL_INICIO = Array(6, 7, 1, 1, 1, 1, 8) ' Exemplo: F, G, A, etc.

    ' Configurações de execução
    EXIBE_MENSAGENS = True
    EXIBE_BARRA_PROGRESSO = True
    GERA_RELATORIO = True
End Sub