 Function resumo_exec(detalhe01 As String, detalhe02 As String, data As Date) As Double
 Dim tmp01, tmp02 As Double

If detalhe01 <> "" Then
tmp01 = resumo_exec_tmp(detalhe01, data)
Else
tmp01 = 0
End If

If detalhe02 <> "" Then
tmp02 = resumo_exec_tmp(detalhe02, data)
Else
tmp02 = 0
End If

resumo_exec = tmp01 + tmp02

End Function
 
Function resumo_exec_tmp(detalhe As String, data As Date) As Double
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim getPivotDataFormula As String
    Dim Ano As Long
    Dim Mes As Long
        
    Ano = Year(data)
    Mes = Month(data)
        
    ' Seleciona a planilha e a Tabela Dinâmica
    Set ws = ThisWorkbook.Sheets("tab_mes")
    Set pt = ws.PivotTables(1) ' Assume que há pelo menos uma Tabela Dinâmica na planilha
    
        ' Tenta obter um exemplo de valor usando GetPivotData
    On Error Resume Next
    getPivotDataFormula = "=GETPIVOTDATA(""Valor""," & ws.Name & "!$I$3," & _
                          """Data""," & Mes & "," & _
                          """Detalhe"",""" & detalhe & """," & _
                          """Anos""," & Ano & ")"
    resumo_exec_tmp = Evaluate(getPivotDataFormula) ' Avalia a fórmula no VBA
    On Error GoTo 0
End Function

Function resumo_exec_cartao(planilha As String, data As Date) As Double
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim getPivotDataFormula As String
    Dim Ano As Long
    Dim Mes As Long
        
    Ano = Year(data)
    Mes = Month(data)
        
    ' Seleciona a planilha e a Tabela Dinâmica
    Set ws = ThisWorkbook.Sheets(planilha)
    Set pt = ws.PivotTables(1) ' Assume que há pelo menos uma Tabela Dinâmica na planilha
    
        ' Tenta obter um exemplo de valor usando GetPivotData
    On Error Resume Next
    getPivotDataFormula = "=GETPIVOTDATA(""Valor""," & ws.Name & "!$K$1," & _
                          """Data""," & Mes & "," & _
                          """Anos""," & Ano & ")"
    resumo_exec_cartao = Evaluate(getPivotDataFormula) ' Avalia a fórmula no VBA
    On Error GoTo 0
End Function

Function resumo_exec_cheque(planilha As String, data As Date) As Double
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim getPivotDataFormula As String
    Dim Ano As Long
    Dim Mes As Long
        
    Ano = Year(data)
    Mes = Month(data)
        
    ' Seleciona a planilha e a Tabela Dinâmica
    Set ws = ThisWorkbook.Sheets(planilha)
    Set pt = ws.PivotTables(1) ' Assume que há pelo menos uma Tabela Dinâmica na planilha
    
        ' Tenta obter um exemplo de valor usando GetPivotData
    On Error Resume Next
    getPivotDataFormula = "=GETPIVOTDATA(""Valor""," & ws.Name & "!$J$1," & _
                          """Data""," & Mes & "," & _
                          """Anos""," & Ano & ")"
    resumo_exec_cheque = Evaluate(getPivotDataFormula) ' Avalia a fórmula no VBA
    On Error GoTo 0
End Function

Function resumo_exec_dinheiro(planilha As String, data As Date) As Double
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim getPivotDataFormula As String
    Dim Ano As Long
    Dim Mes As Long
        
    Ano = Year(data)
    Mes = Month(data)
        
    ' Seleciona a planilha e a Tabela Dinâmica
    Set ws = ThisWorkbook.Sheets(planilha)
    Set pt = ws.PivotTables(1) ' Assume que há pelo menos uma Tabela Dinâmica na planilha
    
        ' Tenta obter um exemplo de valor usando GetPivotData
    On Error Resume Next
        getPivotDataFormula = "=GETPIVOTDATA(""Valor""," & ws.Name & "!$J$1," & _
                              """Data""," & Mes & "," & _
                              """Anos""," & Ano & ")"
    resumo_exec_dinheiro = Evaluate(getPivotDataFormula) ' Avalia a fórmula no VBA
    On Error GoTo 0
End Function