Attribute VB_Name = "Módulo_Relatório"
Sub progress(pctCompl As Single)

UserForm_Processando.Texto.Caption = pctCompl & "% Completo"
UserForm_Processando.Barra.Width = pctCompl * 3
DoEvents

End Sub
Sub Get_Data()
    
    'Travamento de tela é ativado no procedimento que chama a macro
    '(Código da Userform Processando)e destivado ao final da mesma.
    
    Dim ARGeral As String
    ARGeral = UserForm_Settings.TextBox_ArquivoAging.Text
    'Cria uma variável referente ao nome do arquivo aberto com os dados
    'do Aging e atribui ao texto inserido na caixa do menu opções.
    
    
    Dim i As Integer, pctCompl As Single
    UserForm_Processando.Label_Processando.Caption = "Processando..."
    i = 0
    pctCompl = i
    progress pctCompl
    'Inicia a contabilização do avanço na barra de progresso
    
    On Error GoTo Erro_nome_do_arquivo
    Windows(ARGeral).Activate
    'Tratamento de erro para arquivo aging não reconhecido: caso falhe ao ativar o aging, o código será
    'levado à uma seção onde informa o usuário sobre o erro, como proceder e interrompe o procedimento.
    
    'Caso não haja erro, o código prosseguirá com a filtragem e manipulação dos dados do Aging.
    
    'Os procedimentos abaixo separam as colunas importantes e realizam algumas exceções por filtragem.
    Range("A1").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$BC$34532").AutoFilter Field:=1, Criteria1:= _
        "<>1010405", Operator:=xlAnd
    ActiveSheet.Range("$A$1:$BA$34532").AutoFilter Field:=7, Criteria1:="<>IL" _
        , Operator:=xlAnd
    Range("A:A,B:B,G:G,I:I,J:J,K:K,L:L,M:M,AE:AE").Select
    Selection.Copy
    
    Sheets.Add After:=ActiveSheet
    'Adiciona uma nova planilha para jogar e organizar os dados que serão utilizados
    
    ActiveSheet.Paste
    Worksheets(1).Activate
    Columns("P:P").Select
    Selection.Copy
    Worksheets(2).Activate
    Columns("J:J").Select
    ActiveSheet.Paste
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("C1").Value = "Tipo"
    Worksheets(1).Activate
    Columns("Z:Z").Select
    Selection.Copy
    Worksheets(2).Activate
    Columns("L:L").Select
    ActiveSheet.Paste
    
    Range("M1").FormulaR1C1 = "=COUNT(C[-12])+1"
    'conta as linhas preenchidas e adiciona +1 (cabeçalho) para descobrir o Nº da última linha preenchida.
    '(também poderia ser feito por cálculo na memória, através do
    'método "worksheet.function", sem a necessidade de hospedar a fórmula numa célula)
    
    Range("C2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[9]=""DST"",""DIS"",IF(OR(RC[9]=""C26"",RC[9]=""C87""),""PUB"",""PRI""))"
    'Fórmula que classifica os clientes em públicos e privados
    
    Selection.Copy

    
    Dim Linha As String
    Linha = Range("M1").Value
    'Cria uma variável e a atribui ao número da última linha preenchida calculado na célula M1
    
    Range("C" & "2" & ":" & "C" & Linha).Select
    'Faz a seleção de "C2" à ultima célula preenchida
    'Ao invés de range também pode ser utilizado o método "Cells" que utiliza coordenadas ao invés
    'do nome, sem a necessidade de aninhar os textos com a variável para formar o range.
    ActiveSheet.Paste
    'Cola em a fórmula em todas as linhas com conteúdo.
    
    i = 10
    pctCompl = i
    progress pctCompl
    'Barra de progresso em 10%
    
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< FECHAMENTO <<<<<<<<<<<<<<<<<<<<<<<<<<<<
    Dim DataFechamento As Date
    DataFechamento = UserForm_RelatórioPDD.Label_DataFech.Caption '<<<<
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'Cria uma variável para data de fechamento selecionada no formulário, que será usada para calcular os dias vencidos.

    Range("N1").Select
    ActiveCell.FormulaR1C1 = DataFechamento
    'insere a data na célula para ser usada de referência (a variável também poderia ser usada direto na fórmula).
    
    Range("N2").Select
    ActiveCell.FormulaR1C1 = "=IF((R1C14-RC[-5])>180,R1C14-RC[-5],"""")"
    'Fórmula que será usada para pré selecionar somente as notas vencidas acima de 180 dias,
    'já que todos os critérios só enquadram notas acima desse número.
    
    Selection.Copy
    Range("N" & "2" & ":" & "N" & Linha).Select
    ActiveSheet.Paste
    'novamente cola a fórmula somente nas linhas com conteúdo
    
    'O bloco abaixo exclui dos dados mais algumas exceções (notas intercompany)
    ActiveSheet.Range("$A:$N").AutoFilter Field:=4, Criteria1:="EX"
    ActiveSheet.Range("$A:$N").AutoFilter Field:=2, Criteria1:= _
        "<>GAMMA LTDA", Operator:=xlAnd, Criteria2:="<>EL ALAMO SA"
    ActiveSheet.Range("$A:$N").AutoFilter Field:=1, Criteria1:= _
        "<>5225882", Operator:=xlAnd
    With Worksheets(2).AutoFilter.Range
    Range("A" & .Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row).Select
    End With
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.Offset(0, 13)).Select
    Selection.ClearContents
    ActiveSheet.ShowAllData
    
    ActiveSheet.Range("$A:$N").AutoFilter Field:=14, Criteria1:="<>", _
        Operator:=xlAnd
    'Remove as notas vencidas a menos de 180 dias (que ficaram com a célula vazia após a fórmula anterior)
    
    Columns("A:N").Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Paste
    Range("N1").Value = "Dias Vencidos"
    
    'Os blocos abaixo iniciam o processo de filtragem dos 3 primeiros critérios referentes à legislação
    'antiga (anterior à 08/07/2014) e os organizam numa outra planilha que será o formato final trazido no relatório.
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'<<<<< CRITÉRIO 1: VENCIDO > 180 DIAS ATÉ  R$ 5.000,00 - Até 07/10/14 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    
    'como só restaram as notas vencidas acima de 180 dias não será necessárui filtrar dias vencidos, somente valor e data.
    Range("N1").AutoFilter
    ActiveSheet.Range("$A:$N").AutoFilter Field:=11, Criteria1:="<5000" _
        , Operator:=xlAnd
    'Filtra o valor
    
    i = 20
    pctCompl = i
    progress pctCompl
    'Barra de progresso 20%
    

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    ActiveSheet.Range("$A$1:$N$8033").AutoFilter Field:=9, Criteria1:="<10/08/2014", Operator:=xlAnd '<<<<<
'<<<<<<<<<<<<<<*VBA considera a o formato americano na execução do AutoFilter ("mm/dd/aaaa")<<<<<<<<<<<<<<<
    'Filtra a data
    
    
    Columns("A:K").Select
    Selection.Copy
    'Copia as notas que entraram no primeiro critério após as filtragens
    
    Sheets.Add After:=ActiveSheet
    'Adiciona a nova planilha para onde serão jogadas as notas que entraram em cada critério após os filtros.
    ActiveSheet.Paste
    'Cola o critério 1
    
    Range("L1").Value = "Critério"
    Range("M1").FormulaR1C1 = "=COUNTA(C[-12])"
    'Outra fórmula par descobrir a última linha preenchida.

    Dim LinhaFinalCritério_1 As Long
    LinhaFinalCritério_1 = Range("M1").Value
    'Cria uma variável para a última nota enquadrada no critério 1 (última linha preenchida no momento).
    
    Worksheets(3).Activate
    'Volta á planilha de filtragem para a aplicação do próximo critério.
    
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'<<<<< CRITÉRIO 2: VENCIDO > 360 DIAS ACIMA DE  R$ 30.000,00 EM JUÍZO - Até 07/10/14 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    
    Range("A1").Select
    ActiveSheet.Range("$A:$N").AutoFilter Field:=14, Criteria1:=">360", _
        Operator:=xlAnd
    'Dias vencidos agora acima de 360
    
    ActiveSheet.Range("$A:$N").AutoFilter Field:=11, Criteria1:=">30000" _
        , Operator:=xlAnd
    'Altera a filtragem do valor
    
    ActiveSheet.Range("$A:$N").AutoFilter Field:=10, Criteria1:="=L", _
        Operator:=xlAnd
    'As linhas com "L" na coluna estão em juizo, um dos requisitos do critério 2
    
    Columns("A:K").Select
    Selection.Copy
    'Copia as notas que se enquadraram no critério 2.
    
    Worksheets(4).Activate
    'Ativa a planilha para qual serão jogadas.
    
    Range("A1").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    ActiveSheet.Paste
    'Cola as notas do critério 2 ao final das do critério 1.
    
    Rows(LinhaFinalCritério_1 + 1 & ":" & LinhaFinalCritério_1 + 1).Select
    'seleciona o cabeçalho da filtragem do critério 2 que foi levado junto.
    Selection.Delete Shift:=xlUp
    'deleta o cabeçalho
    
    i = 30
    pctCompl = i
    progress pctCompl
    'Progresso 30%
    
    Dim LinhaFinalCritério_2 As Long
    LinhaFinalCritério_2 = Range("M1").Value
    'Cria uma variável para a última nota enquadrada no critério 2 (última linha preenchida no momento).
    
    Dim LinhaInicialCritério_2 As Long
    LinhaInicialCritério_2 = LinhaFinalCritério_1 + 1
    'Cria uma variável para a primeira nota enquadrada no critério 2 (linha seguinte à última do critério 1)
    
    Worksheets(3).Activate
    'Volta á planilha de filtragem para a aplicação do próximo critério.
    
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'<<<<< CRITÉRIO 3: VENCIDO > 360 DIAS, ACIMA DE  R$ 5.000,00 ATÉ R$ 30.000,00 - Até 07/10/14 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    
    ActiveSheet.ShowAllData
    'Dessa vez foi necessário limpar os filtros para remover a limitação "em juizo".
    
    ActiveSheet.Range("$A:$N").AutoFilter Field:=14, Criteria1:=">360", Operator:=xlAnd
    'Filtra a quantidade de dias vencidos.
    
    ActiveSheet.Range("$A:$N").AutoFilter Field:=11, Criteria1:=">5000", Operator:=xlAnd, Criteria2:="<=30000"
    'Filtra o valor em aberto.
    
    ActiveSheet.Range("$A:$N").AutoFilter Field:=9, Criteria1:="<10/08/2014", Operator:=xlAnd
    'Filtra a data novamente.
    
    Selection.Copy
    'Como a planilha já estava selecionada e a seleção não foi alterada,
    'copia agora as notas que se enquadraram no critério 3.
    
    Worksheets(4).Activate
    Range("A1").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    ActiveSheet.Paste
    Rows(LinhaFinalCritério_2 + 1 & ":" & LinhaFinalCritério_2 + 1).Select
    Selection.Delete Shift:=xlUp
    'Novamente lança a filtragem para a planilha onde está sendo organizado o relatório e apaga o cabeçalho.
    
    Dim LinhaFinalCritério_3 As Long
    LinhaFinalCritério_3 = Range("M1").Value
    Dim LinhaInicialCritério_3 As Long
    LinhaInicialCritério_3 = LinhaFinalCritério_2 + 1
    'Novamente são criadas variáveis para linha final e inicial.
    
    
    Worksheets(3).Activate
    
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'<<<<< CRITÉRIO 4: VENCIDO > 180 DIAS ATÉ R$ 15.000,00 - Após 07/10/14 <<<<<<<<<<<<<<<<<<<<<<<<<<  <<<
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

    ActiveSheet.ShowAllData
    ActiveSheet.Range("$A$1:$N$7850").AutoFilter Field:=11, Criteria1:=">0" _
        , Operator:=xlAnd, Criteria2:="<=15000"
    
    ActiveSheet.Range("$A$1:$N$8033").AutoFilter Field:=9, Criteria1:=">10/07/2014", Operator:=xlAnd
    Selection.Copy
    Worksheets(4).Activate
    Range("A1").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    ActiveSheet.Paste
    Rows(LinhaFinalCritério_3 + 1 & ":" & LinhaFinalCritério_3 + 1).Select
    Selection.Delete Shift:=xlUp
    
    i = 40
    pctCompl = i
    progress pctCompl


    Dim LinhaFinalCritério_4 As Long
    LinhaFinalCritério_4 = Range("M1").Value
    Dim LinhaInicialCritério_4 As Long
    LinhaInicialCritério_4 = LinhaFinalCritério_3 + 1
    
    
    Worksheets(3).Activate
    
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'<<<<< CRITÉRIO 5: VENCIDO > 360 DIAS, ACIMA DE  R$ 15.000,00 ATÉ R$ 100.000,00 - Após 07/10/14 <<<<<<<<<<<<<<<<<<<<
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
 
    
    ActiveSheet.Range("$A$1:$N$8033").AutoFilter Field:=14, Criteria1:=">360", _
        Operator:=xlAnd
    ActiveSheet.Range("$A$1:$N$7850").AutoFilter Field:=11, Criteria1:=">15000" _
        , Operator:=xlAnd, Criteria2:="<=100000"
    Selection.Copy
    Worksheets(4).Activate
    Range("A1").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    ActiveSheet.Paste
    Rows(LinhaFinalCritério_4 + 1 & ":" & LinhaFinalCritério_4 + 1).Select
    Selection.Delete Shift:=xlUp
    Dim LinhaFinalCritério_5 As Long
    LinhaFinalCritério_5 = Range("M1").Value
    Dim LinhaInicialCritério_5 As Long
    LinhaInicialCritério_5 = LinhaFinalCritério_4 + 1
    
    Worksheets(3).Activate
    
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'<<<<< CRITÉRIO 6: VENCIDO > 720 DIAS, ACIMA DE  R$ 100.000,00 - Após 07/10/14 <<<<<<<<<<<<<<<<<<<<<<<<<<
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    
    ActiveSheet.Range("$A$1:$N$7850").AutoFilter Field:=11, Criteria1:=">100000" _
        , Operator:=xlAnd
    ActiveSheet.Range("$A$1:$N$8033").AutoFilter Field:=14, Criteria1:=">360", _
        Operator:=xlAnd
    ActiveSheet.Range("$A$1:$N$8033").AutoFilter Field:=10, Criteria1:="=L", _
        Operator:=xlAnd
    Selection.Copy
    Worksheets(4).Activate
    Range("A1").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    ActiveSheet.Paste
    Rows(LinhaFinalCritério_5 + 1 & ":" & LinhaFinalCritério_5 + 1).Select
    Selection.Delete Shift:=xlUp
    
    i = 50
    pctCompl = i
    progress pctCompl
    

    Dim LinhaFinalCritério_6 As Long
    LinhaFinalCritério_6 = Range("M1").Value
    Dim LinhaInicialCritério_6 As Long
    LinhaInicialCritério_6 = LinhaFinalCritério_5 + 1
    Columns("K:K").Select
    Selection.Copy
    Columns("R:R").Select
    ActiveSheet.Paste
    Range("K" & "2" & ":" & "K" & LinhaFinalCritério_1).Select
    Selection.Copy
    Range("L" & "2" & ":" & "L" & LinhaFinalCritério_1).Select
    ActiveSheet.Paste
    Range("K" & LinhaInicialCritério_2 & ":" & "K" & LinhaFinalCritério_2).Select
    Selection.Copy
    Range("M" & LinhaInicialCritério_2 & ":" & "M" & LinhaFinalCritério_2).Select
    ActiveSheet.Paste
    Range("K" & LinhaInicialCritério_3 & ":" & "K" & LinhaFinalCritério_3).Select
    Selection.Copy
    Range("N" & LinhaInicialCritério_3 & ":" & "N" & LinhaFinalCritério_3).Select
    ActiveSheet.Paste
    Range("K" & LinhaInicialCritério_4 & ":" & "K" & LinhaFinalCritério_4).Select
    Selection.Copy
    Range("O" & LinhaInicialCritério_4 & ":" & "O" & LinhaFinalCritério_4).Select
    ActiveSheet.Paste
    Range("K" & LinhaInicialCritério_5 & ":" & "K" & LinhaFinalCritério_5).Select
    Selection.Copy
    Range("P" & LinhaInicialCritério_5 & ":" & "P" & LinhaFinalCritério_5).Select
    ActiveSheet.Paste
    Range("K" & LinhaInicialCritério_6 & ":" & "K" & LinhaFinalCritério_6).Select
    Selection.Copy
    Range("Q" & LinhaInicialCritério_6 & ":" & "Q" & LinhaFinalCritério_6).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    Windows("AFDA Report.xlsm").Activate
    
    Call Clear_Report
    
    i = 75
    pctCompl = i
    progress pctCompl
    
    Windows(ARGeral).Activate

    Range("A2").Select
    Range(Selection, Selection.Offset(0, 17)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.Offset(1, 0)).Select
    Selection.Copy

    Windows("AFDA Report.xlsm").Activate

    Range("A10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("A10").Select
    Range(Selection, Selection.Offset(0, 17)).Select
    Selection.Copy
    Range(Selection, Selection.End(xlDown)).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
    Range("K" & LinhaFinalCritério_6 + 9).Activate
    ActiveCell.FormulaR1C1 = "=SUM(R10C11:R[-1]C)"
    Range("L" & LinhaFinalCritério_6 + 9).Activate
    ActiveCell.FormulaR1C1 = "=SUM(R10C12:R[-1]C)"
    Range("M" & LinhaFinalCritério_6 + 9).Activate
    ActiveCell.FormulaR1C1 = "=SUM(R10C13:R[-1]C)"
    Range("N" & LinhaFinalCritério_6 + 9).Activate
    ActiveCell.FormulaR1C1 = "=SUM(R10C14:R[-1]C)"
    Range("O" & LinhaFinalCritério_6 + 9).Activate
    ActiveCell.FormulaR1C1 = "=SUM(R10C15:R[-1]C)"
    Range("P" & LinhaFinalCritério_6 + 9).Activate
    ActiveCell.FormulaR1C1 = "=SUM(R10C16:R[-1]C)"
    Range("Q" & LinhaFinalCritério_6 + 9).Activate
    ActiveCell.FormulaR1C1 = "=SUM(R10C17:R[-1]C)"
    Range("R" & LinhaFinalCritério_6 + 9).Activate
    ActiveCell.FormulaR1C1 = "=SUM(R10C18:R[-1]C)"
    Range(Selection, Selection.Offset(0, -7)).Select
      Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlDouble
        .Color = -6750208
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlDouble
        .Color = -6750208
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlDouble
        .Color = -6750208
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlDouble
        .Color = -6750208
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlDouble
        .Color = -6750208
        .TintAndShade = 0
        .Weight = xlThick
    End With
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Selection.Font.Bold = True
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    Application.CutCopyMode = False


    i = 100
    pctCompl = i
    progress pctCompl
    UserForm_Processando.Label_Processando.Caption = "Concluído!"
    Application.ScreenUpdating = True

    Windows(ARGeral).Visible = True

    
    Exit Sub
    
Erro_nome_do_arquivo:

    Unload UserForm_Processando
If UserForm_RelatórioPDD.CommandButton_Process.Caption = "Processar" Then
    MsgBox "Certifique-se de que o Aging está aberto e nomeado como """ & ARGeral & """(Manual Página 4)", vbOKOnly, "Aging não encontrado"
    End If
If UserForm_RelatórioPDD.CommandButton_Process.Caption = "Procesar" Then
    MsgBox "Asegúrese de que el archivo esté abierto y nombrado como """ & ARGeral & """(Manual Página 4)", vbOKOnly, "Aging no encontrado"
    End If
    
End Sub



