Attribute VB_Name = "Módulo_Relatório"
Sub progress(pctCompl As Single)

UserForm_Processando.Texto.Caption = pctCompl & "% Completo"
UserForm_Processando.Barra.Width = pctCompl * 3
DoEvents

End Sub
Sub Get_Data()
    
    'Travamento de tela é ativado no procedimento que chama a macro
    '(Código da Userform Processando)e destivado ao final da mesma.
    
    On Error GoTo Tool_File_Name_Error
    Workbooks("AFDA Report Tool.xlsm").Activate
    On Error GoTo 0
    'Verifica o nome do arquivo da ferramenta, pois será necessário no final para voltar e trazer os dados _
    (ao voltar poderia usar a propriedade index, porém é imprecisa caso o usuário esteja usando outras planilhas)
    
    Dim ARGeral As String
    ARGeral = UserForm_Settings.TextBox_ArquivoAging.Text
    'Cria uma variável referente ao nome do arquivo aberto com os dados _
    do Aging e atribui ao texto inserido na caixa do menu opções.
    
    Dim i As Integer, pctCompl As Single
    If UserForm_Idioma.ToggleButton_Português = True Then
    UserForm_Processando.Label_Processando.Caption = "Processando..."
    Else
    UserForm_Processando.Label_Processando.Caption = "Procesando..."
    End If
    
    i = 0
    pctCompl = i
    progress pctCompl
    'Inicia a contabilização do avanço na barra de progresso
    
    On Error GoTo Aging_File_Name_Error
    Workbooks(ARGeral).Activate
    'Tratamento de erro para arquivo aging não reconhecido: caso falhe ao ativar o aging, o código será
    'levado à uma seção onde informa o usuário sobre o erro, como proceder e interrompe o procedimento.
    On Error GoTo 0
    'Caso não haja erro, o código prosseguirá com a filtragem e manipulação dos dados do Aging.
    
    'Os procedimentos abaixo separam as colunas importantes e realizam algumas exceções por filtragem.
    Range("A1").AutoFilter
    ActiveSheet.Range("$A$1:$BC$34532").AutoFilter Field:=1, Criteria1:= _
        "<>1010405", Operator:=xlAnd
    ActiveSheet.Range("$A$1:$BA$34532").AutoFilter Field:=7, Criteria1:="<>IL" _
        , Operator:=xlAnd
    
    'Adiciona uma nova planilha para "jogar" e organizar somente os dados que serão utilizados
    Sheets.Add After:=ActiveSheet
    
    Worksheets(1).Activate
    Range("A:A,B:B,G:G,I:I,J:J,K:K,L:L,M:M,AE:AE").Copy Destination:=Worksheets(2).Range("A1")
    Columns("P:P").Copy Destination:=Worksheets(2).Columns("J:J")
  
    Worksheets(2).Activate
    Columns("C:C").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("C1").Value = "Tipo"
    Worksheets(1).Columns("Z:Z").Copy Destination:=Worksheets(2).Columns("L:L")
    
    'conta as linhas preenchidas e adiciona para descobrir o Nº da última linha.
    '(também pode ser feito por cálculo na memória, através do
    'método "WorksheetFunction", sem a necessidade de "hospedar" a fórmula numa célula)
    
    Dim NLinha As String
    NLinha = WorksheetFunction.CountA(Range("A:A"))
    'Cria uma variável e a atribui ao número da última linha preenchida
   
    Range("C2:C" & NLinha).FormulaR1C1 = _
        "=IF(RC[9]=""DST"",""DIS"",IF(OR(RC[9]=""C26"",RC[9]=""C87""),""PUB"",""PRI""))"
    'Fórmula que classifica os clientes em públicos e privados inserida na coluna ate a ultima linha
    Range("C2:C" & NLinha).Value = Range("C2:C" & NLinha).Value
    'Mantem somente o valor resultante da formula, como não precisará ser atualizada isso economiza processamento
    
    i = 10
    pctCompl = i
    progress pctCompl
    'Avança a barra de progresso
    
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< FECHAMENTO <<<<<<<<<<<<<<<<<<<<<<<<<<<<
    Dim DataFechamento As Date                                    '<<<<
    DataFechamento = UserForm_RelatórioPDD.Label_DataFech.Caption '<<<<
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'Cria uma variável para data de fechamento selecionada no formulário, que será usada para calcular os dias vencidos.
    Range("N1").Value = DataFechamento
    Range("N" & "2" & ":" & "N" & NLinha).FormulaR1C1 = "=IF((R1C14-RC[-5])>180,R1C14-RC[-5],"""")"
    'Fórmula que será usada para separar somente as notas vencidas acima de 180 dias,
    'pois somente à partir daí entram em algum critério de PDD
    Range("N" & "2" & ":" & "N" & NLinha).Value = Range("N" & "2" & ":" & "N" & NLinha).Value
    'Mantem somente o valor resultante da formula, como não precisará ser atualizada isso economiza processamento
    'O bloco abaixo exclui dos dados mais algumas exceções (notas intercompany)
    
    With ActiveSheet.Range("$A:$N")
    .AutoFilter Field:=4, Criteria1:="EX"
    .AutoFilter Field:=2, Criteria1:= _
        "<>GAMMA LTDA", Operator:=xlAnd, Criteria2:="<>EL ALAMO SA"
    .AutoFilter Field:=1, Criteria1:= _
        "<>5225882", Operator:=xlAnd
    End With
    
    With Worksheets(2).AutoFilter.Range
    Dim Exceptions As Range
    Set Exceptions = Range(Range("A" & .Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row), Range("N" & .Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row).End(xlDown))
    End With
    Exceptions.Delete Shift:=xlUp
    ActiveSheet.ShowAllData
    
    ActiveSheet.Range("$A:$N").AutoFilter Field:=14, Criteria1:="", _
        Operator:=xlAnd
    'Separa somente as notas vencidas a mais de 180 dias (excluindo as que ficaram vazias (vazio = critério "falso" na fórmula))
    With Worksheets(2).AutoFilter.Range
    Dim MenorIgua180 As Range
    Set MenorIgua180 = Range(Range("A" & .Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row), Range("N" & .Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row).End(xlDown))
    End With
    MenorIgua180.Delete Shift:=xlUp
    ActiveSheet.ShowAllData
     
    Range("N1").Value = "Dias Vencidos"
    
    'Os blocos abaixo iniciam o processo de filtragem dos 3 primeiros critérios referentes à legislação
    'antiga (anterior à 08/07/2014) e os organizam numa outra planilha que será o formato final trazido no relatório.
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'<<<<< CRITÉRIO 1: VENCIDO > 180 DIAS ATÉ  R$ 5.000,00 - Até 07/10/14 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    
    'como só restaram as notas vencidas acima de 180 dias não será necessárui filtrar dias vencidos, somente valor e data.
    ActiveSheet.Range("$A:$N").AutoFilter Field:=11, Criteria1:="<5000" _
        , Operator:=xlAnd
    'Filtra o valor
    
    i = 20
    pctCompl = i
    progress pctCompl
    'Barra de progresso 20%
    

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    ActiveSheet.Range("A:N").AutoFilter Field:=9, Criteria1:="<10/08/2014", Operator:=xlAnd '<<<<<
                            '*MENOR QUE 08/07 É O MESMO QUE MENOR IGUAL A 07/10             '<<<<<
'<<<<<<<<<<<<<<*VBA considera a o formato americano("mm/dd/aaaa")<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    'Filtra a data de vencimento
    
    Sheets.Add After:=ActiveSheet
    'Adiciona a nova planilha para onde serão jogadas as notas que entrarno relatório.
    
    Worksheets(2).Columns("A:K").Copy Destination:=Worksheets(3).Range("A1")
    'Copia as notas que entraram no primeiro critério após as filtragens e cola na nova planilha.

    Range("L1").Value = "Critério"

    Dim LinhaFinalCritério_1 As Long
    LinhaFinalCritério_1 = WorksheetFunction.CountA(Range("A:A"))
    'Cria uma variável para a última nota enquadrada no critério 1 (última linha preenchida no momento).
    
    Worksheets(2).Activate
    'Volta á planilha de filtragem para a aplicação do próximo critério.
    
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'<<<<< CRITÉRIO 2: VENCIDO > 360 DIAS ACIMA DE  R$ 30.000,00 EM JUÍZO - Até 07/10/14 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    
    With ActiveSheet.Range("$A:$N")
    .AutoFilter Field:=14, Criteria1:=">360", _
        Operator:=xlAnd
    'Dias vencidos agora acima de 360
    
    .AutoFilter Field:=11, Criteria1:=">30000" _
        , Operator:=xlAnd
    'Altera a filtragem do valor
    
    .AutoFilter Field:=10, Criteria1:="=L", _
        Operator:=xlAnd
    'As linhas com "L" na coluna estão em juizo, um dos requisitos do critério 2
    End With
    
    Columns("A:K").Copy Destination:=Worksheets(3).Range("A1").End(xlDown).Offset(1, 0)
    'Copia as notas que se enquadraram no critério 2 e cola ao final das do critério 1.
    
    i = 30
    pctCompl = i
    progress pctCompl
    'Progresso 30%
    
    Worksheets(3).Activate
    
    'Cria uma variável para a última nota enquadrada no critério 2 (última linha preenchida no momento).
    
    Dim LinhaInicialCritério_2 As Long
    LinhaInicialCritério_2 = LinhaFinalCritério_1 + 1
    'Cria uma variável para a primeira nota enquadrada no critério 2 (linha seguinte à última do critério 1)
    Rows(LinhaInicialCritério_2).Delete Shift:=xlUp
    'deleta o cabeçalho da filtragem do critério 2 que foi levado junto e ficou abaixo da ultima do critério 1.
    Dim LinhaFinalCritério_2 As Long
    LinhaFinalCritério_2 = WorksheetFunction.CountA(Range("A:A"))
    Worksheets(2).Activate
    'Volta á planilha de filtragem para a aplicação do próximo critério.
    
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'<<<<< CRITÉRIO 3: VENCIDO > 360 DIAS, ACIMA DE  R$ 5.000,00 ATÉ R$ 30.000,00 - Até 07/10/14 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    
    ActiveSheet.ShowAllData
    'Dessa vez foi necessário limpar os filtros para remover a limitação "em juizo".
    
    With ActiveSheet.Range("$A:$N")
    .AutoFilter Field:=14, Criteria1:=">360", Operator:=xlAnd
    'Filtra a quantidade de dias vencidos.
    
    .AutoFilter Field:=11, Criteria1:=">5000", Operator:=xlAnd, Criteria2:="<=30000"
    'Filtra o valor em aberto.
    
    .AutoFilter Field:=9, Criteria1:="<10/08/2014", Operator:=xlAnd
    'Filtra a data novamente.
    End With
    
    Columns("A:K").Copy Destination:=Worksheets(3).Range("A1").End(xlDown).Offset(1, 0)
    
    'copia agora as notas que se enquadraram no critério 3 e novamente _
    lança a filtragem para a planilha onde está sendo organizado o relatório
    
    Worksheets(3).Activate
    
    Dim LinhaInicialCritério_3 As Long
    LinhaInicialCritério_3 = LinhaFinalCritério_2 + 1
    
    Rows(LinhaInicialCritério_3).Delete Shift:=xlUp
    'Apaga o cabeçalho.
    
    Dim LinhaFinalCritério_3 As Long
    LinhaFinalCritério_3 = WorksheetFunction.CountA(Range("A:A"))
    'Novamente são criadas variáveis para linha final e inicial.

    
    Worksheets(2).Activate
    
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'<<<<< CRITÉRIO 4: VENCIDO > 180 DIAS ATÉ R$ 15.000,00 - Após 07/10/14 <<<<<<<<<<<<<<<<<<<<<<<<<<  <<<
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

    ActiveSheet.ShowAllData
    
    With ActiveSheet.Range("A:N")
    .AutoFilter Field:=11, Criteria1:=">0" _
        , Operator:=xlAnd, Criteria2:="<=15000"
    
    .AutoFilter Field:=9, Criteria1:=">10/07/2014", Operator:=xlAnd
    End With
    
    Columns("A:K").Copy Destination:=Worksheets(3).Range("A1").End(xlDown).Offset(1, 0)
    
    i = 40
    pctCompl = i
    progress pctCompl
    
    Worksheets(3).Activate

    Dim LinhaInicialCritério_4 As Long
    LinhaInicialCritério_4 = LinhaFinalCritério_3 + 1
    
    Rows(LinhaInicialCritério_4).Delete Shift:=xlUp

    Dim LinhaFinalCritério_4 As Long
    LinhaFinalCritério_4 = WorksheetFunction.CountA(Range("A:A"))
    
    Worksheets(2).Activate
    
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'<<<<< CRITÉRIO 5: VENCIDO > 360 DIAS, ACIMA DE  R$ 15.000,00 ATÉ R$ 100.000,00 - Após 07/10/14 <<<<<<<<<<<<<<<<<<<<
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
 
    
    With ActiveSheet.Range("A:N")
    
    .AutoFilter Field:=14, Criteria1:=">360", _
        Operator:=xlAnd
    .AutoFilter Field:=11, Criteria1:=">15000" _
        , Operator:=xlAnd, Criteria2:="<=100000"
    End With
    
    Columns("A:K").Copy Destination:=Worksheets(3).Range("A1").End(xlDown).Offset(1, 0)
    
    Worksheets(3).Activate
    Dim LinhaInicialCritério_5 As Long
    LinhaInicialCritério_5 = LinhaFinalCritério_4 + 1
  
    Rows(LinhaInicialCritério_5).Delete Shift:=xlUp
    
    Dim LinhaFinalCritério_5 As Long
    LinhaFinalCritério_5 = WorksheetFunction.CountA(Range("A:A"))
      
    Worksheets(2).Activate
    
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'<<<<< CRITÉRIO 6: VENCIDO > 720 DIAS, ACIMA DE  R$ 100.000,00 - Após 07/10/14 <<<<<<<<<<<<<<<<<<<<<<<<<<
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    
    With ActiveSheet.Range("A:N")
    .AutoFilter Field:=11, Criteria1:=">100000" _
        , Operator:=xlAnd
    .AutoFilter Field:=14, Criteria1:=">360", _
        Operator:=xlAnd
    .AutoFilter Field:=10, Criteria1:="=L", _
        Operator:=xlAnd
    End With
    
    Columns("A:K").Copy Destination:=Worksheets(3).Range("A1").End(xlDown).Offset(1, 0)
    
    i = 50
    pctCompl = i
    progress pctCompl
    
    Worksheets(3).Activate

    Dim LinhaInicialCritério_6 As Long
    LinhaInicialCritério_6 = LinhaFinalCritério_5 + 1
    
    Rows(LinhaInicialCritério_6).Delete Shift:=xlUp
    
    Dim LinhaFinalCritério_6 As Long
    LinhaFinalCritério_6 = WorksheetFunction.CountA(Range("A:A"))

    Columns("K:K").Copy Destination:=Columns("R:R")

    Range("K2:K" & LinhaFinalCritério_1).Copy Destination:=Range("L2:L" & LinhaFinalCritério_1)
    
    Range("K" & LinhaInicialCritério_2 & ":K" & LinhaFinalCritério_2).Copy _
    Destination:=Range("M" & LinhaInicialCritério_2 & ":" & "M" & LinhaFinalCritério_2)
    
    Range("K" & LinhaInicialCritério_3 & ":" & "K" & LinhaFinalCritério_3).Copy _
    Destination:=Range("N" & LinhaInicialCritério_3 & ":" & "N" & LinhaFinalCritério_3)

    Range("K" & LinhaInicialCritério_4 & ":" & "K" & LinhaFinalCritério_4).Copy _
    Destination:=Range("O" & LinhaInicialCritério_4 & ":" & "O" & LinhaFinalCritério_4)
    
    Range("K" & LinhaInicialCritério_5 & ":" & "K" & LinhaFinalCritério_5).Copy _
    Destination:=Range("P" & LinhaInicialCritério_5 & ":" & "P" & LinhaFinalCritério_5)
   
    Range("K" & LinhaInicialCritério_6 & ":" & "K" & LinhaFinalCritério_6).Copy _
    Destination:=Range("Q" & LinhaInicialCritério_6 & ":" & "Q" & LinhaFinalCritério_6)
    
    Workbooks("AFDA Report Tool.xlsm").Activate
    
    Call Clear_Report
    
    i = 75
    pctCompl = i
    progress pctCompl

    Range("A10:R10").Copy
    Range(Range("A10:R10"), Range("A10:R" & LinhaFinalCritério_6 + 8)).PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    
    Range("A10:R" & LinhaFinalCritério_6 + 8).Value = Workbooks(ARGeral).Worksheets(3).Range("A2:R" & LinhaFinalCritério_6).Value
 
    

    
    Range("K" & LinhaFinalCritério_6 + 9).FormulaR1C1 = "=SUM(R10C11:R[-1]C)"
    Range("L" & LinhaFinalCritério_6 + 9).FormulaR1C1 = "=SUM(R10C12:R[-1]C)"
    Range("M" & LinhaFinalCritério_6 + 9).FormulaR1C1 = "=SUM(R10C13:R[-1]C)"
    Range("N" & LinhaFinalCritério_6 + 9).FormulaR1C1 = "=SUM(R10C14:R[-1]C)"
    Range("O" & LinhaFinalCritério_6 + 9).FormulaR1C1 = "=SUM(R10C15:R[-1]C)"
    Range("P" & LinhaFinalCritério_6 + 9).FormulaR1C1 = "=SUM(R10C16:R[-1]C)"
    Range("Q" & LinhaFinalCritério_6 + 9).FormulaR1C1 = "=SUM(R10C17:R[-1]C)"
    Range("R" & LinhaFinalCritério_6 + 9).FormulaR1C1 = "=SUM(R10C18:R[-1]C)"
    Dim Totais As Range
    Set Totais = Range("K" & LinhaFinalCritério_6 + 9 & ":R" & LinhaFinalCritério_6 + 9)
    
    Totais.Value = Totais.Value
    'Mantém apenas o resultado das fórmulas
    
    'Destaque dos Totais com formatação
    Totais.Borders(xlDiagonalDown).LineStyle = xlNone
    Totais.Borders(xlDiagonalUp).LineStyle = xlNone
    With Totais.Borders(xlEdgeLeft)
        .LineStyle = xlDouble
        .Color = -6750208
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Totais.Borders(xlEdgeTop)
        .LineStyle = xlDouble
        .Color = -6750208
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Totais.Borders(xlEdgeBottom)
        .LineStyle = xlDouble
        .Color = -6750208
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Totais.Borders(xlEdgeRight)
        .LineStyle = xlDouble
        .Color = -6750208
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Totais.Borders(xlInsideVertical)
        .LineStyle = xlDouble
        .Color = -6750208
        .TintAndShade = 0
        .Weight = xlThick
    End With
    Totais.Borders(xlInsideHorizontal).LineStyle = xlNone
    Totais.Font.Bold = True

    i = 100
    pctCompl = i
    progress pctCompl
    If UserForm_Idioma.ToggleButton_Português = True Then
    UserForm_Processando.Label_Processando.Caption = "Concluído!"
    Else
    UserForm_Processando.Label_Processando.Caption = "Hecho!"
    End If
    
    Application.ScreenUpdating = True
    
        Application.DisplayAlerts = False

        Workbooks(ARGeral).Close

        Application.DisplayAlerts = True
    
    Exit Sub
    
Aging_File_Name_Error:

Unload UserForm_Processando
If UserForm_Idioma.ToggleButton_Português.Value = True Then
    MsgBox "Certifique-se de que o arquivo Aging está aberto, com edição habilitada e nomeado como """ & ARGeral & """(Manual Página 4)", vbOKOnly, "Aging não Reconhecido"
    Else
    MsgBox "Asegúrese de que el archivo Aging esté abierto, con edición habilitada y nombrado como """ & ARGeral & """(Manual Página 4)", vbOKOnly, "Aging no Reconocido"
    End If
    Exit Sub

Tool_File_Name_Error:
Unload UserForm_Processando

If UserForm_Idioma.ToggleButton_Português.Value = True Then
    MsgBox "Certifique-se de que o arquivo da ferramenta está nomeado como ""AFDA Report Tool.xlsm""", vbOKOnly, "Aging não Reconhecido"
    Else
    MsgBox "Asegúrese de que el archivo de la herramienta esté nombrado como ""AFDA Report Tool.xlsm""", vbOKOnly, "Aging no Reconocido"
End If

End Sub



