Attribute VB_Name = "M�dulo_Relat�rio"
Sub progress(pctCompl As Single)

UserForm_Processando.Texto.Caption = pctCompl & "% Completo"
UserForm_Processando.Barra.Width = pctCompl * 3
DoEvents

End Sub
Sub Get_Data()
    
    'Travamento de tela � ativado no procedimento que chama a macro
    '(C�digo da Userform Processando)e destivado ao final da mesma.
    
    Dim ARGeral As String
    ARGeral = UserForm_Settings.TextBox_ArquivoAging.Text
    'Cria uma vari�vel referente ao nome do arquivo aberto com os dados
    'do Aging e atribui ao texto inserido na caixa do menu op��es.
    
    
    Dim i As Integer, pctCompl As Single
    UserForm_Processando.Label_Processando.Caption = "Processando..."
    i = 0
    pctCompl = i
    progress pctCompl
    'Inicia a contabiliza��o do avan�o na barra de progresso
    
    On Error GoTo Erro_nome_do_arquivo
    Windows(ARGeral).Activate
    'Tratamento de erro para arquivo aging n�o reconhecido: caso falhe ao ativar o aging, o c�digo ser�
    'levado � uma se��o onde informa o usu�rio sobre o erro, como proceder e interrompe o procedimento.
    
    'Caso n�o haja erro, o c�digo prosseguir� com a filtragem e manipula��o dos dados do Aging.
    
    'Os procedimentos abaixo separam as colunas importantes e realizam algumas exce��es por filtragem.
    Range("A1").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$BC$34532").AutoFilter Field:=1, Criteria1:= _
        "<>1010405", Operator:=xlAnd
    ActiveSheet.Range("$A$1:$BA$34532").AutoFilter Field:=7, Criteria1:="<>IL" _
        , Operator:=xlAnd
    Range("A:A,B:B,G:G,I:I,J:J,K:K,L:L,M:M,AE:AE").Select
    Selection.Copy
    
    Sheets.Add After:=ActiveSheet
    'Adiciona uma nova planilha para jogar e organizar os dados que ser�o utilizados
    
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
    'conta as linhas preenchidas e adiciona +1 (cabe�alho) para descobrir o N� da �ltima linha preenchida.
    '(tamb�m poderia ser feito por c�lculo na mem�ria, atrav�s do
    'm�todo "worksheet.function", sem a necessidade de hospedar a f�rmula numa c�lula)
    
    Range("C2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[9]=""DST"",""DIS"",IF(OR(RC[9]=""C26"",RC[9]=""C87""),""PUB"",""PRI""))"
    'F�rmula que classifica os clientes em p�blicos e privados
    
    Selection.Copy

    
    Dim Linha As String
    Linha = Range("M1").Value
    'Cria uma vari�vel e a atribui ao n�mero da �ltima linha preenchida calculado na c�lula M1
    
    Range("C" & "2" & ":" & "C" & Linha).Select
    'Faz a sele��o de "C2" � ultima c�lula preenchida
    'Ao inv�s de range tamb�m pode ser utilizado o m�todo "Cells" que utiliza coordenadas ao inv�s
    'do nome, sem a necessidade de aninhar os textos com a vari�vel para formar o range.
    ActiveSheet.Paste
    'Cola em a f�rmula em todas as linhas com conte�do.
    
    i = 10
    pctCompl = i
    progress pctCompl
    'Barra de progresso em 10%
    
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< FECHAMENTO <<<<<<<<<<<<<<<<<<<<<<<<<<<<
    Dim DataFechamento As Date
    DataFechamento = UserForm_Relat�rioPDD.Label_DataFech.Caption '<<<<
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'Cria uma vari�vel para data de fechamento selecionada no formul�rio, que ser� usada para calcular os dias vencidos.

    Range("N1").Select
    ActiveCell.FormulaR1C1 = DataFechamento
    'insere a data na c�lula para ser usada de refer�ncia (a vari�vel tamb�m poderia ser usada direto na f�rmula).
    
    Range("N2").Select
    ActiveCell.FormulaR1C1 = "=IF((R1C14-RC[-5])>180,R1C14-RC[-5],"""")"
    'F�rmula que ser� usada para pr� selecionar somente as notas vencidas acima de 180 dias,
    'j� que todos os crit�rios s� enquadram notas acima desse n�mero.
    
    Selection.Copy
    Range("N" & "2" & ":" & "N" & Linha).Select
    ActiveSheet.Paste
    'novamente cola a f�rmula somente nas linhas com conte�do
    
    'O bloco abaixo exclui dos dados mais algumas exce��es (notas intercompany)
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
    'Remove as notas vencidas a menos de 180 dias (que ficaram com a c�lula vazia ap�s a f�rmula anterior)
    
    Columns("A:N").Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Paste
    Range("N1").Value = "Dias Vencidos"
    
    'Os blocos abaixo iniciam o processo de filtragem dos 3 primeiros crit�rios referentes � legisla��o
    'antiga (anterior � 08/07/2014) e os organizam numa outra planilha que ser� o formato final trazido no relat�rio.
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'<<<<< CRIT�RIO 1: VENCIDO > 180 DIAS AT�  R$ 5.000,00 - At� 07/10/14 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    
    'como s� restaram as notas vencidas acima de 180 dias n�o ser� necess�rui filtrar dias vencidos, somente valor e data.
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
'<<<<<<<<<<<<<<*VBA considera a o formato americano na execu��o do AutoFilter ("mm/dd/aaaa")<<<<<<<<<<<<<<<
    'Filtra a data
    
    
    Columns("A:K").Select
    Selection.Copy
    'Copia as notas que entraram no primeiro crit�rio ap�s as filtragens
    
    Sheets.Add After:=ActiveSheet
    'Adiciona a nova planilha para onde ser�o jogadas as notas que entraram em cada crit�rio ap�s os filtros.
    ActiveSheet.Paste
    'Cola o crit�rio 1
    
    Range("L1").Value = "Crit�rio"
    Range("M1").FormulaR1C1 = "=COUNTA(C[-12])"
    'Outra f�rmula par descobrir a �ltima linha preenchida.

    Dim LinhaFinalCrit�rio_1 As Long
    LinhaFinalCrit�rio_1 = Range("M1").Value
    'Cria uma vari�vel para a �ltima nota enquadrada no crit�rio 1 (�ltima linha preenchida no momento).
    
    Worksheets(3).Activate
    'Volta � planilha de filtragem para a aplica��o do pr�ximo crit�rio.
    
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'<<<<< CRIT�RIO 2: VENCIDO > 360 DIAS ACIMA DE  R$ 30.000,00 EM JU�ZO - At� 07/10/14 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<
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
    'As linhas com "L" na coluna est�o em juizo, um dos requisitos do crit�rio 2
    
    Columns("A:K").Select
    Selection.Copy
    'Copia as notas que se enquadraram no crit�rio 2.
    
    Worksheets(4).Activate
    'Ativa a planilha para qual ser�o jogadas.
    
    Range("A1").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    ActiveSheet.Paste
    'Cola as notas do crit�rio 2 ao final das do crit�rio 1.
    
    Rows(LinhaFinalCrit�rio_1 + 1 & ":" & LinhaFinalCrit�rio_1 + 1).Select
    'seleciona o cabe�alho da filtragem do crit�rio 2 que foi levado junto.
    Selection.Delete Shift:=xlUp
    'deleta o cabe�alho
    
    i = 30
    pctCompl = i
    progress pctCompl
    'Progresso 30%
    
    Dim LinhaFinalCrit�rio_2 As Long
    LinhaFinalCrit�rio_2 = Range("M1").Value
    'Cria uma vari�vel para a �ltima nota enquadrada no crit�rio 2 (�ltima linha preenchida no momento).
    
    Dim LinhaInicialCrit�rio_2 As Long
    LinhaInicialCrit�rio_2 = LinhaFinalCrit�rio_1 + 1
    'Cria uma vari�vel para a primeira nota enquadrada no crit�rio 2 (linha seguinte � �ltima do crit�rio 1)
    
    Worksheets(3).Activate
    'Volta � planilha de filtragem para a aplica��o do pr�ximo crit�rio.
    
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'<<<<< CRIT�RIO 3: VENCIDO > 360 DIAS, ACIMA DE  R$ 5.000,00 AT� R$ 30.000,00 - At� 07/10/14 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    
    ActiveSheet.ShowAllData
    'Dessa vez foi necess�rio limpar os filtros para remover a limita��o "em juizo".
    
    ActiveSheet.Range("$A:$N").AutoFilter Field:=14, Criteria1:=">360", Operator:=xlAnd
    'Filtra a quantidade de dias vencidos.
    
    ActiveSheet.Range("$A:$N").AutoFilter Field:=11, Criteria1:=">5000", Operator:=xlAnd, Criteria2:="<=30000"
    'Filtra o valor em aberto.
    
    ActiveSheet.Range("$A:$N").AutoFilter Field:=9, Criteria1:="<10/08/2014", Operator:=xlAnd
    'Filtra a data novamente.
    
    Selection.Copy
    'Como a planilha j� estava selecionada e a sele��o n�o foi alterada,
    'copia agora as notas que se enquadraram no crit�rio 3.
    
    Worksheets(4).Activate
    Range("A1").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    ActiveSheet.Paste
    Rows(LinhaFinalCrit�rio_2 + 1 & ":" & LinhaFinalCrit�rio_2 + 1).Select
    Selection.Delete Shift:=xlUp
    'Novamente lan�a a filtragem para a planilha onde est� sendo organizado o relat�rio e apaga o cabe�alho.
    
    Dim LinhaFinalCrit�rio_3 As Long
    LinhaFinalCrit�rio_3 = Range("M1").Value
    Dim LinhaInicialCrit�rio_3 As Long
    LinhaInicialCrit�rio_3 = LinhaFinalCrit�rio_2 + 1
    'Novamente s�o criadas vari�veis para linha final e inicial.
    
    
    Worksheets(3).Activate
    
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'<<<<< CRIT�RIO 4: VENCIDO > 180 DIAS AT� R$ 15.000,00 - Ap�s 07/10/14 <<<<<<<<<<<<<<<<<<<<<<<<<<  <<<
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
    Rows(LinhaFinalCrit�rio_3 + 1 & ":" & LinhaFinalCrit�rio_3 + 1).Select
    Selection.Delete Shift:=xlUp
    
    i = 40
    pctCompl = i
    progress pctCompl


    Dim LinhaFinalCrit�rio_4 As Long
    LinhaFinalCrit�rio_4 = Range("M1").Value
    Dim LinhaInicialCrit�rio_4 As Long
    LinhaInicialCrit�rio_4 = LinhaFinalCrit�rio_3 + 1
    
    
    Worksheets(3).Activate
    
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'<<<<< CRIT�RIO 5: VENCIDO > 360 DIAS, ACIMA DE  R$ 15.000,00 AT� R$ 100.000,00 - Ap�s 07/10/14 <<<<<<<<<<<<<<<<<<<<
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
    Rows(LinhaFinalCrit�rio_4 + 1 & ":" & LinhaFinalCrit�rio_4 + 1).Select
    Selection.Delete Shift:=xlUp
    Dim LinhaFinalCrit�rio_5 As Long
    LinhaFinalCrit�rio_5 = Range("M1").Value
    Dim LinhaInicialCrit�rio_5 As Long
    LinhaInicialCrit�rio_5 = LinhaFinalCrit�rio_4 + 1
    
    Worksheets(3).Activate
    
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'<<<<< CRIT�RIO 6: VENCIDO > 720 DIAS, ACIMA DE  R$ 100.000,00 - Ap�s 07/10/14 <<<<<<<<<<<<<<<<<<<<<<<<<<
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
    Rows(LinhaFinalCrit�rio_5 + 1 & ":" & LinhaFinalCrit�rio_5 + 1).Select
    Selection.Delete Shift:=xlUp
    
    i = 50
    pctCompl = i
    progress pctCompl
    

    Dim LinhaFinalCrit�rio_6 As Long
    LinhaFinalCrit�rio_6 = Range("M1").Value
    Dim LinhaInicialCrit�rio_6 As Long
    LinhaInicialCrit�rio_6 = LinhaFinalCrit�rio_5 + 1
    Columns("K:K").Select
    Selection.Copy
    Columns("R:R").Select
    ActiveSheet.Paste
    Range("K" & "2" & ":" & "K" & LinhaFinalCrit�rio_1).Select
    Selection.Copy
    Range("L" & "2" & ":" & "L" & LinhaFinalCrit�rio_1).Select
    ActiveSheet.Paste
    Range("K" & LinhaInicialCrit�rio_2 & ":" & "K" & LinhaFinalCrit�rio_2).Select
    Selection.Copy
    Range("M" & LinhaInicialCrit�rio_2 & ":" & "M" & LinhaFinalCrit�rio_2).Select
    ActiveSheet.Paste
    Range("K" & LinhaInicialCrit�rio_3 & ":" & "K" & LinhaFinalCrit�rio_3).Select
    Selection.Copy
    Range("N" & LinhaInicialCrit�rio_3 & ":" & "N" & LinhaFinalCrit�rio_3).Select
    ActiveSheet.Paste
    Range("K" & LinhaInicialCrit�rio_4 & ":" & "K" & LinhaFinalCrit�rio_4).Select
    Selection.Copy
    Range("O" & LinhaInicialCrit�rio_4 & ":" & "O" & LinhaFinalCrit�rio_4).Select
    ActiveSheet.Paste
    Range("K" & LinhaInicialCrit�rio_5 & ":" & "K" & LinhaFinalCrit�rio_5).Select
    Selection.Copy
    Range("P" & LinhaInicialCrit�rio_5 & ":" & "P" & LinhaFinalCrit�rio_5).Select
    ActiveSheet.Paste
    Range("K" & LinhaInicialCrit�rio_6 & ":" & "K" & LinhaFinalCrit�rio_6).Select
    Selection.Copy
    Range("Q" & LinhaInicialCrit�rio_6 & ":" & "Q" & LinhaFinalCrit�rio_6).Select
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
    
    Range("K" & LinhaFinalCrit�rio_6 + 9).Activate
    ActiveCell.FormulaR1C1 = "=SUM(R10C11:R[-1]C)"
    Range("L" & LinhaFinalCrit�rio_6 + 9).Activate
    ActiveCell.FormulaR1C1 = "=SUM(R10C12:R[-1]C)"
    Range("M" & LinhaFinalCrit�rio_6 + 9).Activate
    ActiveCell.FormulaR1C1 = "=SUM(R10C13:R[-1]C)"
    Range("N" & LinhaFinalCrit�rio_6 + 9).Activate
    ActiveCell.FormulaR1C1 = "=SUM(R10C14:R[-1]C)"
    Range("O" & LinhaFinalCrit�rio_6 + 9).Activate
    ActiveCell.FormulaR1C1 = "=SUM(R10C15:R[-1]C)"
    Range("P" & LinhaFinalCrit�rio_6 + 9).Activate
    ActiveCell.FormulaR1C1 = "=SUM(R10C16:R[-1]C)"
    Range("Q" & LinhaFinalCrit�rio_6 + 9).Activate
    ActiveCell.FormulaR1C1 = "=SUM(R10C17:R[-1]C)"
    Range("R" & LinhaFinalCrit�rio_6 + 9).Activate
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
    UserForm_Processando.Label_Processando.Caption = "Conclu�do!"
    Application.ScreenUpdating = True

    Windows(ARGeral).Visible = True

    
    Exit Sub
    
Erro_nome_do_arquivo:

    Unload UserForm_Processando
If UserForm_Relat�rioPDD.CommandButton_Process.Caption = "Processar" Then
    MsgBox "Certifique-se de que o Aging est� aberto e nomeado como """ & ARGeral & """(Manual P�gina 4)", vbOKOnly, "Aging n�o encontrado"
    End If
If UserForm_Relat�rioPDD.CommandButton_Process.Caption = "Procesar" Then
    MsgBox "Aseg�rese de que el archivo est� abierto y nombrado como """ & ARGeral & """(Manual P�gina 4)", vbOKOnly, "Aging no encontrado"
    End If
    
End Sub



