Attribute VB_Name = "M�dulo_Relat�rio"
Sub progress(pctCompl As Single)

UserForm_Processando.Texto.Caption = pctCompl & "% Completo"
UserForm_Processando.Barra.Width = pctCompl * 3
DoEvents

End Sub
Sub Get_Data()
    
    'Travamento de tela � ativado no procedimento que chama a macro
    '(C�digo da Userform Processando)e destivado ao final da mesma.
    
    On Error GoTo Tool_File_Name_Error
    Workbooks("AFDA Report Tool.xlsm").Activate
    On Error GoTo 0
    'Verifica o nome do arquivo da ferramenta, pois ser� necess�rio no final para voltar e trazer os dados _
    (ao voltar poderia usar a propriedade index, por�m � imprecisa caso o usu�rio esteja usando outras planilhas)
    
    Dim ARGeral As String
    ARGeral = UserForm_Settings.TextBox_ArquivoAging.Text
    'Cria uma vari�vel referente ao nome do arquivo aberto com os dados _
    do Aging e atribui ao texto inserido na caixa do menu op��es.
    
    Dim i As Integer, pctCompl As Single
    If UserForm_Idioma.ToggleButton_Portugu�s = True Then
    UserForm_Processando.Label_Processando.Caption = "Processando..."
    Else
    UserForm_Processando.Label_Processando.Caption = "Procesando..."
    End If
    
    i = 0
    pctCompl = i
    progress pctCompl
    'Inicia a contabiliza��o do avan�o na barra de progresso
    
    On Error GoTo Aging_File_Name_Error
    Workbooks(ARGeral).Activate
    'Tratamento de erro para arquivo aging n�o reconhecido: caso falhe ao ativar o aging, o c�digo ser�
    'levado � uma se��o onde informa o usu�rio sobre o erro, como proceder e interrompe o procedimento.
    On Error GoTo 0
    'Caso n�o haja erro, o c�digo prosseguir� com a filtragem e manipula��o dos dados do Aging.
    
    'Os procedimentos abaixo separam as colunas importantes e realizam algumas exce��es por filtragem.
    Range("A1").AutoFilter
    ActiveSheet.Range("$A$1:$BC$34532").AutoFilter Field:=1, Criteria1:= _
        "<>1010405", Operator:=xlAnd
    ActiveSheet.Range("$A$1:$BA$34532").AutoFilter Field:=7, Criteria1:="<>IL" _
        , Operator:=xlAnd
    
    'Adiciona uma nova planilha para "jogar" e organizar somente os dados que ser�o utilizados
    Sheets.Add After:=ActiveSheet
    
    Worksheets(1).Activate
    Range("A:A,B:B,G:G,I:I,J:J,K:K,L:L,M:M,AE:AE").Copy Destination:=Worksheets(2).Range("A1")
    Columns("P:P").Copy Destination:=Worksheets(2).Columns("J:J")
  
    Worksheets(2).Activate
    Columns("C:C").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("C1").Value = "Tipo"
    Worksheets(1).Columns("Z:Z").Copy Destination:=Worksheets(2).Columns("L:L")
    
    'conta as linhas preenchidas e adiciona para descobrir o N� da �ltima linha.
    '(tamb�m pode ser feito por c�lculo na mem�ria, atrav�s do
    'm�todo "WorksheetFunction", sem a necessidade de "hospedar" a f�rmula numa c�lula)
    
    Dim NLinha As String
    NLinha = WorksheetFunction.CountA(Range("A:A"))
    'Cria uma vari�vel e a atribui ao n�mero da �ltima linha preenchida
   
    Range("C2:C" & NLinha).FormulaR1C1 = _
        "=IF(RC[9]=""DST"",""DIS"",IF(OR(RC[9]=""C26"",RC[9]=""C87""),""PUB"",""PRI""))"
    'F�rmula que classifica os clientes em p�blicos e privados inserida na coluna ate a ultima linha
    Range("C2:C" & NLinha).Value = Range("C2:C" & NLinha).Value
    'Mantem somente o valor resultante da formula, como n�o precisar� ser atualizada isso economiza processamento
    
    i = 10
    pctCompl = i
    progress pctCompl
    'Avan�a a barra de progresso
    
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< FECHAMENTO <<<<<<<<<<<<<<<<<<<<<<<<<<<<
    Dim DataFechamento As Date                                    '<<<<
    DataFechamento = UserForm_Relat�rioPDD.Label_DataFech.Caption '<<<<
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'Cria uma vari�vel para data de fechamento selecionada no formul�rio, que ser� usada para calcular os dias vencidos.
    Range("N1").Value = DataFechamento
    Range("N" & "2" & ":" & "N" & NLinha).FormulaR1C1 = "=IF((R1C14-RC[-5])>180,R1C14-RC[-5],"""")"
    'F�rmula que ser� usada para separar somente as notas vencidas acima de 180 dias,
    'pois somente � partir da� entram em algum crit�rio de PDD
    Range("N" & "2" & ":" & "N" & NLinha).Value = Range("N" & "2" & ":" & "N" & NLinha).Value
    'Mantem somente o valor resultante da formula, como n�o precisar� ser atualizada isso economiza processamento
    'O bloco abaixo exclui dos dados mais algumas exce��es (notas intercompany)
    
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
    'Separa somente as notas vencidas a mais de 180 dias (excluindo as que ficaram vazias (vazio = crit�rio "falso" na f�rmula))
    With Worksheets(2).AutoFilter.Range
    Dim MenorIgua180 As Range
    Set MenorIgua180 = Range(Range("A" & .Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row), Range("N" & .Offset(1, 0).SpecialCells(xlCellTypeVisible)(1).Row).End(xlDown))
    End With
    MenorIgua180.Delete Shift:=xlUp
    ActiveSheet.ShowAllData
     
    Range("N1").Value = "Dias Vencidos"
    
    'Os blocos abaixo iniciam o processo de filtragem dos 3 primeiros crit�rios referentes � legisla��o
    'antiga (anterior � 08/07/2014) e os organizam numa outra planilha que ser� o formato final trazido no relat�rio.
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'<<<<< CRIT�RIO 1: VENCIDO > 180 DIAS AT�  R$ 5.000,00 - At� 07/10/14 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    
    'como s� restaram as notas vencidas acima de 180 dias n�o ser� necess�rui filtrar dias vencidos, somente valor e data.
    ActiveSheet.Range("$A:$N").AutoFilter Field:=11, Criteria1:="<5000" _
        , Operator:=xlAnd
    'Filtra o valor
    
    i = 20
    pctCompl = i
    progress pctCompl
    'Barra de progresso 20%
    

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    ActiveSheet.Range("A:N").AutoFilter Field:=9, Criteria1:="<10/08/2014", Operator:=xlAnd '<<<<<
                            '*MENOR QUE 08/07 � O MESMO QUE MENOR IGUAL A 07/10             '<<<<<
'<<<<<<<<<<<<<<*VBA considera a o formato americano("mm/dd/aaaa")<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    'Filtra a data de vencimento
    
    Sheets.Add After:=ActiveSheet
    'Adiciona a nova planilha para onde ser�o jogadas as notas que entrarno relat�rio.
    
    Worksheets(2).Columns("A:K").Copy Destination:=Worksheets(3).Range("A1")
    'Copia as notas que entraram no primeiro crit�rio ap�s as filtragens e cola na nova planilha.

    Range("L1").Value = "Crit�rio"

    Dim LinhaFinalCrit�rio_1 As Long
    LinhaFinalCrit�rio_1 = WorksheetFunction.CountA(Range("A:A"))
    'Cria uma vari�vel para a �ltima nota enquadrada no crit�rio 1 (�ltima linha preenchida no momento).
    
    Worksheets(2).Activate
    'Volta � planilha de filtragem para a aplica��o do pr�ximo crit�rio.
    
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'<<<<< CRIT�RIO 2: VENCIDO > 360 DIAS ACIMA DE  R$ 30.000,00 EM JU�ZO - At� 07/10/14 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<
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
    'As linhas com "L" na coluna est�o em juizo, um dos requisitos do crit�rio 2
    End With
    
    Columns("A:K").Copy Destination:=Worksheets(3).Range("A1").End(xlDown).Offset(1, 0)
    'Copia as notas que se enquadraram no crit�rio 2 e cola ao final das do crit�rio 1.
    
    i = 30
    pctCompl = i
    progress pctCompl
    'Progresso 30%
    
    Worksheets(3).Activate
    
    'Cria uma vari�vel para a �ltima nota enquadrada no crit�rio 2 (�ltima linha preenchida no momento).
    
    Dim LinhaInicialCrit�rio_2 As Long
    LinhaInicialCrit�rio_2 = LinhaFinalCrit�rio_1 + 1
    'Cria uma vari�vel para a primeira nota enquadrada no crit�rio 2 (linha seguinte � �ltima do crit�rio 1)
    Rows(LinhaInicialCrit�rio_2).Delete Shift:=xlUp
    'deleta o cabe�alho da filtragem do crit�rio 2 que foi levado junto e ficou abaixo da ultima do crit�rio 1.
    Dim LinhaFinalCrit�rio_2 As Long
    LinhaFinalCrit�rio_2 = WorksheetFunction.CountA(Range("A:A"))
    Worksheets(2).Activate
    'Volta � planilha de filtragem para a aplica��o do pr�ximo crit�rio.
    
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'<<<<< CRIT�RIO 3: VENCIDO > 360 DIAS, ACIMA DE  R$ 5.000,00 AT� R$ 30.000,00 - At� 07/10/14 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    
    ActiveSheet.ShowAllData
    'Dessa vez foi necess�rio limpar os filtros para remover a limita��o "em juizo".
    
    With ActiveSheet.Range("$A:$N")
    .AutoFilter Field:=14, Criteria1:=">360", Operator:=xlAnd
    'Filtra a quantidade de dias vencidos.
    
    .AutoFilter Field:=11, Criteria1:=">5000", Operator:=xlAnd, Criteria2:="<=30000"
    'Filtra o valor em aberto.
    
    .AutoFilter Field:=9, Criteria1:="<10/08/2014", Operator:=xlAnd
    'Filtra a data novamente.
    End With
    
    Columns("A:K").Copy Destination:=Worksheets(3).Range("A1").End(xlDown).Offset(1, 0)
    
    'copia agora as notas que se enquadraram no crit�rio 3 e novamente _
    lan�a a filtragem para a planilha onde est� sendo organizado o relat�rio
    
    Worksheets(3).Activate
    
    Dim LinhaInicialCrit�rio_3 As Long
    LinhaInicialCrit�rio_3 = LinhaFinalCrit�rio_2 + 1
    
    Rows(LinhaInicialCrit�rio_3).Delete Shift:=xlUp
    'Apaga o cabe�alho.
    
    Dim LinhaFinalCrit�rio_3 As Long
    LinhaFinalCrit�rio_3 = WorksheetFunction.CountA(Range("A:A"))
    'Novamente s�o criadas vari�veis para linha final e inicial.

    
    Worksheets(2).Activate
    
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'<<<<< CRIT�RIO 4: VENCIDO > 180 DIAS AT� R$ 15.000,00 - Ap�s 07/10/14 <<<<<<<<<<<<<<<<<<<<<<<<<<  <<<
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

    Dim LinhaInicialCrit�rio_4 As Long
    LinhaInicialCrit�rio_4 = LinhaFinalCrit�rio_3 + 1
    
    Rows(LinhaInicialCrit�rio_4).Delete Shift:=xlUp

    Dim LinhaFinalCrit�rio_4 As Long
    LinhaFinalCrit�rio_4 = WorksheetFunction.CountA(Range("A:A"))
    
    Worksheets(2).Activate
    
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'<<<<< CRIT�RIO 5: VENCIDO > 360 DIAS, ACIMA DE  R$ 15.000,00 AT� R$ 100.000,00 - Ap�s 07/10/14 <<<<<<<<<<<<<<<<<<<<
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
 
    
    With ActiveSheet.Range("A:N")
    
    .AutoFilter Field:=14, Criteria1:=">360", _
        Operator:=xlAnd
    .AutoFilter Field:=11, Criteria1:=">15000" _
        , Operator:=xlAnd, Criteria2:="<=100000"
    End With
    
    Columns("A:K").Copy Destination:=Worksheets(3).Range("A1").End(xlDown).Offset(1, 0)
    
    Worksheets(3).Activate
    Dim LinhaInicialCrit�rio_5 As Long
    LinhaInicialCrit�rio_5 = LinhaFinalCrit�rio_4 + 1
  
    Rows(LinhaInicialCrit�rio_5).Delete Shift:=xlUp
    
    Dim LinhaFinalCrit�rio_5 As Long
    LinhaFinalCrit�rio_5 = WorksheetFunction.CountA(Range("A:A"))
      
    Worksheets(2).Activate
    
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'<<<<< CRIT�RIO 6: VENCIDO > 720 DIAS, ACIMA DE  R$ 100.000,00 - Ap�s 07/10/14 <<<<<<<<<<<<<<<<<<<<<<<<<<
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

    Dim LinhaInicialCrit�rio_6 As Long
    LinhaInicialCrit�rio_6 = LinhaFinalCrit�rio_5 + 1
    
    Rows(LinhaInicialCrit�rio_6).Delete Shift:=xlUp
    
    Dim LinhaFinalCrit�rio_6 As Long
    LinhaFinalCrit�rio_6 = WorksheetFunction.CountA(Range("A:A"))

    Columns("K:K").Copy Destination:=Columns("R:R")

    Range("K2:K" & LinhaFinalCrit�rio_1).Copy Destination:=Range("L2:L" & LinhaFinalCrit�rio_1)
    
    Range("K" & LinhaInicialCrit�rio_2 & ":K" & LinhaFinalCrit�rio_2).Copy _
    Destination:=Range("M" & LinhaInicialCrit�rio_2 & ":" & "M" & LinhaFinalCrit�rio_2)
    
    Range("K" & LinhaInicialCrit�rio_3 & ":" & "K" & LinhaFinalCrit�rio_3).Copy _
    Destination:=Range("N" & LinhaInicialCrit�rio_3 & ":" & "N" & LinhaFinalCrit�rio_3)

    Range("K" & LinhaInicialCrit�rio_4 & ":" & "K" & LinhaFinalCrit�rio_4).Copy _
    Destination:=Range("O" & LinhaInicialCrit�rio_4 & ":" & "O" & LinhaFinalCrit�rio_4)
    
    Range("K" & LinhaInicialCrit�rio_5 & ":" & "K" & LinhaFinalCrit�rio_5).Copy _
    Destination:=Range("P" & LinhaInicialCrit�rio_5 & ":" & "P" & LinhaFinalCrit�rio_5)
   
    Range("K" & LinhaInicialCrit�rio_6 & ":" & "K" & LinhaFinalCrit�rio_6).Copy _
    Destination:=Range("Q" & LinhaInicialCrit�rio_6 & ":" & "Q" & LinhaFinalCrit�rio_6)
    
    Workbooks("AFDA Report Tool.xlsm").Activate
    
    Call Clear_Report
    
    i = 75
    pctCompl = i
    progress pctCompl

    Range("A10:R10").Copy
    Range(Range("A10:R10"), Range("A10:R" & LinhaFinalCrit�rio_6 + 8)).PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    
    Range("A10:R" & LinhaFinalCrit�rio_6 + 8).Value = Workbooks(ARGeral).Worksheets(3).Range("A2:R" & LinhaFinalCrit�rio_6).Value
 
    

    
    Range("K" & LinhaFinalCrit�rio_6 + 9).FormulaR1C1 = "=SUM(R10C11:R[-1]C)"
    Range("L" & LinhaFinalCrit�rio_6 + 9).FormulaR1C1 = "=SUM(R10C12:R[-1]C)"
    Range("M" & LinhaFinalCrit�rio_6 + 9).FormulaR1C1 = "=SUM(R10C13:R[-1]C)"
    Range("N" & LinhaFinalCrit�rio_6 + 9).FormulaR1C1 = "=SUM(R10C14:R[-1]C)"
    Range("O" & LinhaFinalCrit�rio_6 + 9).FormulaR1C1 = "=SUM(R10C15:R[-1]C)"
    Range("P" & LinhaFinalCrit�rio_6 + 9).FormulaR1C1 = "=SUM(R10C16:R[-1]C)"
    Range("Q" & LinhaFinalCrit�rio_6 + 9).FormulaR1C1 = "=SUM(R10C17:R[-1]C)"
    Range("R" & LinhaFinalCrit�rio_6 + 9).FormulaR1C1 = "=SUM(R10C18:R[-1]C)"
    Dim Totais As Range
    Set Totais = Range("K" & LinhaFinalCrit�rio_6 + 9 & ":R" & LinhaFinalCrit�rio_6 + 9)
    
    Totais.Value = Totais.Value
    'Mant�m apenas o resultado das f�rmulas
    
    'Destaque dos Totais com formata��o
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
    If UserForm_Idioma.ToggleButton_Portugu�s = True Then
    UserForm_Processando.Label_Processando.Caption = "Conclu�do!"
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
If UserForm_Idioma.ToggleButton_Portugu�s.Value = True Then
    MsgBox "Certifique-se de que o arquivo Aging est� aberto, com edi��o habilitada e nomeado como """ & ARGeral & """(Manual P�gina 4)", vbOKOnly, "Aging n�o Reconhecido"
    Else
    MsgBox "Aseg�rese de que el archivo Aging est� abierto, con edici�n habilitada y nombrado como """ & ARGeral & """(Manual P�gina 4)", vbOKOnly, "Aging no Reconocido"
    End If
    Exit Sub

Tool_File_Name_Error:
Unload UserForm_Processando

If UserForm_Idioma.ToggleButton_Portugu�s.Value = True Then
    MsgBox "Certifique-se de que o arquivo da ferramenta est� nomeado como ""AFDA Report Tool.xlsm""", vbOKOnly, "Aging n�o Reconhecido"
    Else
    MsgBox "Aseg�rese de que el archivo de la herramienta est� nombrado como ""AFDA Report Tool.xlsm""", vbOKOnly, "Aging no Reconocido"
End If

End Sub



