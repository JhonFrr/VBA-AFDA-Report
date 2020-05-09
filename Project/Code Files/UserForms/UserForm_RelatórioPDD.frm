VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_RelatórioPDD 
   Caption         =   "Preparação do Relatório"
   ClientHeight    =   3525
   ClientLeft      =   300
   ClientTop       =   1170
   ClientWidth     =   6255
   OleObjectBlob   =   "UserForm_RelatórioPDD.frx":0000
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "UserForm_RelatórioPDD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Public AnoAtual As Integer, MêsAtual As Byte, DiaFech As Byte



Private Sub UserForm_Activate()

If UserForm_Idioma.ToggleButton_Español.Value = True Then
Label1.Caption = "Por favor, confirma los datos abajo:"
Label_TitleMêsRef.Caption = "Mes y Año de Referencia"
UserForm_RelatórioPDD.Caption = "Preparación del Informe"
Label_TitleDataFech.Caption = "Fecha Cierre"
CommandButton_Process.Caption = "Procesar"

End If

If UserForm_Idioma.ToggleButton_Português.Value = True Then
Label1.Caption = "Por favor, confirme os dados abaixo:"
Label_TitleMêsRef.Caption = "Mês e Ano Referente"
UserForm_RelatórioPDD.Caption = "Preparação do Relatório"
Label_TitleDataFech.Caption = "Data Fechamento"
CommandButton_Process.Caption = "Processar"
End If

AnoAtual = WorksheetFunction.Text(Date, "YYYY")
MêsAtual = WorksheetFunction.Text(Date, "MM")

SpinButton_AnoRef.Value = WorksheetFunction.Text(WorksheetFunction.WorkDay(DateSerial(AnoAtual, MêsAtual, 1), -1), "YYYY")
SpinButton_MêsRef.Value = WorksheetFunction.Text(WorksheetFunction.WorkDay(DateSerial(AnoAtual, MêsAtual, 1), -1), "MM")


Label_AnoRef.Caption = WorksheetFunction.Text(WorksheetFunction.WorkDay(DateSerial(AnoAtual, MêsAtual, 1), -1), "YYYY")
Label_MêsRef.Caption = WorksheetFunction.Text(WorksheetFunction.WorkDay(DateSerial(AnoAtual, MêsAtual, 1), -1), "MM")
Label_DataFech.Caption = WorksheetFunction.Text(WorksheetFunction.WorkDay(DateSerial(AnoAtual, MêsAtual, 1), -1), "DD/MM/YYYY")


End Sub
Private Sub SpinButton_AnoRef_Change()


AnoRef = SpinButton_AnoRef.Value
MêsRef = SpinButton_MêsRef.Value

Label_DataFech.Caption = WorksheetFunction.Text(WorksheetFunction.WorkDay(DateSerial(AnoRef, MêsRef + 1, 1), -1), "DD/MM/YYYY")
Label_MêsRef.Caption = WorksheetFunction.Text(WorksheetFunction.WorkDay(DateSerial(AnoRef, MêsRef + 1, 1), -1), "MM")
Label_AnoRef.Caption = WorksheetFunction.Text(WorksheetFunction.WorkDay(DateSerial(AnoRef, MêsRef + 1, 1), -1), "YYYY")
SpinButton_DiaFech.Value = WorksheetFunction.Text(WorksheetFunction.WorkDay(DateSerial(AnoRef, MêsRef + 1, 1), -1), "DD")

End Sub
Private Sub SpinButton_MêsRef_Change()

AnoRef = SpinButton_AnoRef.Value
MêsRef = SpinButton_MêsRef.Value

Label_DataFech.Caption = WorksheetFunction.Text(WorksheetFunction.WorkDay(DateSerial(AnoRef, MêsRef + 1, 1), -1), "DD/MM/YYYY")
Label_MêsRef.Caption = WorksheetFunction.Text(WorksheetFunction.WorkDay(DateSerial(AnoRef, MêsRef + 1, 1), -1), "MM")
Label_AnoRef.Caption = WorksheetFunction.Text(WorksheetFunction.WorkDay(DateSerial(AnoRef, MêsRef + 1, 1), -1), "YYYY")
SpinButton_DiaFech.Value = WorksheetFunction.Text(WorksheetFunction.WorkDay(DateSerial(AnoRef, MêsRef + 1, 1), -1), "DD")

End Sub

Private Sub SpinButton_DiaFech_Change()

AnoRef = SpinButton_AnoRef.Value
MêsRef = SpinButton_MêsRef.Value
DiaFech = SpinButton_DiaFech.Value

Label_DataFech.Caption = WorksheetFunction.Text(DateSerial(AnoRef, MêsRef, DiaFech), "DD/MM/YYYY")

End Sub


Private Sub CommandButton_Home_Click()

UserForm_RelatórioPDD.Hide
UserForm_Menu.Show

End Sub


Private Sub CommandButton_Process_Click()

UserForm_RelatórioPDD.Hide
UserForm_Processando.Show

End Sub

Private Sub CommandButton_Settings_Click()

UserForm_RelatórioPDD.Hide
UserForm_Settings.Show

End Sub

