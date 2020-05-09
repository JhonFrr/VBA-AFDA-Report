VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_Relat�rioPDD 
   Caption         =   "Prepara��o do Relat�rio"
   ClientHeight    =   3525
   ClientLeft      =   300
   ClientTop       =   1170
   ClientWidth     =   6255
   OleObjectBlob   =   "UserForm_Relat�rioPDD.frx":0000
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "UserForm_Relat�rioPDD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Public AnoAtual As Integer, M�sAtual As Byte, DiaFech As Byte



Private Sub UserForm_Activate()

If UserForm_Idioma.ToggleButton_Espa�ol.Value = True Then
Label1.Caption = "Por favor, confirma los datos abajo:"
Label_TitleM�sRef.Caption = "Mes y A�o de Referencia"
UserForm_Relat�rioPDD.Caption = "Preparaci�n del Informe"
Label_TitleDataFech.Caption = "Fecha Cierre"
CommandButton_Process.Caption = "Procesar"

End If

If UserForm_Idioma.ToggleButton_Portugu�s.Value = True Then
Label1.Caption = "Por favor, confirme os dados abaixo:"
Label_TitleM�sRef.Caption = "M�s e Ano Referente"
UserForm_Relat�rioPDD.Caption = "Prepara��o do Relat�rio"
Label_TitleDataFech.Caption = "Data Fechamento"
CommandButton_Process.Caption = "Processar"
End If

AnoAtual = WorksheetFunction.Text(Date, "YYYY")
M�sAtual = WorksheetFunction.Text(Date, "MM")

SpinButton_AnoRef.Value = WorksheetFunction.Text(WorksheetFunction.WorkDay(DateSerial(AnoAtual, M�sAtual, 1), -1), "YYYY")
SpinButton_M�sRef.Value = WorksheetFunction.Text(WorksheetFunction.WorkDay(DateSerial(AnoAtual, M�sAtual, 1), -1), "MM")


Label_AnoRef.Caption = WorksheetFunction.Text(WorksheetFunction.WorkDay(DateSerial(AnoAtual, M�sAtual, 1), -1), "YYYY")
Label_M�sRef.Caption = WorksheetFunction.Text(WorksheetFunction.WorkDay(DateSerial(AnoAtual, M�sAtual, 1), -1), "MM")
Label_DataFech.Caption = WorksheetFunction.Text(WorksheetFunction.WorkDay(DateSerial(AnoAtual, M�sAtual, 1), -1), "DD/MM/YYYY")


End Sub
Private Sub SpinButton_AnoRef_Change()


AnoRef = SpinButton_AnoRef.Value
M�sRef = SpinButton_M�sRef.Value

Label_DataFech.Caption = WorksheetFunction.Text(WorksheetFunction.WorkDay(DateSerial(AnoRef, M�sRef + 1, 1), -1), "DD/MM/YYYY")
Label_M�sRef.Caption = WorksheetFunction.Text(WorksheetFunction.WorkDay(DateSerial(AnoRef, M�sRef + 1, 1), -1), "MM")
Label_AnoRef.Caption = WorksheetFunction.Text(WorksheetFunction.WorkDay(DateSerial(AnoRef, M�sRef + 1, 1), -1), "YYYY")
SpinButton_DiaFech.Value = WorksheetFunction.Text(WorksheetFunction.WorkDay(DateSerial(AnoRef, M�sRef + 1, 1), -1), "DD")

End Sub
Private Sub SpinButton_M�sRef_Change()

AnoRef = SpinButton_AnoRef.Value
M�sRef = SpinButton_M�sRef.Value

Label_DataFech.Caption = WorksheetFunction.Text(WorksheetFunction.WorkDay(DateSerial(AnoRef, M�sRef + 1, 1), -1), "DD/MM/YYYY")
Label_M�sRef.Caption = WorksheetFunction.Text(WorksheetFunction.WorkDay(DateSerial(AnoRef, M�sRef + 1, 1), -1), "MM")
Label_AnoRef.Caption = WorksheetFunction.Text(WorksheetFunction.WorkDay(DateSerial(AnoRef, M�sRef + 1, 1), -1), "YYYY")
SpinButton_DiaFech.Value = WorksheetFunction.Text(WorksheetFunction.WorkDay(DateSerial(AnoRef, M�sRef + 1, 1), -1), "DD")

End Sub

Private Sub SpinButton_DiaFech_Change()

AnoRef = SpinButton_AnoRef.Value
M�sRef = SpinButton_M�sRef.Value
DiaFech = SpinButton_DiaFech.Value

Label_DataFech.Caption = WorksheetFunction.Text(DateSerial(AnoRef, M�sRef, DiaFech), "DD/MM/YYYY")

End Sub


Private Sub CommandButton_Home_Click()

UserForm_Relat�rioPDD.Hide
UserForm_Menu.Show

End Sub


Private Sub CommandButton_Process_Click()

UserForm_Relat�rioPDD.Hide
UserForm_Processando.Show

End Sub

Private Sub CommandButton_Settings_Click()

UserForm_Relat�rioPDD.Hide
UserForm_Settings.Show

End Sub

