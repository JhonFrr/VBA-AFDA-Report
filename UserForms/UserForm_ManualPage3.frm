VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_ManualPage3 
   Caption         =   "Manual Página 3 (Preparação do Relatório)"
   ClientHeight    =   6180
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   11790
   OleObjectBlob   =   "UserForm_ManualPage3.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm_ManualPage3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Label_TextMesAno_Click()

End Sub

Private Sub Label10_Click()

End Sub

Private Sub Label_Text_Process_Click()

End Sub

Private Sub Label14_Click()

End Sub

Private Sub UserForm_Activate()

SpinButton_ManualPage.Value = 0

If UserForm_Idioma.ToggleButton_Português.Value = True Then

Label_Title.Caption = "Preparação do Relatório"

Label_Title_MesAno.Caption = """Mês e Ano Referente"""
Label_Text_MesAno.Caption = "" & vbNewLine & _
"" & vbNewLine & _
"Seleciona o mês e ano ao" & vbNewLine & _
"qual se refere o relatório" & vbNewLine & _
"(Sugere automáticamente" & vbNewLine & _
"o mês anterior ao atual," & vbNewLine & _
"permitindo alteração" & vbNewLine & _
"manual pelo usuário)"

Label_Title_DataFech.Caption = """Data Fechamento"""

Label_Text_DataFech.Caption = "Seleciona a data de" & vbNewLine & _
"fechamento do mês para" & vbNewLine & _
"cálculo dos dias vencidos." & vbNewLine & _
"(Sugere automáticamente" & vbNewLine & _
"o último dia útil do mês" & vbNewLine & _
"selecionado, permitindo a" & vbNewLine & _
"alteração manual pelo usuário)"

Label_Text_Config.Caption = "" & vbNewLine & _
"Botão " & """Configuração""" & vbNewLine & _
"(Página    )"

Label_Text_Process.Caption = "" & vbNewLine & _
"Botão " & """Processar""" & vbNewLine & _
"Inicia o Processo"

End If

If UserForm_Idioma.ToggleButton_Español.Value = True Then

Label_Title.Caption = "Preparación del Informe"

Label_Title_MesAno.Caption = """Mes y Año de Referencia"""
Label_Text_MesAno.Caption = "" & vbNewLine & _
"" & vbNewLine & _
"Selecciona el mes y año al" & vbNewLine & _
"que se refiere el informe" & vbNewLine & _
"(Sugiere automáticamente" & vbNewLine & _
"el mes anterior al" & vbNewLine & _
"actual permitiendo el " & vbNewLine & _
"cambio por el usuario)"

Label_Title_DataFech.Caption = """Fecha Cierre"""

Label_Text_DataFech.Caption = "" & vbNewLine & _
"" & vbNewLine & _
"Selecciona la fecha" & vbNewLine & _
"cierre del mes para el" & vbNewLine & _
"cálculo de días vencidos." & vbNewLine & _
"(Sugiere automáticamente" & vbNewLine & _
"el último día hábil del mes" & vbNewLine & _
"seleccionado, permitiendo el" & vbNewLine & _
"cambio por el usuario)"

Label_Text_Config.Caption = "" & vbNewLine & _
"Botón " & """Configuración""" & vbNewLine & _
"(Página    )"

Label_Text_Process.Caption = "" & vbNewLine & _
"Botón " & """Procesar""" & vbNewLine & _
"Inicia el Proceso"

End If


End Sub

Private Sub SpinButton_ManualPage_Change()

If SpinButton_ManualPage.Value = 1 Then
Unload Me
UserForm_ManualPage4.Show
End If

If SpinButton_ManualPage.Value = -1 Then
Unload Me
UserForm_ManualPage2.Show
End If

End Sub

Private Sub CommandButton_Home_Click()

Unload Me
UserForm_Menu.Show

End Sub
