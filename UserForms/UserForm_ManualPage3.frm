VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_ManualPage3 
   Caption         =   "Manual P�gina 3 (Prepara��o do Relat�rio)"
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

If UserForm_Idioma.ToggleButton_Portugu�s.Value = True Then

Label_Title.Caption = "Prepara��o do Relat�rio"

Label_Title_MesAno.Caption = """M�s e Ano Referente"""
Label_Text_MesAno.Caption = "" & vbNewLine & _
"" & vbNewLine & _
"Seleciona o m�s e ano ao" & vbNewLine & _
"qual se refere o relat�rio" & vbNewLine & _
"(Sugere autom�ticamente" & vbNewLine & _
"o m�s anterior ao atual," & vbNewLine & _
"permitindo altera��o" & vbNewLine & _
"manual pelo usu�rio)"

Label_Title_DataFech.Caption = """Data Fechamento"""

Label_Text_DataFech.Caption = "Seleciona a data de" & vbNewLine & _
"fechamento do m�s para" & vbNewLine & _
"c�lculo dos dias vencidos." & vbNewLine & _
"(Sugere autom�ticamente" & vbNewLine & _
"o �ltimo dia �til do m�s" & vbNewLine & _
"selecionado, permitindo a" & vbNewLine & _
"altera��o manual pelo usu�rio)"

Label_Text_Config.Caption = "" & vbNewLine & _
"Bot�o " & """Configura��o""" & vbNewLine & _
"(P�gina    )"

Label_Text_Process.Caption = "" & vbNewLine & _
"Bot�o " & """Processar""" & vbNewLine & _
"Inicia o Processo"

End If

If UserForm_Idioma.ToggleButton_Espa�ol.Value = True Then

Label_Title.Caption = "Preparaci�n del Informe"

Label_Title_MesAno.Caption = """Mes y A�o de Referencia"""
Label_Text_MesAno.Caption = "" & vbNewLine & _
"" & vbNewLine & _
"Selecciona el mes y a�o al" & vbNewLine & _
"que se refiere el informe" & vbNewLine & _
"(Sugiere autom�ticamente" & vbNewLine & _
"el mes anterior al" & vbNewLine & _
"actual permitiendo el " & vbNewLine & _
"cambio por el usuario)"

Label_Title_DataFech.Caption = """Fecha Cierre"""

Label_Text_DataFech.Caption = "" & vbNewLine & _
"" & vbNewLine & _
"Selecciona la fecha" & vbNewLine & _
"cierre del mes para el" & vbNewLine & _
"c�lculo de d�as vencidos." & vbNewLine & _
"(Sugiere autom�ticamente" & vbNewLine & _
"el �ltimo d�a h�bil del mes" & vbNewLine & _
"seleccionado, permitiendo el" & vbNewLine & _
"cambio por el usuario)"

Label_Text_Config.Caption = "" & vbNewLine & _
"Bot�n " & """Configuraci�n""" & vbNewLine & _
"(P�gina    )"

Label_Text_Process.Caption = "" & vbNewLine & _
"Bot�n " & """Procesar""" & vbNewLine & _
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
