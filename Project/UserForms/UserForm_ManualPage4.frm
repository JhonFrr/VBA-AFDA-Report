VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_ManualPage4 
   Caption         =   "Manual Página 4 (Opções)"
   ClientHeight    =   6180
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   11790
   OleObjectBlob   =   "UserForm_ManualPage4.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm_ManualPage4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Activate()

If UserForm_Idioma.ToggleButton_Português.Value = True Then

Label_Title_Arquivo.Caption = """Nome do Arquivo Aging"""

Label_Arquivo.Caption = "" & vbNewLine & _
"Para gerar o relatório, o Aging atualizado" & vbNewLine & _
"deve estar aberto ao mesmo tempo, para" & vbNewLine & _
"que ele seja reconhecido digite seu" & vbNewLine & _
"nome aqui ou renomeie o arquivo."

End If

If UserForm_Idioma.ToggleButton_Español.Value = True Then

Label_Title_Arquivo.Caption = """Nombre del Archivo Aging"""

Label_Arquivo.Caption = "" & vbNewLine & _
"Para generar el informe, el Aging actualizado" & vbNewLine & _
"debe estar abierto al mismo tiempo, para" & vbNewLine & _
"que sea reconocido ingresa su" & vbNewLine & _
"nombre aquí o renombra el archivo."

End If

End Sub

Private Sub CommandButton_Back_Click()

Unload Me
UserForm_ManualPage3.Show

End Sub
Private Sub CommandButton_Home_Click()

Unload Me
UserForm_Menu.Show

End Sub
