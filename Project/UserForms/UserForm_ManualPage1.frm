VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_ManualPage1 
   Caption         =   "Manual Página 1 (Menu Principal)"
   ClientHeight    =   6180
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   11790
   OleObjectBlob   =   "UserForm_ManualPage1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm_ManualPage1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub UserForm_Activate()

SpinButton_ManualPage.Value = 0



If UserForm_Idioma.ToggleButton_Português.Value = True Then

Application.ScreenUpdating = False

UserForm_ManualPage1.Caption = "Manual Página 1 (Menu Principal)"
Label_Title_Menu.Caption = "Menu Principal"

Label_Text_Iniciar.Caption = "" & vbNewLine & _
"" & vbNewLine & _
"Tela de preparação" & vbNewLine & _
"do Relatório." & vbNewLine & _
"(Página    )"

Label_Title_Iniciar.Caption = "Botão " & """Iniciar"""

Label_Text_Manual.Caption = "" & vbNewLine & _
"" & vbNewLine & _
"Você está aqui."


Label_Title_Manual.Caption = "Botão " & """Manual"""

Label_Text_Idioma.Caption = "" & vbNewLine & _
"" & vbNewLine & _
"Tela de seleção" & vbNewLine & _
"do Idioma." & vbNewLine & _
"(Página    )"

Label_Title_Idioma.Caption = "Botão " & """Idioma"""

Application.ScreenUpdating = True

End If

If UserForm_Idioma.ToggleButton_Español.Value = True Then

Application.ScreenUpdating = False

UserForm_ManualPage1.Caption = "Manual Página 1 (Menú Principal)"
Label_Title_Menu.Caption = "Menú Principal"

Label_Text_Iniciar.Caption = "" & vbNewLine & _
"" & vbNewLine & _
"Pantalla de preparación" & vbNewLine & _
"del Informe." & vbNewLine & _
"(Página    )"

Label_Title_Iniciar.Caption = "Botón " & """Empezar"""

Label_Text_Manual.Caption = "" & vbNewLine & _
"" & vbNewLine & _
"Estás aquí."

Label_Title_Manual.Caption = "Botón " & """Manual"""

Label_Text_Idioma.Caption = "" & vbNewLine & _
"" & vbNewLine & _
"Pantalla de selección" & vbNewLine & _
"del Idioma." & vbNewLine & _
"(Página    )"

Label_Title_Idioma.Caption = "Botón " & """Idioma"""

Application.ScreenUpdating = True

End If

End Sub

Private Sub SpinButton_ManualPage_Change()

If SpinButton_ManualPage.Value = 1 Then
Unload Me
UserForm_ManualPage2.Show
End If

If SpinButton_ManualPage.Value = -1 Then
Unload Me
UserForm_ManualIntro.Show
End If

End Sub


Private Sub CommandButton_Home_Click()

Unload Me
UserForm_Menu.Show

End Sub
