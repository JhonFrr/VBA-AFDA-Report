VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_ManualPage1 
   Caption         =   "Manual P�gina 1 (Menu Principal)"
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



If UserForm_Idioma.ToggleButton_Portugu�s.Value = True Then

Application.ScreenUpdating = False

UserForm_ManualPage1.Caption = "Manual P�gina 1 (Menu Principal)"
Label_Title_Menu.Caption = "Menu Principal"

Label_Text_Iniciar.Caption = "" & vbNewLine & _
"" & vbNewLine & _
"Tela de prepara��o" & vbNewLine & _
"do Relat�rio." & vbNewLine & _
"(P�gina    )"

Label_Title_Iniciar.Caption = "Bot�o " & """Iniciar"""

Label_Text_Manual.Caption = "" & vbNewLine & _
"" & vbNewLine & _
"Voc� est� aqui."


Label_Title_Manual.Caption = "Bot�o " & """Manual"""

Label_Text_Idioma.Caption = "" & vbNewLine & _
"" & vbNewLine & _
"Tela de sele��o" & vbNewLine & _
"do Idioma." & vbNewLine & _
"(P�gina    )"

Label_Title_Idioma.Caption = "Bot�o " & """Idioma"""

Application.ScreenUpdating = True

End If

If UserForm_Idioma.ToggleButton_Espa�ol.Value = True Then

Application.ScreenUpdating = False

UserForm_ManualPage1.Caption = "Manual P�gina 1 (Men� Principal)"
Label_Title_Menu.Caption = "Men� Principal"

Label_Text_Iniciar.Caption = "" & vbNewLine & _
"" & vbNewLine & _
"Pantalla de preparaci�n" & vbNewLine & _
"del Informe." & vbNewLine & _
"(P�gina    )"

Label_Title_Iniciar.Caption = "Bot�n " & """Empezar"""

Label_Text_Manual.Caption = "" & vbNewLine & _
"" & vbNewLine & _
"Est�s aqu�."

Label_Title_Manual.Caption = "Bot�n " & """Manual"""

Label_Text_Idioma.Caption = "" & vbNewLine & _
"" & vbNewLine & _
"Pantalla de selecci�n" & vbNewLine & _
"del Idioma." & vbNewLine & _
"(P�gina    )"

Label_Title_Idioma.Caption = "Bot�n " & """Idioma"""

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
