VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_ManualPage2 
   Caption         =   "ManuaI P�gina 2 (Sele��o do Idioma)"
   ClientHeight    =   6180
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   11790
   OleObjectBlob   =   "UserForm_ManualPage2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm_ManualPage2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub UserForm_Activate()

SpinButton_ManualPage.Value = 0

If UserForm_Idioma.ToggleButton_Portugu�s.Value = True Then

Label_Title.Caption = "Sele��o do Idioma"

Label_Title_Idiomas.Caption = "Idiomas"
Label_Text_Idiomas.Caption = "" & vbNewLine & _
"" & vbNewLine & _
"A interface de usu�rio est� dispon�vel em dois idiomas," & vbNewLine & _
"a sele��o � feita aqui."

Label_Title_Home.Caption = "Bot�o " & """Home"""

Label_Text_Home.Caption = "" & vbNewLine & _
"" & vbNewLine & _
"Volta � tela do" & vbNewLine & _
"Menu Principal" & vbNewLine & _
"(Presente em" & vbNewLine & _
"todas as telas)"

End If

If UserForm_Idioma.ToggleButton_Espa�ol.Value = True Then

Label_Title.Caption = "Selecci�n del Idioma"

Label_Title_Idiomas.Caption = "Idiomas"
Label_Text_Idiomas.Caption = "" & vbNewLine & _
"" & vbNewLine & _
"La interfaz de usuario est� disponible en dos idiomas," & vbNewLine & _
"la selecci�n se realiza aqu�."

Label_Title_Home.Caption = "Bot�n " & """Home"""

Label_Text_Home.Caption = "" & vbNewLine & _
"" & vbNewLine & _
"Vuelve a la pantalla" & vbNewLine & _
"del Men� Principal" & vbNewLine & _
"(Presente en" & vbNewLine & _
"todas las pantallas)"

End If

End Sub

Private Sub SpinButton_ManualPage_Change()

If SpinButton_ManualPage.Value = 1 Then
Unload Me
UserForm_ManualPage3.Show
End If

If SpinButton_ManualPage.Value = -1 Then
Unload Me
UserForm_ManualPage1.Show
End If

End Sub

Private Sub CommandButton_Home_Click()

Unload Me
UserForm_Menu.Show

End Sub
