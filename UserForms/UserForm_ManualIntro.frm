VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_ManualIntro 
   Caption         =   "Manual (Introdu��o)"
   ClientHeight    =   6180
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   11790
   OleObjectBlob   =   "UserForm_ManualIntro.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm_ManualIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Activate()

If UserForm_Idioma.ToggleButton_Portugu�s.Value = True Then
Application.ScreenUpdating = False
UserForm_ManualIntro.Caption = "Manual (Introdu��o)"
Label_Title.Caption = "Manual de Usu�rio"
Label_Intro.Caption = "Esta ferramenta foi desenvolvida com o intuito de automatizar" & vbNewLine & _
"a gera��o do relat�rio Provis�o de Devedores Duvidosos." & vbNewLine & _
"O manual a seguir possui algumas informa��es b�sicas sobre" & vbNewLine & _
"sua utiliza��o, em caso de maiores d�vidas, problemas, ou" & vbNewLine & _
"sugest�es, por favor entre em contato:"
Label_Contact.Caption = "Jo�o Ferreira" & vbNewLine & _
"Credit & Collect Apprentice" & vbNewLine & _
"Avenida Dr. Chucri Zaidan, n� 1.240" & vbNewLine & _
"Torre B, 12� andar, conjuntos 1.201 e 1.204" & vbNewLine & _
"Vila S�o Francisco � S�o Paulo / SP / Brasil" & vbNewLine & _
"CEP: 04711-130" & vbNewLine & _
"T 5511 5694.8452  Tie Line 705-8452" & vbNewLine & _
"joao_ferreira@baxter.com"
Application.ScreenUpdating = True
End If

If UserForm_Idioma.ToggleButton_Espa�ol.Value = True Then
Application.ScreenUpdating = False
UserForm_ManualIntro.Caption = "Manual (Introduccion)"
Label_Title.Caption = "Manual de Usuario"
Label_Intro.Caption = "Esta herramienta fue desarrollada con el intuito de" & vbNewLine & _
"automatizar la generaci�n del informe Provisi�n de Incobrables." & vbNewLine & _
"El siguiente manual tiene algunas informaciones b�sicas sobre" & vbNewLine & _
"su utilizaci�n, en caso de mayores dudas, problemas, o" & vbNewLine & _
"sugestiones, por favor entre en contacto:"
Label_Contact.Caption = "Jo�o Ferreira" & vbNewLine & _
"Credit & Collect Apprentice" & vbNewLine & _
"Avenida Dr. Chucri Zaidan, n� 1.240" & vbNewLine & _
"Torre B, piso 12, conjuntos 1.201 e 1.204" & vbNewLine & _
"Vila S�o Francisco � S�o Paulo / SP / Brasil" & vbNewLine & _
"CEP: 04711-130" & vbNewLine & _
"T 5511 5694.8452  Tie Line 705-8452" & vbNewLine & _
"joao_ferreira@baxter.com"
Application.ScreenUpdating = True
End If

End Sub
Private Sub CommandButton_Next_Click()
Application.ScreenUpdating = False
Unload Me
UserForm_ManualPage1.Show
Application.ScreenUpdating = True
End Sub
Private Sub CommandButton_Home_Click()
Application.ScreenUpdating = False
Unload Me
UserForm_Menu.Show
Application.ScreenUpdating = True
End Sub
