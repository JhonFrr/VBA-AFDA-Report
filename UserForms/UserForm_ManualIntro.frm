VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_ManualIntro 
   Caption         =   "Manual (Introdução)"
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

If UserForm_Idioma.ToggleButton_Português.Value = True Then
Application.ScreenUpdating = False
UserForm_ManualIntro.Caption = "Manual (Introdução)"
Label_Title.Caption = "Manual de Usuário"
Label_Intro.Caption = "Esta ferramenta foi desenvolvida com o intuito de automatizar" & vbNewLine & _
"a geração do relatório Provisão de Devedores Duvidosos." & vbNewLine & _
"O manual a seguir possui algumas informações básicas sobre" & vbNewLine & _
"sua utilização, em caso de maiores dúvidas, problemas, ou" & vbNewLine & _
"sugestões, por favor entre em contato:"
Label_Contact.Caption = "João Ferreira" & vbNewLine & _
"Credit & Collect Apprentice" & vbNewLine & _
"Avenida Dr. Chucri Zaidan, nº 1.240" & vbNewLine & _
"Torre B, 12º andar, conjuntos 1.201 e 1.204" & vbNewLine & _
"Vila São Francisco – São Paulo / SP / Brasil" & vbNewLine & _
"CEP: 04711-130" & vbNewLine & _
"T 5511 5694.8452  Tie Line 705-8452" & vbNewLine & _
"joao_ferreira@baxter.com"
Application.ScreenUpdating = True
End If

If UserForm_Idioma.ToggleButton_Español.Value = True Then
Application.ScreenUpdating = False
UserForm_ManualIntro.Caption = "Manual (Introduccion)"
Label_Title.Caption = "Manual de Usuario"
Label_Intro.Caption = "Esta herramienta fue desarrollada con el intuito de" & vbNewLine & _
"automatizar la generación del informe Provisión de Incobrables." & vbNewLine & _
"El siguiente manual tiene algunas informaciones básicas sobre" & vbNewLine & _
"su utilización, en caso de mayores dudas, problemas, o" & vbNewLine & _
"sugestiones, por favor entre en contacto:"
Label_Contact.Caption = "João Ferreira" & vbNewLine & _
"Credit & Collect Apprentice" & vbNewLine & _
"Avenida Dr. Chucri Zaidan, nº 1.240" & vbNewLine & _
"Torre B, piso 12, conjuntos 1.201 e 1.204" & vbNewLine & _
"Vila São Francisco – São Paulo / SP / Brasil" & vbNewLine & _
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
