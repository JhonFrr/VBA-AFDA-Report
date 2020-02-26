VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_Menu 
   Caption         =   "Menu"
   ClientHeight    =   3525
   ClientLeft      =   240
   ClientTop       =   930
   ClientWidth     =   6255
   OleObjectBlob   =   "UserForm_Menu.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm_Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Activate()

Application.ScreenUpdating = False

If UserForm_Idioma.ToggleButton_Português.Value = True Then
Label1.Caption = "Relatório" & vbNewLine & "Provisão de Devedores Duvidosos"
CommandButton_Start.Caption = "Iniciar"
UserForm_Menu.Caption = "Menu"
End If

If UserForm_Idioma.ToggleButton_Español.Value = True Then
Label1.Caption = "Informe" & vbNewLine & "Provisión de Incobrables"
CommandButton_Start.Caption = "Empezar"
UserForm_Menu.Caption = "Menú"
End If

Application.ScreenUpdating = True

End Sub
Private Sub CommandButton_Idioma_Click()

UserForm_Menu.Hide
UserForm_Idioma.Show

End Sub

Private Sub CommandButton_Manual_Click()

UserForm_Menu.Hide
UserForm_ManualIntro.Show

End Sub

Private Sub CommandButton_Start_Click()

UserForm_Menu.Hide
UserForm_RelatórioPDD.Show

End Sub

