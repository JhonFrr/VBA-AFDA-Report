VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_Settings 
   Caption         =   "Configuração"
   ClientHeight    =   3530
   ClientLeft      =   300
   ClientTop       =   1170
   ClientWidth     =   6255
   OleObjectBlob   =   "UserForm_Settings.frx":0000
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "UserForm_Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CommandButton_Back_Click()
UserForm_Settings.Hide
UserForm_RelatórioPDD.Show
End Sub

Private Sub UserForm_Activate()

If UserForm_Idioma.ToggleButton_Español.Value = True Then
Application.ScreenUpdating = False
Label1.Caption = "Ingresa abajo el nombre del archivo" & vbNewLine & _
"aging o renómbralo para:"
Application.ScreenUpdating = True
End If

If UserForm_Idioma.ToggleButton_Português.Value = True Then
Application.ScreenUpdating = False
Label1.Caption = "Digite abaixo o nome do arquivo" & vbNewLine & _
"aging ou o renomeie para:"

End If

End Sub

Private Sub CommandButton_Home_Click()

UserForm_Settings.Hide
UserForm_Menu.Show

End Sub


