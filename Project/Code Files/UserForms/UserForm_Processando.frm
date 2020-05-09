VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_Processando 
   Caption         =   "Indicador de Progresso"
   ClientHeight    =   3525
   ClientLeft      =   360
   ClientTop       =   1365
   ClientWidth     =   6255
   OleObjectBlob   =   "UserForm_Processando.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm_Processando"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Activate()

If UserForm_Idioma.ToggleButton_Español.Value = True Then
UserForm_Processando.Caption = "Indicador de Progreso"
Else
UserForm_Processando.Caption = "Indicador de Progresso"
End If

Application.ScreenUpdating = False
Call Get_Data
Application.ScreenUpdating = True
End Sub
