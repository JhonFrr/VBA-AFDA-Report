VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_Idioma 
   Caption         =   "Idioma"
   ClientHeight    =   3525
   ClientLeft      =   225
   ClientTop       =   1365
   ClientWidth     =   6270
   OleObjectBlob   =   "UserForm_Idioma.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm_Idioma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Activate()

If UserForm_Menu.CommandButton_Start.Caption = "Iniciar" Then
ToggleButton_Portugu�s.Value = True
End If

If UserForm_Menu.CommandButton_Start.Caption = "Empezar" Then
ToggleButton_Espa�ol.Value = True
End If

End Sub
Private Sub ToggleButton_Portugu�s_Click()

If ToggleButton_Portugu�s.Value = True Then
Application.ScreenUpdating = False
ToggleButton_Portugu�s.Locked = True
ToggleButton_Espa�ol.Locked = False
ToggleButton_Espa�ol.Value = False
Label_Portugu�s.ForeColor = 65280
Label_Espa�ol.ForeColor = 16777215

Call Restore_Layout

End If

End Sub


Private Sub ToggleButton_Espa�ol_Click()

If ToggleButton_Espa�ol.Value = True Then

Application.ScreenUpdating = False
ToggleButton_Espa�ol.Locked = True
ToggleButton_Portugu�s.Locked = False
ToggleButton_Portugu�s.Value = False
Label_Espa�ol.ForeColor = 8421631
Label_Portugu�s.ForeColor = 16777215

Call Restore_Layout

End If

End Sub

Private Sub CommandButton_Home_Click()

UserForm_Idioma.Hide
UserForm_Menu.Show

End Sub



