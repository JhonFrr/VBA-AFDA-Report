VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EstaPastaDeTrabalho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()

Load UserForm_Idioma
If Worksheets(1).Range("B9").Value = "RAZ�N SOCIAL" Then
UserForm_Idioma.ToggleButton_Espa�ol.Value = True

Call Restore_Layout

MsgBox "ATAJOS:" & vbNewLine & _
"CTRL+M: ABRIR MEN�" & vbNewLine & _
"CTRL+L: BORRAR INFORME" & vbNewLine & _
"CTRL+R: RECUPERAR LAYOUT"

Else

Call Restore_Layout

MsgBox "ATALHOS:" & vbNewLine & _
"CTRL+M: ABRIR MENU" & vbNewLine & _
"CTRL+L: LIMPAR RELAT�RIO" & vbNewLine & _
"CTRL+R: RECUPERAR LAYOUT"
End If

UserForm_Menu.Show

End Sub

