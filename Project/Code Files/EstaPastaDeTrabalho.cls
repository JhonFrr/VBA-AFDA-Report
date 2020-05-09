Private Sub Workbook_Open()

Load UserForm_Idioma
If Worksheets(1).Range("B9").Value = "RAZÓN SOCIAL" Then
UserForm_Idioma.ToggleButton_Español.Value = True

Call Restore_Layout

MsgBox "ATAJOS:" & vbNewLine & _
"CTRL+M: ABRIR MENÚ" & vbNewLine & _
"CTRL+L: BORRAR INFORME" & vbNewLine & _
"CTRL+R: RECUPERAR LAYOUT"

Else

Call Restore_Layout

MsgBox "ATALHOS:" & vbNewLine & _
"CTRL+M: ABRIR MENU" & vbNewLine & _
"CTRL+L: LIMPAR RELATÓRIO" & vbNewLine & _
"CTRL+R: RECUPERAR LAYOUT"
End If

UserForm_Menu.Show

End Sub

