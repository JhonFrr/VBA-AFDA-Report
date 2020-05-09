Attribute VB_Name = "Módulo_Atalhos"
Sub Menu()
Attribute Menu.VB_ProcData.VB_Invoke_Func = "m\n14"

' Atalho do teclado: Ctrl+m

UserForm_Menu.Show

End Sub
Sub Clear_Report()
Attribute Clear_Report.VB_ProcData.VB_Invoke_Func = "l\n14"

' Atalho do teclado: Ctrl+l

    Application.ScreenUpdating = False
    
    Rows("11:11").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Clear
    Rows("10:10").Select
    Selection.ClearContents
    
    Application.ScreenUpdating = True
    
End Sub

Sub Restore_Layout()
Attribute Restore_Layout.VB_ProcData.VB_Invoke_Func = "r\n14"

' Atalho do teclado: Ctrl+r

Application.ScreenUpdating = False

If UserForm_Idioma.ToggleButton_Español.Value = True Then
Worksheets("LAYOUT BACKUP").Range("A11:R20").Copy
Worksheets(1).Activate
Range("A1").Activate
ActiveSheet.Paste
Sheets(1).name = "INFORME"
Else

Worksheets("LAYOUT BACKUP").Range("A1:R10").Copy
Worksheets(1).Activate
Range("A1").Activate
ActiveSheet.Paste
Sheets(1).name = "RELATÓRIO"
End If

If Not ActiveSheet.AutoFilterMode Then
Range("A9:Q9").AutoFilter
End If

Application.ScreenUpdating = True

End Sub
Sub EditHide_LayoutBackup()
Attribute EditHide_LayoutBackup.VB_ProcData.VB_Invoke_Func = "R\n14"

' Atalho do teclado: Ctrl+Shift+r



If Worksheets("LAYOUT BACKUP").Visible = xlSheetHidden Then

Worksheets("LAYOUT BACKUP").Visible = xlSheetVisible
Worksheets("LAYOUT BACKUP").Activate

Else

Worksheets("LAYOUT BACKUP").Visible = xlSheetHidden

End If

End Sub
