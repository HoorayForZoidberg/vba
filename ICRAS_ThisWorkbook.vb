Private Sub Workbook_Activate()
    Call AddToCellMenu
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
'
' by Doug Adams
'

Application.StatusBar = "................So long, and thanks for all the fish!................."
If Application.Wait(Now + TimeValue("0:00:02")) Then
    Application.StatusBar = False
    Exit Sub
End If

End Sub

Private Sub Workbook_Deactivate()
    Call DeleteFromCellMenu
End Sub

Private Sub Workbook_Open()
'
' (by Nick Lizop)
' prompt the user to save his file in an appropriate format
' note: for simplicity, this program relies on the naming convention of templates as "fy + [2 digit end year]"
' this method will stop working in 2030, by which time I really hope someone will have taken a closer look at this file and updated my code
'
    If InStr(ActiveWorkbook.Name, ".xlt") Or InStr(ActiveWorkbook.Name, "fy1") Or InStr(ActiveWorkbook.Name, "fy2") Then
        Application.Dialogs(xlDialogSaveAs).Show ("Please save this file as a macro enabled workbook")
    End If
End Sub