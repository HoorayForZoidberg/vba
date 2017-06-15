Sub PrepForMAX()
'
' (by Nick Lizop)
' Formatting for publishing to MAX
'

'let's make extra super sure the user didn't just click this by accident
Dim answer As Integer
    answer = MsgBox("Are you absolutley sure you want to format this document for MAX?", vbYesNo + vbExclamation, "Format for MAX")

If answer = vbNo Then
    Exit Sub
End If

'proceed
Application.ScreenUpdating = False
Application.EnableEvents = False

    'convert everything to values
    Sheets("SRR").Select
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    'convert everything to values
    Sheets("PRR").Select
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    'convert everything to values
    Sheets("PKI").Select
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False

    'convert everything to values
    Sheets("PRIVPKI").Select
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    'some general formatting
    Sheets("PKI").Select
    Range("A2:C2").Select
    Selection.Merge
    Range("A3:I3").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Columns("J:J").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlToLeft
    Rows("72:72").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    
    'same as above
    Sheets("PRIVPKI").Select
    Range("A2:B2").Select
    Selection.Merge
    Columns("J:J").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlToLeft
    
    'delete all other worksheets
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        Select Case ws.Name
            Case "PKI", "PRIVPKI", "SRR", "Previous SRR", "PRR", "Previous PRR"
                'do nothing
            Case Else
                Application.DisplayAlerts = False
                ws.Delete
                Application.DisplayAlerts = True
        End Select
    Next
        
Application.ScreenUpdating = True
Application.EnableEvents = True

End Sub
Sub IMFformat()
'
' (by Nick Lizop)
' IMF table formatting
'
    Application.ScreenUpdating = False
    
    Dim rng As Range
    Dim firstColumn As Long
    Dim lastColumn As Long
    
    Set rng = Selection
    
    'improve column widths
    rng.Columns.ColumnWidth = 15
    rng.Columns(1).ColumnWidth = 50

    'set font to uniform specs
    With Selection.Font
        .Name = "Arial"
        .Size = 11
        .Bold = False
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    
    'eliminate wrap text formatting etc.
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlTop
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    Application.ScreenUpdating = True
    
    'delete all blank cells and shift remaining ones left
    Dim answer As Integer
    answer = MsgBox("Would you like to delete all blank cells in the highlighted selection?", vbYesNo + vbQuestion, "Delete Blanks")
    
    Application.ScreenUpdating = False
    
    If answer = vbYes Then

        On Error Resume Next
    
        Dim blanks As Range
        Set blanks = rng.SpecialCells(xlCellTypeBlanks)
        blanks.Delete Shift:=xlShiftToLeft
    
    End If
    
    'unsmash the cells in the first column
    Call Unsmash(rng.Columns(1))
    
    Application.ScreenUpdating = True

End Sub



Sub AddToCellMenu()
'
' (by Nick Lizop)
' adds a few custom functions to the context menu
'

    Dim ContextMenu As CommandBar
    Dim MySubMenu As CommandBarControl

    ' Delete the controls first to avoid duplicates.
    Call DeleteFromCellMenu

    ' Set ContextMenu to the Cell context menu.
    Set ContextMenu = Application.CommandBars("Cell")

    ' Add a custom submenu with three buttons.
    Set MySubMenu = ContextMenu.Controls.Add(Type:=msoControlPopup, before:=3)

    With MySubMenu
        .Caption = "Change Data"
        .Tag = "My_Cell_Control_Tag"

        With .Controls.Add(Type:=msoControlButton)
            .OnAction = "'" & ThisWorkbook.Name & "'!" & "toPercent"
            .FaceId = 383
            .Caption = "Change to %"
        End With
        With .Controls.Add(Type:=msoControlButton)
            .OnAction = "'" & ThisWorkbook.Name & "'!" & "multiply1000"
            .FaceId = 376
            .Caption = "Multiply by 1000"
        End With
        With .Controls.Add(Type:=msoControlButton)
            .OnAction = "'" & ThisWorkbook.Name & "'!" & "divide1000"
            .FaceId = 377
            .Caption = "Divide by 1000"
        End With
    End With

    ' Add separators to the Cell context menu.
    ContextMenu.Controls(3).BeginGroup = True
    ContextMenu.Controls(4).BeginGroup = True

End Sub

Sub DeleteFromCellMenu()
    Dim ContextMenu As CommandBar
    Dim ctrl As CommandBarControl

    ' Set ContextMenu to the Cell context menu.
    Set ContextMenu = Application.CommandBars("Cell")

    ' Delete the custom controls with the Tag : My_Cell_Control_Tag.
    For Each ctrl In ContextMenu.Controls
        If ctrl.Tag = "My_Cell_Control_Tag" Then
            ctrl.Delete
        End If
    Next ctrl

    ' Delete the custom built-in Save button.
    On Error Resume Next
    ContextMenu.FindControl(ID:=3).Delete
    On Error GoTo 0
End Sub

Sub toPercent()
    Dim CaseRange As Range
    Dim CalcMode As Long
    Dim cell As Range

    On Error Resume Next
    Set CaseRange = Selection
    On Error GoTo 0
    If CaseRange Is Nothing Then Exit Sub

    With Application
        CalcMode = .Calculation
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .EnableEvents = False
    End With

    For Each cell In CaseRange.Cells
        cell.NumberFormat = "0.00%"
        cell.Value = cell.Value * 0.01
    Next cell

    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = CalcMode
    End With
End Sub

Sub multiply1000()
    Dim CaseRange As Range
    Dim CalcMode As Long
    Dim cell As Range

    On Error Resume Next
    Set CaseRange = Selection
    On Error GoTo 0
    If CaseRange Is Nothing Then Exit Sub

    With Application
        CalcMode = .Calculation
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .EnableEvents = False
    End With

    For Each cell In CaseRange.Cells
        cell.Value = cell.Value * 1000
    Next cell

    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = CalcMode
    End With
End Sub

Sub divide1000()
    Dim CaseRange As Range
    Dim CalcMode As Long
    Dim cell As Range

    On Error Resume Next
    Set CaseRange = Selection
    On Error GoTo 0
    If CaseRange Is Nothing Then Exit Sub

    With Application
        CalcMode = .Calculation
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .EnableEvents = False
    End With

    For Each cell In CaseRange.Cells
        cell.Value = cell.Value / 1000
    Next cell

    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = CalcMode
    End With
End Sub



