Function Map(F As String, ByVal A As Variant) As Variant
' Map function F across 1-Dimensional Array

Dim i As Long

    For i = LBound(A) To UBound(A)
        A(i) = Application.Run(F, A(i))
    Next i
    
    Map = A
    
End Function

Function Unsmash(rng As Range)
'
' (by Nick Lizop)
' Recursively unsmashes cells which contain multiple lines
'
    
    For Each cell In rng.Cells
        
        If InStr(1, cell.Value2, Chr(10), vbTextCompare) Then
            
            Dim arr() As String
            
            arr = Split(cell.Value2, Chr(10), 2, vbTextCompare)
            
            cell.Value2 = arr(1)
            
            With cell.Font
                .Name = "Arial"
                .Size = 11
                .Bold = False
            End With
            
            cell.EntireRow.Insert
            
            cell.Offset(-1, 0).Select
            
            With ActiveCell
                .Value2 = arr(0)
                .Font.Name = "Arial"
                .Font.Size = 11
                .Font.Bold = False
            End With
            
            'recursively call the function until no two lines are left in the same cell
            Call Unsmash(ActiveCell.Offset(1, 0))
            
        End If
        
    Next

End Function
