Option Explicit

Sub KopirovatDataZSouboru()
    Dim wbTarget As Workbook
    Dim wsTarget As Worksheet
    Dim targetRow As Long
    Dim i As Long
    Dim max_i As Long
    Dim cesta As String
    Dim fileName As String

' určím si cílový sešit excelu a list, na který se budou vkládat data
      Set wbTarget = ThisWorkbook
      Set wsTarget = wbTarget.Sheets("PhxDB")
      max_i = PhxDB.Range("O23")
      
' Vyhledání první prázdné buňky ve sloupci A
    targetRow = wsTarget.Cells(wsTarget.Rows.Count, "A").End(xlUp).Row + 1

    ' Iterace přes všechny soubory
      For i = 1 To max_i
        fileName = "C:\Users\josef.sedlacek\OneDrive - Trelleborg AB\Desktop\phoenix\phx (" & i & ").xlsx"
        Workbooks.Open fileName
        
      With ActiveWorkbook.Sheets(1) ' předpokládá se, že data jsou v prvním listu
            wsTarget.Cells(targetRow, 1).Value = .Range("B24").Value
            wsTarget.Cells(targetRow, 2).Value = .Range("B25").Value
            wsTarget.Cells(targetRow, 3).Value = .Range("B26").Value
            wsTarget.Cells(targetRow, 4).Value = .Range("B27").Value
            wsTarget.Cells(targetRow, 5).Value = .Range("B28").Value
            wsTarget.Cells(targetRow, 6).Value = .Range("B29").Value
            wsTarget.Cells(targetRow, 7).Value = .Range("B30").Value
            wsTarget.Cells(targetRow, 8).Value = .Range("B31").Value
            wsTarget.Cells(targetRow, 9).Value = .Range("B32").Value
            wsTarget.Cells(targetRow, 10).Value = .Range("B33").Value
            wsTarget.Cells(targetRow, 11).Value = .Range("B34").Value
            wsTarget.Cells(targetRow, 12).Value = .Range("B35").Value
      End With
                
        ' Uzavření souboru
        ActiveWorkbook.Close False
        
        ' Přesun na další prázdný řádek
        targetRow = wsTarget.Cells(wsTarget.Rows.Count, "A").End(xlUp).Row + 1
    Next i
    
    MsgBox "Data-mining proběhl úspěšně, kámo"
    
End Sub
