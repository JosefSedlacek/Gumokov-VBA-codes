'Oprávnění na listu Časová Osa
'Pouze uživatel s oprávněním může psát do buněk G24:G123
Private Sub Worksheet_Change(ByVal Target As Range)
  If Not Intersect(Target, Range("G24:G123")) Is Nothing Then
  
  Dim uzivatel As String
  uzivatel = Environ("Username")
  
  Dim uzivatel_s_opravnenim(1 To 4)
  uzivatel_s_opravnenim(1) = "josef.sedlacek"
  uzivatel_s_opravnenim(2) = "..."
  uzivatel_s_opravnenim(3) = "..."
  uzivatel_s_opravnenim(4) = "přidej dalšího uživatele"
  
        If Not (uzivatel = uzivatel_s_opravnenim(1) Or uzivatel = uzivatel_s_opravnenim(2) _
                                                    Or uzivatel = uzivatel_s_opravnenim(3) _
                                                    Or uzivatel = uzivatel_s_opravnenim(4)) Then
              Application.EnableEvents = False
              Application.Undo
              Application.EnableEvents = True
              MsgBox "Nemáte oprávnění, obraťte se na administrativu vašeho oddělení"
              Exit Sub
        End If
  End If
End Sub

'###############################################
'Skrytí prázdných řádků v Admin Listu - přístupy
Sub HideEmptyPristupyList()
  Dim xRG As Range
  Application.ScreenUpdating = False
  ActiveSheet.Unprotect
        For Each xRG In A_Pristupy.Range("D7:D306")
              If xRG.Value = "x" Then
              xRG.EntireRow.Hidden = True
              Else
              xRG.EntireRow.Hidden = False
              End If
        Next xRG
  ActiveSheet.Protect
  Application.ScreenUpdating = True
End Sub

'##################################################
'Zobrazení prázdných řádků v Admin Listu - přístupy
Sub UnhideEmptyPristupyList()
  Dim xRG As Range
  Application.ScreenUpdating = False
  ActiveSheet.Unprotect
        For Each xRG In A_Pristupy.Range("D7:D306")
              If xRG.Value = "x" Then
              xRG.EntireRow.Hidden = False
              Else
              xRG.EntireRow.Hidden = False
              End If
        Next xRG
  ActiveSheet.Protect
  Application.ScreenUpdating = True
End Sub

'###############################################################################
'Makro pro tlačítko Zápis 1 - dostane nás do zápisového listu prvního pracovníka
Sub Zapis1()
  Dim list_na_otevreni As String
  list_na_otevreni = "Z1"
  Dim listy As Range
  Dim sloupec_opravneni As Integer
  Dim radek_opravneni As Integer
  Dim opravneni As String
  Set listy = A_Pristupy.Range("A6:AR6")
  
  On Error GoTo Errorhandler
        sloupec_opravneni = Application.WorksheetFunction.Match(list_na_otevreni, listy, 0)
        radek_opravneni = Application.WorksheetFunction.Match(Environ("Username"), A_Pristupy.Range("E6:E306"), 0) + 5
        opravneni = A_Pristupy.Cells(radek_opravneni, sloupec_opravneni).Value
        Select Case opravneni
              Case "x"
                    MsgBox "Nemáte oprávnění."
                    Exit Sub
              Case "z"
                    ActiveWorkbook.Sheets(list_na_otevreni).Activate
                    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
                                 , AllowFiltering:=True
        End Select
        Exit Sub
  Errorhandler:
  Select Case Err.Number
        Case 1004:  MsgBox "Uživatelské jméno není v seznamu oprávnění"
  End Select
End Sub

'################################################################
'Makro, které vkládá obdélník s textem do časové osy jako událost
'šířka makra se přizpůsobí podle zadaných dnů trvání události
Sub VytvorObdelnik()
  Dim ws As Worksheet
  Dim shp As Shape
  Dim novynazev As String
  Dim sirka As Double
  Dim poznamka As String
  
  On Error GoTo Errorhandler
  sirka = InputBox("Zadejte trvání události ve dnech: ", "TRVÁNÍ", 3)
  poznamka = InputBox("Popište vkládanou událost: ", "POPISEK", "Popište, co vkládáte")
  
  Set ws = ActiveSheet
  novynazev = Format(Now, "yyyymmddhhmmss") & "udalost"
  Set shp = ws.Shapes.AddShape(msoShapeRectangle, 170, 93, sirka * 18.5, 35)
        With shp
          .Name = novynazev
          .Line.Visible = msoFalse ' Bez obrysu
          .Fill.ForeColor.RGB = RGB(192, 192, 192) ' Šedá výplň
          .TextFrame.Characters.Text = poznamka
          .TextFrame.Characters.Font.Color = RGB(0, 0, 0) ' Černý text
          .TextFrame.Characters.Font.Name = "Arial Nova" ' Font Arial Nova
          .TextFrame.Characters.Font.Size = 10 ' Velikost písma 10
      End With
  
  Errorhandler:
        Select Case Err.Number
              Case 13:    MsgBox "musíte zadat trvání a poznámky pro vytvoření události"
                          Exit Sub
              Case Else: Exit Sub
        End Select
End Sub

'#############################################
'Oprávnění psát připomínky do zápisového listu
Sub opravneni_pripominky()
  Dim uzivatel As String
  uzivatel = Environ("Username")
  Dim uzivatel_s_opravnenim(1 To 4)
  uzivatel_s_opravnenim(1) = "vedouci_1"
  uzivatel_s_opravnenim(2) = "vedouci_2"
  uzivatel_s_opravnenim(3) = "vedouci_3"
  uzivatel_s_opravnenim(4) = "vedouci_4"
  If Not (uzivatel = uzivatel_s_opravnenim(1) Or uzivatel = uzivatel_s_opravnenim(2) _
                    Or uzivatel = uzivatel_s_opravnenim(3) Or uzivatel = uzivatel_s_opravnenim(4)) Then
        Application.EnableEvents = False
        Application.Undo
        Application.EnableEvents = True
        MsgBox "Zde zapisuje pouze vedoucí oddělení"
  
        Exit Sub
  End If
End Sub
