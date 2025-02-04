Sub UlozitJakoPDF()

    Dim ws As Worksheet
    Dim cilovaCesta As String
    Dim souborNazev As String
    Dim plnyNazevSouboru As String
    
    Set ws = ThisWorkbook.Sheets("PŘEHLED LIKVIDACE")

    cilovaCesta = ws.Range("U3").Value
    souborNazev = ws.Range("U4").Value
    plnyNazevSouboru = cilovaCesta & souborNazev
    
    ws.PageSetup.PrintArea = ws.Range("A1:R74").Address ' Upravit oblast tisku podle potřeby

    On Error Resume Next
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=plnyNazevSouboru, Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    On Error GoTo 0

    MsgBox "List 'PŘEHLED LIKVIDACE' byl úspěšně uložen jako PDF do složky: " & vbCrLf & cilovaCesta, vbInformation

End Sub
