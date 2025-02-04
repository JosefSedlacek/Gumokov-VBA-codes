Option Explicit

Sub ExportToPDF()

Dim ws As Worksheet
Dim wwss As Worksheet
Dim pdfFilePath As String
Dim pdfFileName As String
Dim savePath As String
Dim pageNumber As Integer

' Nastavte list, který chcete exportovat
Set ws = ThisWorkbook.Sheets("Price List")
Set wwss = ThisWorkbook.Sheets("Categories Invoices")

' Získejte cestu kde ukládat soubory
savePath = "P:\All Access\Pro účtárnu\Silao - fakturace\"

' Získejte informace o názvu souboru
pdfFileName = "Gumokov price list " & wwss.Range("G10").Value & " " & wwss.Range("I10").Value & ".pdf"

' Získejte počet stránek, které chcete exportovat
pageNumber = 1

' Sestavte plnou cestu k PDF souboru
pdfFilePath = savePath & pdfFileName

' Pokud soubor již existuje, odstraňte ho
If Dir(pdfFilePath) <> "" Then
    Kill pdfFilePath
End If

' Exportujte list do PDF
ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfFilePath, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False, From:=1, To:=pageNumber

' Zpráva o dokončení
MsgBox "Export byl dokončen. Soubor uložen jako: " & pdfFilePath

End Sub

Sub ExportToPDF2()

Dim ws As Worksheet
Dim pdfFilePath As String
Dim pdfFileName As String
Dim savePath As String
Dim pageNumber As Integer

' Nastavte list, který chcete exportovat
Set ws = ThisWorkbook.Sheets("Categories Invoices")

' Získejte cestu kde ukládat soubory
savePath = "P:\All Access\Pro účtárnu\Silao - fakturace\"

' Získejte informace o názvu souboru
pdfFileName = "Silao - departments " & ws.Range("G10").Value & " " & ws.Range("I10").Value & ".pdf"

' Získejte počet stránek, které chcete exportovat
pageNumber = 10
'ws.Range("Y4").Value

' Sestavte plnou cestu k PDF souboru
pdfFilePath = savePath & pdfFileName

' Pokud soubor již existuje, odstraňte ho
If Dir(pdfFilePath) <> "" Then
    Kill pdfFilePath
End If

' Exportujte list do PDF
ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfFilePath, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False, From:=1, To:=pageNumber

' Zpráva o dokončení
MsgBox "Export byl dokončen. Soubor uložen jako: " & pdfFilePath

End Sub

Sub ExportToPDF3()

Dim ws As Worksheet
Dim wwss As Worksheet
Dim pdfFilePath As String
Dim pdfFileName As String
Dim savePath As String
Dim pageNumber As Integer

' Nastavte list, který chcete exportovat
Set ws = ThisWorkbook.Sheets("Summary Invoice")
Set wwss = ThisWorkbook.Sheets("Categories Invoices")

' Získejte cestu kde ukládat soubory
savePath = "P:\All Access\Pro účtárnu\Silao - fakturace\"

' Získejte informace o názvu souboru
pdfFileName = "Silao - summary " & wwss.Range("G10").Value & " " & wwss.Range("I10").Value & ".pdf"

' Získejte počet stránek, které chcete exportovat
pageNumber = 1
'ws.Range("Y4").Value

' Sestavte plnou cestu k PDF souboru
pdfFilePath = savePath & pdfFileName

' Pokud soubor již existuje, odstraňte ho
If Dir(pdfFilePath) <> "" Then
    Kill pdfFilePath
End If

' Exportujte list do PDF
ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfFilePath, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False, From:=1, To:=pageNumber

' Zpráva o dokončení
MsgBox "Export byl dokončen. Soubor uložen jako: " & pdfFilePath

End Sub
