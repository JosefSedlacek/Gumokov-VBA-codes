Sub Ulozit_do_PowerBI()

    Dim wsSeznam As Worksheet
    Dim wsTestovani As Worksheet
    Dim novySeznamWorkbook As Workbook
    Dim novyTestovaniWorkbook As Workbook
    Dim cesta As String
    Dim souborSeznam As String
    Dim souborTestovani As String
    
    '---------------------------------------
    '########### Odemknutí listů ###########

    Worksheets("Zaloha").Visible = xlSheetVisible
    
    Worksheets("AKTUALIZACE").Unprotect Password:="123456"
    Worksheets("SEZNAM ŠARŽÍ").Unprotect Password:="123456"
    Worksheets("TESTOVÁNÍ").Unprotect Password:="123456"
    Worksheets("PŘEHLED LIKVIDACE").Unprotect Password:="123456"
    
    Worksheets("AKTUALIZACE").Range("I9") = "Odesílání dat do PowerBI"
    
    Application.ScreenUpdating = False
    
    ' Cesta k adresáři
    cesta = "P:\All Access\TB HRA KPIs\podklady\"
    souborSeznam = cesta & "Směsi přehled šarží.xlsx"
    souborTestovani = cesta & "Směsi testování.xlsx"
    
    ' Nastavení listů
    Set wsSeznam = ThisWorkbook.Worksheets("SEZNAM ŠARŽÍ")
    Set wsTestovani = ThisWorkbook.Worksheets("TESTOVÁNÍ")
    
    ' Kontrola a smazání existujícího souboru pro SEZNAM ŠARŽÍ
    If Dir(souborSeznam) <> "" Then
        Kill souborSeznam
    End If
    
    ' Vytvoření nového souboru pro SEZNAM ŠARŽÍ
    Set novySeznamWorkbook = Application.Workbooks.Add
    wsSeznam.Range("C7:L1500").Copy
        With novySeznamWorkbook.Sheets(1)
            .Range("A1").PasteSpecial Paste:=xlPasteValues
            .Range("A1").PasteSpecial Paste:=xlPasteFormats
            .Name = "SEZNAM ŠARŽÍ"
        End With
    Application.CutCopyMode = False
    novySeznamWorkbook.SaveAs Filename:=souborSeznam, FileFormat:=xlOpenXMLWorkbook
    novySeznamWorkbook.Close SaveChanges:=False
    
    ' Kontrola a smazání existujícího souboru pro TESTOVÁNÍ
    If Dir(souborTestovani) <> "" Then
        Kill souborTestovani
    End If
    
    ' Vytvoření nového souboru pro TESTOVÁNÍ
    Set novyTestovaniWorkbook = Application.Workbooks.Add
    wsTestovani.Range("G7:S2500").Copy
        With novyTestovaniWorkbook.Sheets(1)
            .Range("A1").PasteSpecial Paste:=xlPasteValues
            .Range("A1").PasteSpecial Paste:=xlPasteFormats
            .Name = "TESTOVÁNÍ"
        End With
    Application.CutCopyMode = False
    novyTestovaniWorkbook.SaveAs Filename:=souborTestovani, FileFormat:=xlOpenXMLWorkbook
    novyTestovaniWorkbook.Close SaveChanges:=False
    
    Application.ScreenUpdating = True
    
    MsgBox "Data uložena do PowerBI, následuje aktualizace", vbInformation
    
    Call ZalohaPoznamek

End Sub
