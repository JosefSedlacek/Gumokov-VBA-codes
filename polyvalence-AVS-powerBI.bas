Option Explicit

Sub Ulozit_do_power_BI()

Application.ScreenUpdating = False

MsgBox "Nyní budou data polyvalence odeslána do PowerBI"

Dim ws As Worksheet

'#################################
'######################## POL data

    Dim novySeznam As Workbook
    Dim cilovaCesta As String
    
    cilovaCesta = "P:\All Access\TB HRA KPIs\podklady\Polyvalence\PolyvalAVS\POL_data.xlsx"
    Set ws = ThisWorkbook.Sheets("POL data")
    Set novySeznam = Workbooks.Add
    ws.Copy Before:=novySeznam.Sheets(1)
    Application.DisplayAlerts = False ' Potlačí upozornění na přepsání souboru
    novySeznam.SaveAs Filename:=cilovaCesta, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True ' Opět povolí upozornění
    novySeznam.Close SaveChanges:=False


'#################################
'################# LAST SAVE data

    Dim novySeznam2 As Workbook
    Dim cilovaCesta2 As String
    
    cilovaCesta2 = "P:\All Access\TB HRA KPIs\podklady\Polyvalence\PolyvalAVS\LAST_SAVE_data.xlsx"
    Set ws = ThisWorkbook.Sheets("LAST SAVE data")
    Set novySeznam2 = Workbooks.Add
    ws.Copy Before:=novySeznam2.Sheets(1)
    Application.DisplayAlerts = False ' Potlačí upozornění na přepsání souboru
    novySeznam2.SaveAs Filename:=cilovaCesta2, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True ' Opět povolí upozornění
    novySeznam2.Close SaveChanges:=False


'#################################
'################# Seznam podmínek

    Dim novySeznam3 As Workbook
    Dim cilovaCesta3 As String

    cilovaCesta3 = "P:\All Access\TB HRA KPIs\podklady\Polyvalence\PolyvalAVS\Seznam_podminek.xlsx"
    Set ws = ThisWorkbook.Sheets("Seznam podmínek")
    Set novySeznam3 = Workbooks.Add
    ws.Copy Before:=novySeznam3.Sheets(1)
    Application.DisplayAlerts = False ' Potlačí upozornění na přepsání souboru
    novySeznam3.SaveAs Filename:=cilovaCesta3, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True ' Opět povolí upozornění
    novySeznam3.Close SaveChanges:=False

Application.ScreenUpdating = True
ThisWorkbook.Worksheets("Hodnocení lisaře").Range("O2").Value = Now

MsgBox "Data odeslána"

End Sub
