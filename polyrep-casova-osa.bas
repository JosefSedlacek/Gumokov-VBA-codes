Option Explicit

Dim ZakazkaShape As Shape
Dim SAP_Zakazka As Variant, SAP_tvar As String, SAP_norma As Long, SAP_lis As Variant, VZ_lis As String
Dim SAP_mnozstvi As Long, SAP_Zphl As Long
Dim SAP_start_date As Long, SAP_start_time As Long, SAP_end_date As Long, SAP_end_time As Long
Dim SAP_Obsluznost As Long, Obsluznost As String

Dim Posledni_radek As Long, Vysledek_radek As Long, Posledni_vysledek_radek As Long, Trvani_zakazky As Long
Dim Sc_radek As Long, Sc_start_sloupec As Long, Sc_end_sloupec As Long, Typ_radek As Long
Dim ZakazkaID As String, Tvar As String, Lis As String, Obsluha As String

Dim Prvni_smena_vybraneho_mesice As Long, Posledni_smena_vybraneho_mesice As Long
Dim Start_smena As Long, End_smena As Long


Sub Shedule_Refresh()
''' dodělat chybovou hlášku - pokud ve filtrované oblasti listu ZakazkyDB nejsou data !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
''' doladit správné zobrazování - pokud zakázka začíná před zobrazovanou oblastí - nastavit začátek na začátek časové osy !!!!!!

ActiveWorkbook.RefreshAll

With Reporting
    For Each ZakazkaShape In .Shapes
        On Error Resume Next
        If InStr(ZakazkaShape.Name, "Reporting") > 0 Then ZakazkaShape.Delete
        On Error GoTo 0
    Next ZakazkaShape
    Prvni_smena_vybraneho_mesice = .Range("B25").Value
    Posledni_smena_vybraneho_mesice = .Range("B26").Value
End With

With ZakazkyDB

    If ZakazkyDB.Range("AA2") = 0 Then
        MsgBox "Ve vybraném období nejsou data o zákázkách"
        MsgBox "Nyní vyskočí chybové hlášení - stiskněte END !!!!!"
    End If
            
    Posledni_radek = Cells(Rows.Count, "A").End(xlUp).Row
    If Posledni_radek < 2 Then Exit Sub
 
    Posledni_vysledek_radek = .Range("Z2000").End(xlUp).Row
    If Posledni_vysledek_radek < 2 Then Exit Sub
    
    For Vysledek_radek = 2 To Posledni_vysledek_radek
    
            'získat proměnné - sloupce dat z upravené SAP tabulky
            
            ZakazkaID = .Range("Z" & Vysledek_radek).Value
            SAP_Zakazka = .Range("AA" & Vysledek_radek).Value
            SAP_tvar = .Range("AB" & Vysledek_radek).Value
            Tvar = .Range("AC" & Vysledek_radek).Value
            SAP_norma = .Range("AD" & Vysledek_radek).Value
            SAP_lis = .Range("AE" & Vysledek_radek).Value
            VZ_lis = .Range("AF" & Vysledek_radek).Value
            Lis = .Range("AG" & Vysledek_radek).Value
            SAP_mnozstvi = .Range("AH" & Vysledek_radek).Value
            SAP_Zphl = .Range("AI" & Vysledek_radek).Value
            SAP_start_date = .Range("AJ" & Vysledek_radek).Value
            SAP_start_time = .Range("AK" & Vysledek_radek).Value
            SAP_end_date = .Range("AL" & Vysledek_radek).Value
            SAP_end_time = .Range("AM" & Vysledek_radek).Value
            Start_smena = .Range("AN" & Vysledek_radek).Value
            End_smena = .Range("AO" & Vysledek_radek).Value
            SAP_Obsluznost = .Range("AP" & Vysledek_radek).Value
            Obsluha = .Range("AQ" & Vysledek_radek).Value
            
            'Nastavit začínající sloupec:
            
            If Start_smena <= Prvni_smena_vybraneho_mesice Then                       'Nastavit jako první sloupec
                Sc_start_sloupec = 28
            Else                                                                      'Nastavit začátek - na začátek měsíce
                Sc_start_sloupec = 28 + (Start_smena - Prvni_smena_vybraneho_mesice)  'Začínající sloupec
            End If
            
            'Nastavit konečný sloupec:
            
            If End_smena >= Posledni_smena_vybraneho_mesice Then                      'Konec na nebo po posledním dni měsíce - nastavit končící sloupec
                Sc_end_sloupec = 28 + (Posledni_smena_vybraneho_mesice - Prvni_smena_vybraneho_mesice)
                
            'Konec protáhlý na konec měsíce
            
            Else
                Sc_end_sloupec = 28 + (End_smena - Prvni_smena_vybraneho_mesice)       'Nastavit končící sloupec
            End If
            
            'Determinovat řádek v shedule - časové ose
            
            On Error Resume Next
            Sc_radek = Reporting.Range("AA17:AA100").Find(Lis, , xlValues, xlWhole).Row 'Najít řádek příslušného lisu
            On Error GoTo 0
            If Sc_radek = 0 Or Lis = "" Then GoTo NextTrain
            
            With Reporting
            
            'Vytvořit obrazec pro každou zakázku
            
                .Shapes("SampleShapeZakazka").Duplicate.Name = "Reporting" & ZakazkaID
                With .Shapes("Reporting" & ZakazkaID)
                        .Left = Reporting.Cells(Sc_radek, Sc_start_sloupec).Left
                        .Top = Reporting.Cells(Sc_radek, Sc_start_sloupec).Top
                        .Width = Range(Reporting.Cells(Sc_radek, Sc_start_sloupec), Reporting.Cells(Sc_radek, Sc_end_sloupec)).Width
                        .TextFrame2.TextRange.Text = Tvar                               'Nastavit do obdélníku text - Tvar příslušné zakázky
                        
                End With
            End With
NextTrain:
    Next Vysledek_radek
End With

End Sub

Sub Vybrat_zakazku()
Application.ScreenUpdating = False
    ZakazkaID = Replace(Application.Caller, Left(Application.Caller, 9), "")
    Reporting.Range("E18").Value = ZakazkaID
Application.ScreenUpdating = True
End Sub

Sub Aktualni_tyden()
'Application.ScreenUpdating = False
    Reporting.Range("B27").Value = Reporting.Range("B31").Value
    On Error Resume Next
    Reporting.Range("B23").Value = Reporting.Range("B32").Value
    On Error GoTo 0
Shedule_Refresh
'Application.ScreenUpdating = True
End Sub


Sub posun_o_tyden_zpet()
Application.ScreenUpdating = False
    If Reporting.Range("B27").Value = 1 Then
                If Reporting.Range("B23").Value = 1 Then
                    MsgBox "Jste na prvním dostupném roce databáze"
                Exit Sub
                End If
       Reporting.Range("B23").Value = Reporting.Range("B23").Value - 1 'Snížíme rok o jedna a následně nastavíme poslední týden v roce
                If Reporting.Range("B23").Value = 4 Then
                   Reporting.Range("B27").Value = 53
                Else:
                   Reporting.Range("B27").Value = 52
                End If
    Else
        Reporting.Range("B27").Value = Reporting.Range("B27").Value - 1
    End If
Shedule_Refresh
Application.ScreenUpdating = True
End Sub

Sub posun_o_tyden_dopredu()
Application.ScreenUpdating = False
    If Reporting.Range("B27").Value = 52 Then
                If Reporting.Range("B23").Value = 7 Then
                    MsgBox "Jste na posledním dostupném roce databáze. Proím, proveďte update na listu PomocnaData v buňce J7, kde je seznam dostupných roků"
                Exit Sub
                End If
        Reporting.Range("B23").Value = Reporting.Range("B23").Value + 1
        Reporting.Range("B27").Value = 1
    Else
        Reporting.Range("B27").Value = Reporting.Range("B27").Value + 1
    End If
Shedule_Refresh
Application.ScreenUpdating = True
End Sub

Sub smena_zpet()
Application.ScreenUpdating = False
Dim smena As Variant
smena = Reporting.Range("B34").Value
Dim datumek As Date
datumek = Reporting.Range("B29").Value

Select Case smena
    Case "Ranní"
        Reporting.Range("B29").Value = datumek - 1
        Reporting.Range("B34").Value = "Noční"
    Case "Odpolední"
        Reporting.Range("B34").Value = "Ranní"
    Case "Noční"
        Reporting.Range("B34").Value = "Odpolední"
End Select
Application.ScreenUpdating = True
End Sub

Sub smena_dopredu()
Application.ScreenUpdating = False
Dim smena As Variant
smena = Reporting.Range("B34").Value
Dim datumek As Date
datumek = Reporting.Range("B29").Value

Select Case smena
    Case "Noční"
        Reporting.Range("B34").Value = "Ranní"
        Reporting.Range("B29").Value = datumek + 1
    Case "Odpolední"
        Reporting.Range("B34").Value = "Noční"
    Case "Ranní"
        Reporting.Range("B34").Value = "Odpolední"
End Select
Application.ScreenUpdating = True
End Sub

Sub aktualni_smena()
Application.ScreenUpdating = False
Reporting.Range("B34").Value = "Ranní"
Reporting.Range("B29").Value = Date
Application.ScreenUpdating = True
End Sub
