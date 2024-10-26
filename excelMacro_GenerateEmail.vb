Sub GenerateEmailWithOrderAndButtons()
    Dim OutlookApp As Object
    Dim OutlookMail As Object
    Dim r As Long
    
    Dim CisloObjednavky As String
    Dim PopisObjednavky As String
    Dim Ulice As String
    Dim CisloObjektu As String
    Dim CisloBytu As String
    Dim Kategorie As String

    Dim UliceSCislemDomuABytu As String
    
    Dim TelozpravyMisto As String
    Dim TextObjednavka As String
    Dim TextPodpisPrijemce As String
    Dim TeloObjednavka As String
    Dim TeloTlacitka As String
    Dim TeloPodpis As String
    Dim TeloPata As String
    Dim TeloZpravy As String
    
    Dim TitulOdesilatele As String
    Dim JmenoOdesilatele As String
    Dim PrijmeniOdesilatele As String
    
    Dim HodnostOdesilatele As String
    Dim OdborOdesilatele As String
    Dim OdeleniOdesilatele As String
    Dim MestoUradu As String
    Dim AdresaUradu As String
    Dim TelefonOdesilatele As String
    Dim Web As String
    
    Dim EmailOdesilatele As String
    Dim EmailPrijemce As String
    
    Dim Zapsal As String
    Dim Prijemce As String
    
    Dim Radek As Long
    
    Dim ShodaOdesilatel As Range
    Dim ShodaPrijemce As Range
    
    Dim InfoListOdesilatel As Worksheet
    Dim InfoListPrijemce As Worksheet
    Dim HlavniList As Worksheet
    
    ' Nastavení listů
    Set InfoListOdesilatel = Sheets("InfoOdesilatel")
    Set InfoListPrijemce = Sheets("InfoPrijemce")
    Set HlavniList = Sheets("Seznam požadavků") ' Změňte název na skutečný název hlavního listu
 
    
    ' Získání aktuálního řádku, kde je vybraná buňka
    r = ActiveCell.Row
    
    ' Získání údajů z jednotlivých buněk v aktivním řádku
    CisloObjednavky = Cells(r, 3).Value ' Číslo objednávky
    Ulice = Cells(r, 4).Value ' Ulice
    CisloObjektu = Cells(r, 5).Value ' Číslo objektu
    CisloBytu = Cells(r, 6).Value ' Číslo bytu
    Kategorie = Cells(r, 7).Value ' Kategorie
    PopisObjednavky = Cells(r, 8).Value ' Popis objednávky
    Zapsal = Cells(r, 11).Value ' Iniciály odesílatele
    Prijemce = Cells(r, 17).Value ' Příjemce

    ' Hledání odesílatele na základě iniciál
    Set ShodaOdesilatel = InfoListOdesilatel.Columns(2).Find(What:=Zapsal, LookAt:=xlWhole)
    
    ' Kontrola, zda byl odesílatel nalezen
    If Not Zapsal = "" And Not ShodaOdesilatel Is Nothing Then
        Radek = ShodaOdesilatel.Row
        
        TitulOdesilatele = InfoListOdesilatel.Cells(Radek, 3).Value
        JmenoOdesilatele = InfoListOdesilatel.Cells(Radek, 4).Value
        PrijmeniOdesilatele = InfoListOdesilatel.Cells(Radek, 5).Value
        HodnostOdesilatele = InfoListOdesilatel.Cells(Radek, 6).Value
        OdborOdesilatele = InfoListOdesilatel.Cells(Radek, 7).Value
        OdeleniOdesilatele = InfoListOdesilatel.Cells(Radek, 8).Value
        MestoUradu = InfoListOdesilatel.Cells(Radek, 9).Value
        AdresaUradu = InfoListOdesilatel.Cells(Radek, 10).Value
        TelefonOdesilatele = InfoListOdesilatel.Cells(Radek, 11).Value
        EmailOdesilatele = InfoListOdesilatel.Cells(Radek, 12).Value
        Web = InfoListOdesilatel.Cells(Radek, 13).Value
    Else
        ' Pokud nebyl odesílatel nalezen, zobrazí se chybová zpráva
        MsgBox ("Nebylo nalezeno žádné příjmení začínající na " & Zapsal & ".")
        Exit Sub
    End If
    
    ' Hledání příjemce na základě jména
    Set ShodaPrijemce = InfoListPrijemce.Columns(2).Find(What:=Prijemce, LookAt:=xlWhole)
    
    ' Kontrola, zda byl příjemce nalezen
    If Not ShodaPrijemce = "" And Not ShodaPrijemce Is Nothing Then
        Radek = ShodaPrijemce.Row
        ' Načtení e-mailu příjemce
        EmailPrijemce = InfoListPrijemce.Cells(Radek, 3).Value
    Else
        ' Pokud nebyl příjemce nalezen, zobrazí se upozornění
        MsgBox ("Nebyl nalezen žádný příjemce " & Prijemce & ".")
    End If
    
    ' Sestavení těla zprávy podle toho, zda je vyplněno číslo bytu
    If Not CisloBytu = "" Then
        TelozpravyMisto = "<p>Objednávka je na ulici: <b>" & Ulice & " v bytovém domě s číslem: " & CisloObjektu & " a číslem bytu: " & CisloBytu & "</b>.</p>"
    Else
        TelozpravyMisto = "<p>Objednávka je na ulici: <b>" & Ulice & " v domě s číslem: " & CisloObjektu & "</b>.</p>"
    End If

    If Not CisloBytu = "" Then
        UliceSCislemDomuABytu = "" & Ulice & " " & CisloObjektu & "_" & CisloBytu
    Else
        UliceSCislemDomuABytu = "" & Ulice & " " & CisloObjektu
    End If

    TextObjednavka = "Dobrý den, %0Atímto vytváříme novou objednávku číslo: " & CisloObjednavky & "%0Astručný popis objednávky: " & PopisObjednavky & "."
    TextPodpisPrijemce = ""

    TeloObjednavka = "<div><p style='margin: 0;'>Dobrý den,</p>" & _
                "<p style='margin: 0;'>tímto vytváříme novou objednávku číslo: <b>" & CisloObjednavky & "</b>.</p>" & _
                "<p style='margin: 0;'>stručný popis objednávky: <b>" & PopisObjednavky & "</b>.</p></div>"
                 
    TeloPodpis = "<div><div style='color: blue; line-height: 1; '>" & _
                "<p style='margin-top: 30; margin-bottom: 0'><b>" & TitulOdesilatele & " " & JmenoOdesilatele & " " & PrijmeniOdesilatele & "</b></p>" & _
                "<p style='margin-top: 0; margin-bottom: 0'>" & HodnostOdesilatele & "</p>" & _
                "<p style='margin-top: 0; margin-bottom: 0'>" & OdborOdesilatele & "</p>" & _
                "<p style='margin-top: 0; margin-bottom: 20'>" & OdeleniOdesilatele & "</p>" & _
                "<p style='margin-top: 0; margin-bottom: 0'><b>" & MestoUradu & "</b></p>" & _
                "<p style='margin-top: 0; margin-bottom: 0'>" & AdresaUradu & "</p>" & _
                "<p style='margin-top: 0; margin-bottom: 0'>Tel: " & TelefonOdesilatele & "</p>" & _
                "<p style='margin-top: 0; margin-bottom: 0'>Web: <a href =" & Web & ">" & Web & "</a></p></div>" & _
                "<p style='color: green; margin-top: 20; margin-bottom: 12'>Myslete na přírodu. Skutečně potřebujete vytisknout tento e-mail?</p>" & _
                "<p style='color: blue; margin-top: 20; margin-bottom: 12'>Obsah tohoto e-mailu včetně jeho příloh je důvěrný. Pokud nejste oprávněným adresátem tohoto emailu, <b>nejste oprávněni tuto zprávu odeslat, uložit ji, zveřejnit či naložit s ní jakýmkoliv jiným způsobem</b>. V takovém případě prosím informujte odesílatele a tento e-mail včetně jeho příloh vymažte trvale ze svého systému.</p></div>"

    TeloPata = "<div><hr style='border:none; border-top:1px solid #ccc;'/>" & _
                "<p style='font-size:11px; color:#555; font-style:italic; text-align:center;'>Tento e-mail byl automaticky generován systémem.</p>" & _
                "<p style='font-size:11px; color:#555; text-align:center;'>Vytvořil Šimon Raus | <a href='mailto:simon.raus@email.cz' style='color:#666; text-decoration:none;'>simon.raus@email.cz</a></p></div>"
                 
    TeloTlacitka = "<p style='margin: 0;'>S pozdravem,<br>" & JmenoOdesilatele & " " & PrijmeniOdesilatele & "</p>" & _
                "<p style='margin-top: 30px;'>Zároveň Vás žádáme o potvrzení této objednávky níže uvedeným tlačítkem. Jakmile budete zahajovat realizaci objednávky, tak nás opět informujte kliknutím na tlačítko „Potvrdit realizaci objednávky“</p>" & _
                "<a href='mailto:" & EmailOdesilatele & "?subject=Objednávka " & CisloObjednavky & " " & UliceSCislemDomuABytu & "-PŘIJATO?body=" & TextObjednavka & " ' style='padding:10px 20px; background-color:#4CAF50; color:white; text-decoration:none;'>Potvrdit přijetí objednávky</a>    " & _
                "<a href='mailto:" & EmailOdesilatele & "?subject=Objednávka " & CisloObjednavky & " " & UliceSCislemDomuABytu & "-ZAHÁJENO?body=" & TextObjednavka & "' style='padding:10px 20px; background-color:#008CBA; color:white; text-decoration:none;'>Potvrdit realizaci objednávky</a>    "

    TeloZpravy = "<html><body>" & _
                TeloObjednavka & _
                TeloTlacitka & _
                TeloPodpis & _
                TeloPata & _
                "</body></html>"


    ' Inicializace aplikace Outlook
    On Error Resume Next
    Set OutlookApp = GetObject(Class:="Outlook.Application")
    If OutlookApp Is Nothing Then
        Set OutlookApp = CreateObject(Class:="Outlook.Application")
    End If
    On Error GoTo 0
    
    ' Vytvoření e-mailu
    Set OutlookMail = OutlookApp.CreateItem(0)
    
    ' Nastavení e-mailu (příjemce, předmět, tělo zprávy)
    With OutlookMail
        .To = EmailPrijemce
        .Subject = "Nová objednávka " & CisloObjednavky & " " & UliceSCislemDomuABytu
        .HTMLBody = TeloZpravy
        .Display ' Zobrazí e-mail, lze použít .Send pro okamžité odeslání
    End With
    
    ' Uvolnění paměti
    Set OutlookMail = Nothing
    Set OutlookApp = Nothing
End Sub

