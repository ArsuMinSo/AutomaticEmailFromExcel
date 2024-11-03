' MIT License
'
' Copyright (c) 2024 Arsu MinSo
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.

' -----------------------------------------
' Název makra: GenerateEmailWithOrderAndButtons
' Popis: Automatizace generování e-mailu s objednávkou a interaktivními tlačítky.
' -----------------------------------------


Sub GenerateEmailWithOrderAndButtons()
    Dim OutlookApp As Object
    Dim OutlookMail As Object
    Dim r As Long
    
    Dim CisloObjednavky As String
    Dim PopisObjednavky As String
    Dim KontaktHlaseneho As String
    Dim Nahlasil As String
    Dim Ulice As String
    Dim CisloObjektu As String
    Dim CisloBytu As String
    Dim Kategorie As String
    Dim TerminSplneni As String

    Dim UliceSCislemDomuABytu As String
    
    Dim TelozpravyMisto As String
    Dim TextObjednavka As String
    Dim TextKontaktNaHlasene As String
    Dim TextPozdrav As String
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
    Dim EmailOdesilatele As String
    Dim TextTelefon As String
    Dim Web As String
    Dim TextWeb As String

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
    KontaktHlaseneho = Cells(r, 9).Value ' Kontakt na člověka, co to nahlásil
    Hlasil = Cells(r, 10).Value ' Jméno toho, co to nahláslil
    Zapsal = Cells(r, 11).Value ' Iniciály odesílatele
    Prijemce = Cells(r, 17).Value ' Příjemce
    TerminSplneni = Cells(r, 19).Value ' Termín splnění

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
        UliceSCislemDomuABytu = Ulice & " " & CisloObjektu & "_" & CisloBytu
        TelozpravyMisto = "<p>Objednávka je na ulici: <b>" & Ulice & " v bytovém domě s číslem: " & CisloObjektu & " a číslem bytu: " & CisloBytu & "</b>.</p>"
    Else
        UliceSCislemDomuABytu = Ulice & " " & CisloObjektu
        TelozpravyMisto = "<p>Objednávka je na ulici: <b>" & Ulice & " v domě s číslem: " & CisloObjektu & "</b>.</p>"
    End If

    If Not Web = "" Then
        TextWeb = "<p style='margin-top: 0; margin-bottom: 0'>Web: <a href =" & Web & ">" & Web & "</a></p></div>"
    End If

    If Not TelefonOdesilatele = "" Then
        TextTelefon = "<p style='margin-top: 0; margin-bottom: 0'>Tel: " & TelefonOdesilatele & "</p>"
    End If

    If Not KontaktHlaseneho = "" And Not Hlasil = "" Then
        TextKontaktNaHlasene = "<tr><td>Kontakt:</td><td><b>" & Hlasil & ", tel: " & KontaktHlaseneho & "</b></td></tr>"
    End If
    
    TextPozdrav = "Dobrý den, %0Atímto"
    TextObjednavka = "objednávky číslo: " & CisloObjednavky & _
                     "%0Astručný popis objednávky: " & PopisObjednavky

    TeloObjednavka = "<div><p style='margin: 0;'>Dobrý den,</p><br>" & _
                "<p style='margin: 0;'>zasíláme Vám objednávku č.<b>" & CisloObjednavky & "</b></p><br>" & _
                "<table style='width:100%'>" & _
                    "<tr>" & _
                        "<td>Kategorie:</td>" & _
                        "<td><b>" & Kategorie & "</b></td>" & _
                    "</tr>" & _
                    "<tr>" & _
                        "<td>Popis opravy:</td>" & _
                        "<td><b>" & PopisObjednavky & "</b></td>" & _
                    "</tr>" & _ 
                    "<tr>" & _
                        "<td>Adresa provedení opravy:</td>" & _
                        "<td><b>" & UliceSCislemDomuABytu & "</b></td>" & _
                    "</tr>" & _
                TextKontaktNaHlasene & _
                    "<tr>" & _
                        "<td>Datum a čas vystavení:</td>" & _
                        "<td><b>" & Format(Now, "dd.mm.yyyy HH:MM") & "</b></td>" & _
                    "</tr>" & _
                    "<tr>" & _
                        "<td>Termín provedení opravy</td>" & _
                        "<td><b>" & TerminSplneni & "</b></td>" & _
                    "</tr>" & _
                "</table></div>"
                 
    TeloPodpis = "<br><br>S pozdravem, <div><div style='color: blue; line-height: 1;'>" & _
                "<p style='margin-top: 15; margin-bottom: 0'><b>" & TitulOdesilatele & " " & JmenoOdesilatele & " " & PrijmeniOdesilatele & "</b></p>" & _
                "<p style='margin-top: 0; margin-bottom: 0'>" & HodnostOdesilatele & "</p>" & _
                "<p style='margin-top: 0; margin-bottom: 0'>" & OdborOdesilatele & "</p>" & _
                "<p style='margin-top: 0; margin-bottom: 20'>" & OdeleniOdesilatele & "</p>" & _
                "<p style='margin-top: 0; margin-bottom: 0'><b>" & MestoUradu & "</b></p>" & _
                "<p style='margin-top: 0; margin-bottom: 0'>" & AdresaUradu & "</p>" & _
                TextTelefon & _
                TextWeb & _
                "<p style='color: green; margin-top: 20; margin-bottom: 12'>Myslete na přírodu. Skutečně potřebujete vytisknout tento e-mail?</p>" & _
                "<p style='color: blue; margin-top: 20; margin-bottom: 12'>Obsah tohoto e-mailu včetně jeho příloh je důvěrný. Pokud nejste oprávněným adresátem tohoto emailu, <b>nejste oprávněni tuto zprávu odeslat, uložit ji, zveřejnit či naložit s ní jakýmkoliv jiným způsobem</b>. V takovém případě prosím informujte odesílatele a tento e-mail včetně jeho příloh vymažte trvale ze svého systému.</p></div>"

    ' Moje iniciály
    TeloPata = "<div><hr style='border:none; border-top:1px solid #ccc;'/>" & _
                "<p style='font-size:9px; color:#555; font-style:italic; text-align:center; margin-top: 0; margin-bottom: 0'>This message was generated automatically by the system based on a standard email template.</p>" & _
                "<p style='font-size:9px; color:#555; text-align:center; margin-top: 0; margin-bottom: 0'>Template created by: Šimon Raus | <a href='mailto:simon.raus@email.cz' style='color:#666; text-decoration:none;'>simon.raus@email.cz</a></p></div>"
    
    ' Tlačítka
    TeloTlacitka = "<p style='margin-top: 30px;'>Zároveň Vás žádáme o potvrzení této objednávky níže uvedeným tlačítkem. Jakmile budete zahajovat realizaci objednávky, tak nás opět informujte kliknutím na tlačítko „Potvrdit realizaci objednávky“</p>" & _
                "<table style='width:100%; text-align:center'>" & _
                    "<tr>" & _
                        "<td style='background-color:#4CAF50'><a href='mailto:" & EmailOdesilatele & "?subject=Objednávka " & CisloObjednavky & " " & UliceSCislemDomuABytu & "-PŘIJATO?body=" & TextPozdrav & " potvrzujeme přijetí nové " & TextObjednavka & "' style='color:white; text-decoration:none;'>Potvrdit přijetí objednávky</a></td>" & _
                        "<td style='background-color:#008CBA'><a href='mailto:" & EmailOdesilatele & "?subject=Objednávka " & CisloObjednavky & " " & UliceSCislemDomuABytu & "-ZAHÁJENO?body=" & TextPozdrav & " potvrzujeme zahájení " & TextObjednavka & "' style='color:white; text-decoration:none;'>Potvrdit realizaci objednávky</a></td>" & _
                        "<td style='background-color:#4C0000'><a href='mailto:" & EmailOdesilatele & "?subject=Objednávka " & CisloObjednavky & " " & UliceSCislemDomuABytu & "-UKONČENO?body=" & TextPozdrav & " potvrzujeme dokončení " & TextObjednavka & "' style='color:white; text-decoration:none;'>Potvrdit dokončení objednávky</a></td>" & _
                    "</tr>" & _
                "</table>"
    
    ' tělo zprávy staví kompletní email dohromady
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
        .Subject = "Objednávka " & CisloObjednavky & " " & UliceSCislemDomuABytu & " " & Kategorie
        .HTMLBody = TeloZpravy
        .Display ' Zobrazí e-mail, lze použít .Send pro okamžité odeslání
    End With
    
    ' Uvolnění paměti
    Set OutlookMail = Nothing
    Set OutlookApp = Nothing
End Sub