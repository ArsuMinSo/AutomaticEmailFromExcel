Sub GenerateEmailWithOrderAndButtons()
    ' Definice proměnných pro Outlook
    Dim OutlookApp As Object
    Dim OutlookMail As Object
    Dim r As Long ' Proměnná pro aktuální řádek
    
    ' Definice proměnných pro data z listu
    Dim CisloObjednavky As String
    Dim PopisObjednavky As String
    Dim Ulice As String
    Dim CisloObjektu As String
    Dim CisloBytu As String
    Dim Kategorie As String
    
    ' Definice proměnných pro tělo zprávy
    Dim TeloZpravyMisto As String
    Dim TeloZpravy As String
    
    ' Definice proměnných pro odesílatele
    Dim JmenoOdesilatele As String
    Dim PrijmeniOdesilatele As String
    
    ' Definice proměnných pro e-maily
    Dim EmailOdesilatele As String
    Dim EmailPrijemce As String
    
    ' Definice proměnných pro zapsané informace
    Dim Zapsal As String
    Dim Prijemce As String
    
    ' Pomocné proměnné
    Dim Radek As Long
    Dim ShodaOdesilatel As Range
    Dim ShodaPrijemce As Range
    
    ' Listy, ze kterých se čtou data
    Dim InfoListOdesilatel As Worksheet
    Dim InfoListPrijemce As Worksheet
    Dim HlavniList As Worksheet

    ' Nastavení jednotlivých listů
    Set InfoListOdesilatel = Sheets("InfoOdesilatel")
    Set InfoListPrijemce = Sheets("InfoPrijemce")
    Set HlavniList = Sheets("Seznam požadavků") ' Změňte na skutečný název listu
  
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
        ' Načtení jména, příjmení a e-mailu odesílatele
        JmenoOdesilatele = InfoListOdesilatel.Cells(Radek, 3).Value
        PrijmeniOdesilatele = InfoListOdesilatel.Cells(Radek, 4).Value
        EmailOdesilatele = InfoListOdesilatel.Cells(Radek, 5).Value
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
        TeloZpravyMisto = "<p>Objednávka je na ulici: <b>" & Ulice & " v bytovém domě s číslem: " & CisloObjektu & " a číslem bytu: " & CisloBytu & "</b>.</p>"
    Else
        TeloZpravyMisto = "<p>Objednávka je na ulici: <b>" & Ulice & " v domě s číslem: " & CisloObjektu & "</b>.</p>"
    End If

    ' HTML kód pro tělo e-mailu s tlačítky
    TeloZpravy = "<html><body>" & _
                 "<p>Dobrý den,</p>" & _
                 "<p>tímto vytváříme novou objednávku číslo: <b>" & CisloObjednavky & "</b>.</p>" & _
                 "<p>stručný popis objednávky: <b>" & PopisObjednavky & "</b>.</p>" & _
                 TeloZpravyMisto & _
                 "<p>S pozdravem,<br>" & JmenoOdesilatele & " " & PrijmeniOdesilatele & "</p>" & _
                 "<br><p>Vyberte prosím jednu z následujících možností:</p>" & _
                 "<a href='mailto:" & EmailOdesilatele & "?subject=Práce na objednávce číslo " & CisloObjednavky & " BYLA ZAČATA' style='padding:10px 20px; background-color:#4CAF50; color:white; text-decoration:none;'>Potvrdit začátek práce</a>    " & _
                 "<a href='mailto:" & EmailOdesilatele & "?subject=Práce na objednávce číslo " & CisloObjednavky & " BYLA UKONČENA' style='padding:10px 20px; background-color:#008CBA; color:white; text-decoration:none;'>Potvrdit ukončení práce</a>    " & _
                 "<a href='mailto:" & EmailOdesilatele & "?subject=Nějaká třetí možnost' style='padding:10px 20px; background-color:#f44336; color:white; text-decoration:none;'>Možnost 3</a>    " & _
                 "<br><br>" & _
                 "<hr style='border:none; border-top:1px solid #ccc;'/>" & _
                 "<p style='font-size:9px; color:#666; font-style:italic; text-align:center;'>Tento e-mail byl automaticky generován systémem.</p>" & _
                 "<p style='font-size:9px; color:#666; text-align:center;'>Vytvořil Šimon Raus | <a href='mailto:simon.raus@email.cz' style='color:#666; text-decoration:none;'>simon.raus@email.cz</a></p>" & _
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
        .Subject = "Nová objednávka číslo: " & CisloObjednavky
        .HTMLBody = TeloZpravy
        .Display ' Zobrazí e-mail, lze použít .Send pro okamžité odeslání
    End With
    
    ' Uvolnění paměti
    Set OutlookMail = Nothing
    Set OutlookApp = Nothing
End Sub
