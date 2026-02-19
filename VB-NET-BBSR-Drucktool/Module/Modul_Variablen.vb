#Region "Imports"
Imports System.Drawing
Imports System.Drawing.Imaging
Imports BBSR_Energieausweis.DynaPDF
#End Region
'--------------------------------------------------
Module Modul_Variablen
    '--------------------------------------------------
#Region "Variablen"
    '--------------------------------------------------
    ''' <summary>
    ''' Dieses ist die Variable für die Parameterwerte
    ''' </summary>
    Public Variable_Parameter As VarParameter
    ''' <summary>
    ''' Dieses ist die Variable für die Steuerungswerte
    ''' </summary>
    Public Variable_Steuerung As VarSteuerung
    ''' <summary>
    ''' Dieses ist die Variable für die Importwerte
    ''' </summary>
    Public Variable_XML_Import As VarXML
    '--------------------------------------------------
#End Region
    '--------------------------------------------------
#Region "Structure"
    '---------------------------------------------------------
    ''' <summary>
    ''' Hier findet die Strukturdeklaration für die Parameter die beim Start übergeben werden statt.
    ''' </summary>
    Structure VarParameter
        '---------------------------------------------------------
        Dim Arbeitsverzeichnis As String
        Dim Steuerungsdatei As String
        '---------------------------------------------------------
        Dim Aussteller_ID_DIBT As String
        Dim Aussteller_PWD_DIBT As String
        '---------------------------------------------------------
    End Structure
    '---------------------------------------------------------
    ''' <summary>
    ''' Hier findet die Strukturdeklaration für die Steuerungsdatei statt.
    ''' </summary>
    Structure VarSteuerung
        '---------------------------------------------------------
        Dim Methode As String                       '1= |2= |3= |4= |5= |6=
        Dim Sandbox As Boolean                      'False= /True=
        Dim Testdaten As Boolean                    'False= /True=
        Dim PDF_erzeugen As Boolean                 'False= /True=
        Dim PDF_oeffnen As Boolean                  'False= /True=
        Dim PDF_Registriernummer As Boolean         'False= /True=
        Dim Anwendung_minimieren As Boolean         'False= /True=
        Dim Anwendung_beenden As Boolean            'False= /True=
        '---------------------------------------------------------
        Dim Ausstellungsdatum As Date
        Dim Bundesland As String
        Dim Postleitzahl As String
        Dim Gesetzesgrundlage As String             '
        Dim Gebaeudeart As String                   'Wohngebäude / Nichtwohngebäude / Wohnteil gemischt genutztes Gebäude / Nichtwohnteil gemischt genutztes Gebäude
        Dim Berechnungsart As String                'Energiebedarfsausweis / Energieverbrauchsausweis
        Dim Neubau As Integer                       '1=Neubau /0=Modernisierung
        '---------------------------------------------------------
        Dim Importdatei As String
        Dim Kontrolldatei As String
        Dim Ausgabepfad As String
        Dim Ausgabedatei As String
        '---------------------------------------------------------
        Dim Ausgabequalitaet As Integer             '0=150 dpi | 1=300 dpi | 2=600 dpi
        '---------------------------------------------------------
        Dim Bild_Projekt As String
        Dim Bild_Aussteller As String
        Dim Bild_Unterschrift As String
        '---------------------------------------------------------
        Dim Registriernummer As String
        '---------------------------------------------------------
        Dim Ablauf_Fehler As Boolean
        Dim Internetverbindung_Fehler As Boolean
        '---------------------------------------------------------
        'Ergebnisse nach der Regtrierung
        '---------------------------------------------------------
        Dim Entwurf As Boolean                      'False= /True=
        '---------------------------------------------------------
        Dim Restkontingent_Anzahl As Integer
        Dim Restkontingent_Fehlermeldung As String
        '---------------------------------------------------------
        Dim Datenregistratur_Fehler_Anzahl As Integer
        Dim Datenregistratur_Fehler_ID() As Integer
        Dim Datenregistratur_Fehler_Text() As String
        '---------------------------------------------------------
        Dim Datenregistratur_Pruefung_Datendatei As Integer
        Dim Datenregistratur_Pruefung_Datendatei_Bemerkungen As String
        '---------------------------------------------------------
        Dim KontrolldateiPruefen_Erfolgreich As Boolean
        Dim KontrolldateiPruefen_Fehler_Anzahl As Integer
        Dim KontrolldateiPruefen_Fehler_ID() As String
        Dim KontrolldateiPruefen_Fehler_Kurztext() As String
        Dim KontrolldateiPruefen_Fehler_Langtext() As String
        '---------------------------------------------------------
        Dim ZusatzdatenErfassung_Fehler_Anzahl As Integer
        Dim ZusatzdatenErfassung_Fehler_ID() As String
        Dim ZusatzdatenErfassung_Fehler_Text() As String
        '---------------------------------------------------------
        Dim OffeneKontrolldateien_Anzahl As Integer
        Dim OffeneKontrolldateien_Ausweis_Registriernummer() As String
        Dim OffeneKontrolldateien_Ausweis_NummerErzeugtAm() As Date
        Dim OffeneKontrolldateien_Ausweis_Aussteller() As String
        Dim OffeneKontrolldateien_Fehlermeldung As String
        '---------------------------------------------------------
        Dim Ergebnisprotokoll_Text As String
        Dim Ergebnisprotokoll_Base64 As String
        '---------------------------------------------------------
        Dim Datenregistratur As Boolean
        Dim KontrolldateiPruefen As Boolean
        Dim Restkontingent As Boolean
        Dim ZusatzdatenErfassung As Boolean
        Dim OffeneKontrolldateien As Boolean
        '---------------------------------------------------------
        Dim XML_Kontrolldatei_Prefix As String
        '---------------------------------------------------------
    End Structure
    '---------------------------------------------------------
    ''' <summary>
    ''' Hier findet die Strukturdeklaration für den XML Import statt.   
    ''' </summary>
    Structure VarXML
        '---------------------------------------------------------
        'Arbeitsvariablen
        '---------------------------------------------------------
        Dim Berechnungsverfahren As String  'Bedarfsberechnung_4108-4701 / Bedarfsberechnung_18599 / Bedarfsberechnung_easy / Verbrauchsberechnung
        '---------------------------------------------------------
        Dim Zusaetzliche_Verbrauchsberechnung As Boolean
        '---------------------------------------------------------
        Dim Verbrauchserfassung() As VarVerbrauchsdaten
        Dim Verbrauchserfassung_Anzahl As Integer
        Dim Gebaeudenutzung_Anzahl As Integer
        Dim Energiebedarf_Anzahl As Integer
        Dim Zonen_Anzahl As Integer
        Dim Erneuerbare_Energien_65EE_Regel_Anzahl As Integer
        Dim Erneuerbare_Energien_65EE_keine_Regel_Anzahl As Integer
        '---------------------------------------------------------
        Dim Verzeichnis_Log_Datei As String
        '---------------------------------------------------------
        Dim Modernisierungsempfehlung_Anzahl As Integer
        '---------------------------------------------------------
        Dim Bilddateien() As String
        Dim Bildanzahl As Integer
        '---------------------------------------------------------
        'Aus XML geladen
        '---------------------------------------------------------
        Dim Gebaeudebezogene_Daten As VarGebaeudebezogeneDaten
        Dim Energieausweis_Daten As VarEnergieausweisDaten
        '---------------------------------------------------------
        Dim BauteilOpak() As VarBauteilOpak
        Dim BauteilTransparent() As VarBauteilTransparent
        Dim BauteilDach() As VarBauteilDach
        '---------------------------------------------------------
        Dim Heizungsanlage() As VarHeizungsanlage
        Dim Trinkwarmwasseranlage() As VarTrinkwarmwasseranlage
        Dim Kaelteanlage() As VarKaelteanlage
        Dim RLT_System() As VarRLTSystem
        '---------------------------------------------------------
        Dim Energietraeger() As VarEnergietraeger
        '---------------------------------------------------------
        Dim EE_65EE_Regel() As VarEE_65EE_Regel                         '0-200
        Dim EE_65EE_keine_Regel() As VarEE_65EE_keine_Regel             '0-200
        '---------------------------------------------------------
        Dim Energietraeger_Daten() As VarEnergietraeger_Daten           '1-8
        Dim Zeitraum_Strom_Daten() As VarZeitraum_Strom_Daten           '1-40
        Dim Nutzung_Gebaeudekategorie() As VarNutzung_Gebaeudekategorie '1-5
        '---------------------------------------------------------
        Dim Leerstandszuschlag_Heizung As VarLeerstandszuschlag_Heizung_Daten
        Dim Leerstandszuschlag_Warmwasser As VarLeerstandszuschlag_Warmwasser_Daten
        Dim Leerstandszuschlag_thermisch_erzeugte_Kaelte As VarLeerstandszuschlag_thermisch_erzeugte_Kaelte
        Dim Leerstandszuschlag_Strom As VarLeerstandszuschlag_Strom
        '---------------------------------------------------------
        Dim Warmwasserzuschlag As VarWarmwasserzuschlag_Daten
        Dim Kuehlzuschlag As VarKuehlzuschlag_Daten
        '---------------------------------------------------------
        Dim Modernisierungsempfehlungen() As VarModernisierungsempfehlungen
        '---------------------------------------------------------
    End Structure
    '---------------------------------------------------------
    ''' <summary>
    ''' Hier findet die Strukturdeklaration für den Bereich Verbrauchsdaten statt.   
    ''' </summary>
    Structure VarVerbrauchsdaten
        '---------------------------------------------------------
        Dim Startdatum As VarDate
        Dim Enddatum As VarDate
        '---------------------------------------------------------
        Dim Energietraeger_Verbrauch As VarString
        Dim Primaerenergiefaktor As VarDecimal
        '---------------------------------------------------------
        Dim Energieverbrauch As VarInteger
        '---------------------------------------------------------
        Dim Energieverbrauchsanteil_Warmwasser_zentral As VarInteger
        Dim Energieverbrauchsanteil_Heizung As VarInteger
        Dim Energieverbrauchsanteil_thermisch_erzeugte_Kaelte As VarInteger
        '---------------------------------------------------------
        Dim Energieverbrauch_Strom As VarInteger
        Dim Energieverbrauchsanteil_elektrisch_erzeugte_Kaelte As VarInteger
        '---------------------------------------------------------
        Dim Klimafaktor As VarDecimal
        '---------------------------------------------------------
    End Structure
    '---------------------------------------------------------
    ''' <summary>
    ''' Hier findet die Strukturdeklaration für die Gebäudebezogenen Daten statt.   
    ''' </summary>
    Structure VarGebaeudebezogeneDaten
        '---------------------------------------------------------
        Dim Projektbezeichnung_Aussteller As VarString
        Dim Gebaeudeadresse_Strasse_Nr As VarString
        Dim Gebaeudeadresse_Postleitzahl As VarString
        Dim Gebaeudeadresse_Ort As VarString
        Dim Ausstellervorname As VarString
        Dim Ausstellername As VarString
        Dim Aussteller_Bezeichnung As VarString
        Dim Aussteller_Strasse_Nr As VarString
        Dim Aussteller_PLZ As VarString
        Dim Aussteller_Ort As VarString
        Dim Zusatzinfos_beigefuegt As VarBoolean
        Dim Angaben_erhaeltlich As VarString
        Dim Ergaenzdende_Erlaeuterungen As VarString
        Dim Ergaenzdende_Erlaeuterungen_EE As VarString
        '---------------------------------------------------------
        Dim Treibhausgasemissionen As VarDecimal
        '---------------------------------------------------------
        Dim NWG_Aushang_Daten As Var_NWG_Aushang_Daten
        '---------------------------------------------------------
    End Structure
    '---------------------------------------------------------
    ''' <summary>
    ''' Hier findet die Strukturdeklaration für den NWB Aushang statt.   
    ''' </summary>
    Structure Var_NWG_Aushang_Daten
        '---------------------------------------------------------
        Dim Nutzenergiebedarf_Heizung_Diagramm As VarDecimal
        Dim Nutzenergiebedarf_Trinkwarmwasser_Diagramm As VarDecimal
        Dim Nutzenergiebedarf_Beleuchtung_Diagramm As VarDecimal
        Dim Nutzenergiebedarf_Lueftung_Diagramm As VarDecimal
        Dim Nutzenergiebedarf_Kuehlung_Befeuchtung_Diagramm As VarDecimal
        '---------------------------------------------------------
        Dim Endenergiebedarf_Heizung_Diagramm As VarDecimal
        Dim Endenergiebedarf_Trinkwarmwasser_Diagramm As VarDecimal
        Dim Endenergiebedarf_Beleuchtung_Diagramm As VarDecimal
        Dim Endenergiebedarf_Lueftung_Diagramm As VarDecimal
        Dim Endenergiebedarf_Kuehlung_Befeuchtung_Diagramm As VarDecimal
        '---------------------------------------------------------
        Dim Primaerenergiebedarf_Heizung_Diagramm As VarDecimal
        Dim Primaerenergiebedarf_Trinkwarmwasser_Diagramm As VarDecimal
        Dim Primaerenergiebedarf_Beleuchtung_Diagramm As VarDecimal
        Dim Primaerenergiebedarf_Lueftung_Diagramm As VarDecimal
        Dim Primaerenergiebedarf_Kuehlung_Befeuchtung_Diagramm As VarDecimal
        '---------------------------------------------------------
    End Structure
    '---------------------------------------------------------
    ''' <summary>
    ''' Hier findet die Strukturdeklaration für die Energieausweis Daten statt.   
    ''' </summary>
    Structure VarEnergieausweisDaten
        '---------------------------------------------------------
        Dim Registriernummer As VarString
        Dim Ausstellungsdatum As VarString
        Dim Bundesland As VarString
        Dim Postleitzahl As VarString
        Dim Gebaeudeteil As VarString
        Dim Baujahr_Gebaeude As VarString
        Dim Altersklasse_Gebaeude As VarString
        Dim Baujahr_Waermeerzeuger As VarString
        Dim Altersklasse_Waermeerzeuger As VarString
        Dim wesentliche_Energietraeger_Heizung As VarString
        Dim wesentliche_Energietraeger_Warmwasser As VarString
        Dim Erneuerbare_Art As VarString
        Dim Erneuerbare_Verwendung As VarString
        Dim Lueftungsart_Fensterlueftung As VarBoolean
        Dim Lueftungsart_Schachtlueftung As VarBoolean
        Dim Lueftungsart_Anlage_o_WRG As VarBoolean
        Dim Lueftungsart_Anlage_m_WRG As VarBoolean
        Dim Anlage_zur_Kuehlung As VarBoolean 'nur in 2016
        Dim Kuehlungsart_passive_Kuehlung As VarBoolean
        Dim Kuehlungsart_Strom As VarBoolean
        Dim Kuehlungsart_Waerme As VarBoolean
        Dim Kuehlungsart_gelieferte_Kaelte As VarBoolean
        Dim Keine_inspektionspflichtige_Anlage As VarBoolean
        Dim Anzahl_Klimanlagen As VarInteger
        Dim Anlage_groesser_12kW_ohneGA As VarBoolean
        Dim Anlage_groesser_12kW_mitGA As VarBoolean
        Dim Anlage_groesser_70kW As VarBoolean
        Dim Faelligkeitsdatum_Inspektion As VarDate
        '---------------------------------------------------------
        Dim Nutzung_zur_Erfuellung_von_EE_neue_Anlage As VarBoolean
        Dim EE_Angabe_Warmwasser As VarBoolean
        Dim EE_Angabe_Heizung As VarBoolean
        Dim Keine_Pauschale_Erfuellungsoptionen_Anlagentyp As VarBoolean
        Dim Hausuebergabestation As VarBoolean
        Dim Waermepumpe As VarBoolean
        Dim Stromdirektheizung As VarBoolean
        Dim Solarthermische_Anlage As VarBoolean
        Dim Heizungsanlage_Biomasse_Wasserstoff_Wasserstoffderivale As VarBoolean
        Dim Waermepumpen_Hybridheizung As VarBoolean
        Dim Solarthermie_Hybridheizung As VarBoolean
        Dim Dezentral_elektrische_Warmwasserbereitung As VarBoolean
        Dim Nutzung_bei_Bestandsanlagen As VarBoolean
        '---------------------------------------------------------
        Dim Treibhausgasemissionen As VarDecimal
        Dim Treibhausgasemissionen_Zusaetzliche_Verbrauchsdaten As VarDecimal
        '---------------------------------------------------------
        Dim Ausstellungsanlass As VarString
        Dim Datenerhebung_Aussteller As VarBoolean
        Dim Datenerhebung_Eigentuemer As VarBoolean
        '---------------------------------------------------------
        Dim Wohngebaeude As VarWohngebaeude
        Dim Nichtwohngebaeude As VarNichtwohngebaeude
        '---------------------------------------------------------
        Dim Empfehlungen_moeglich As VarBoolean
        Dim Keine_Modernisierung_Erweiterung_Vorhaben As VarBoolean
        Dim Softwarehersteller_Programm_Version As VarString
        '---------------------------------------------------------
    End Structure
    '---------------------------------------------------------
    ''' <summary>
    ''' Hier findet die Strukturdeklaration für das Wohngebäude statt.   
    ''' </summary>
    Structure VarWohngebaeude
        '---------------------------------------------------------
        Dim Gebaeudetyp As VarString
        Dim Anzahl_Wohneinheiten As VarInteger
        Dim Gebaeudenutzflaeche As VarInteger
        '---------------------------------------------------------
        Dim Verbrauchswerte As VarVerbrauchswerte
        Dim Bedarfswerte_easy As VarBedarfswerte_easy
        Dim Bedarfswerte_4108_4701 As VarBedarfswerte_4108_4701
        Dim Bedarfswerte_18599 As VarBedarfswerte_18599
        '---------------------------------------------------------
    End Structure
    '---------------------------------------------------------
    ''' <summary>
    ''' Hier findet die Strukturdeklaration für das Nichtwohngebäude statt.   
    ''' </summary>
    Structure VarNichtwohngebaeude
        '---------------------------------------------------------
        Dim Hauptnutzung_Gebaeudekategorie As VarString
        Dim Hauptnutzung_Gebaeudekategorie_Sonstiges_Beschreibung As VarString
        Dim Nettogrundflaeche As VarInteger
        '---------------------------------------------------------
        Dim Verbrauchswerte_NWG As VarVerbrauchswerteNWG
        Dim Bedarfswerte_NWG As VarBedarfswerteNWG
        '---------------------------------------------------------
        Dim Zone() As VarZone
        '---------------------------------------------------------
    End Structure
    '---------------------------------------------------------
    ''' <summary>
    ''' Hier findet die Strukturdeklaration für die Verbrauchsdaten WB statt.   
    ''' </summary>
    Structure VarVerbrauchswerte
        '---------------------------------------------------------
        Dim Flaechenermittlung_AN_aus_Wohnflaeche As VarBoolean
        Dim Wohnflaeche As VarInteger
        Dim Keller_beheizt As VarBoolean
        '---------------------------------------------------------
        Dim Mittlerer_Endenergieverbrauch As VarDecimal
        Dim Mittlerer_Primaerenergieverbrauch As VarDecimal
        Dim Energieeffizienzklasse As VarString
        '---------------------------------------------------------
    End Structure
    '---------------------------------------------------------
    ''' <summary>
    ''' Hier findet die Strukturdeklaration für die Verbrauchsdaten NWB statt.   
    ''' </summary>
    Structure VarVerbrauchswerteNWG
        '---------------------------------------------------------
        Dim Warmwasser_enthalten As VarBoolean
        Dim Kuehlung_enthalten As VarBoolean
        '---------------------------------------------------------
        Dim Strom_Daten As VarStrom
        '---------------------------------------------------------
        Dim Endenergieverbrauch_Waerme As VarDecimal
        Dim Endenergieverbrauch_Strom As VarDecimal
        Dim Endenergieverbrauch_Waerme_Vergleichswert As VarDecimal
        Dim Endenergieverbrauch_Strom_Vergleichswert As VarDecimal
        Dim Primaerenergieverbrauch As VarDecimal
        '---------------------------------------------------------
    End Structure
    '---------------------------------------------------------
    ''' <summary>
    ''' Hier findet die Strukturdeklaration für die Bedarfswerte easy statt.   
    ''' </summary>
    Structure VarBedarfswerte_easy
        '---------------------------------------------------------
        Dim Wohngebaeude_Anbaugrad As VarString
        '---------------------------------------------------------
        Dim Anzahl_Geschosse As VarInteger
        Dim Geschoss_Bruttogeschossflaechenumfang As VarInteger
        Dim Geschoss_Bruttogeschossflaeche As VarInteger
        Dim Dach_Bruttogeschossflaechenumfang As VarInteger
        Dim Dach_Bruttogeschossflaeche As VarInteger
        Dim Aufsummierte_Bruttogeschossflaeche As VarInteger
        Dim Mittlere_Geschosshoehe As VarDecimal
        Dim Kompaktheit As VarBoolean
        Dim Deckungsgleichheit As VarBoolean
        Dim Fensterflaechenanteil_Nordost_Nord_Nordwest As VarInteger
        Dim Fensterflaechenanteil_Gesamt As VarInteger
        Dim Dach_transparente_Bauteile_Fensterflaechenanteil As VarDecimal
        Dim Spezielle_Fenstertueren_Flaechenanteil As VarDecimal
        Dim Außentueren_Flaechenanteil As VarDecimal
        Dim Keine_Anlage_zur_Kuehlung As VarBoolean
        Dim Anforderung_Waermebruecken_erfuellt As VarBoolean
        Dim Gebaeudedichtheit As VarBoolean
        Dim Heiz_Warmwassersystem As VarString
        Dim Lueftungsanlagenanforderungen As VarBoolean
        Dim Waermeschutz_Variante As VarString
        Dim Endenergiebedarf As VarDecimal
        Dim Energieeffizienzklasse As VarString
        Dim Primaerenergiebedarf_Ist_Wert As VarDecimal
        Dim Primaerenergiebedarf_Anforderungswert As VarDecimal
        Dim Energetische_Qualitaet_Ist_Wert As VarDecimal
        Dim Energetische_Qualitaet_Anforderungs_Wert As VarDecimal
        '---------------------------------------------------------
        'EnEV 2016 / GEG 2020/2023
        Dim Art_der_Nutzung_erneuerbaren_Energie_1 As VarString
        Dim Deckungsanteil_1 As VarInteger
        Dim Anteil_der_Pflichterfuellung_1 As VarInteger
        Dim Art_der_Nutzung_erneuerbaren_Energie_2 As VarString
        Dim Deckungsanteil_2 As VarInteger
        Dim Anteil_der_Pflichterfuellung_2 As VarInteger
        '---------------------------------------------------------
        'EnEV 2016
        Dim Art_der_Nutzung_erneuerbaren_Energie_3 As VarString
        Dim Deckungsanteil_3 As VarInteger
        '---------------------------------------------------------
    End Structure
    '---------------------------------------------------------
    ''' <summary>
    ''' Hier findet die Strukturdeklaration für die Bedarfswerte 4108/4701 statt.   
    ''' </summary>
    Structure VarBedarfswerte_4108_4701
        '---------------------------------------------------------
        Dim Wohngebaeude_Anbaugrad As VarString
        Dim Bruttovolumen As VarInteger
        Dim durchschnittliche_Geschosshoehe As VarDecimal
        '---------------------------------------------------------
        Dim Waermebrueckenzuschlag As VarDecimal
        Dim Transmissionswaermeverlust As VarInteger
        Dim Luftdichtheit As VarString
        Dim Lueftungswaermeverlust As VarInteger
        Dim Solare_Waermegewinne As VarInteger
        Dim Interne_Waermegewinne As VarInteger
        '---------------------------------------------------------
        Dim Pufferspeicher_Nenninhalt As VarInteger
        Dim Heizkreisauslegungstemperatur As VarString
        Dim Heizungsanlage_innerhalb_Huelle As VarBoolean
        '---------------------------------------------------------
        Dim Trinkwarmwasserspeicher_Nenninhalt As VarInteger
        Dim Trinkwarmwasserverteilung_Zirkulation As VarBoolean
        Dim Vereinfachte_Datenaufnahme As VarBoolean
        Dim spezifischer_Transmissionswaermeverlust_Ist As VarDecimal
        Dim spezifischer_Transmissionswaermeverlust_Hoechstwert As VarDecimal
        Dim angerechneter_lokaler_erneuerbarer_Strom As VarInteger
        Dim Innovationsklausel As VarBoolean
        Dim Quartiersregelung As VarBoolean
        '---------------------------------------------------------
        Dim Primaerenergiebedarf_Hoechstwert_Bestand As VarDecimal
        Dim Endenergiebedarf_Hoechstwert_Bestand As VarDecimal
        Dim Treibhausgasemissionen_Hoechstwert_Bestand As VarDecimal
        '---------------------------------------------------------
        Dim Primaerenergiebedarf_Hoechstwert_Neubau As VarDecimal
        Dim Endenergiebedarf_Hoechstwert_Neubau As VarDecimal
        Dim Treibhausgasemissionen_Hoechstwert_Neubau As VarDecimal
        '---------------------------------------------------------
        Dim Endenergiebedarf_Waerme_AN As VarDecimal
        Dim Endenergiebedarf_Hilfsenergie_AN As VarDecimal
        Dim Endenergiebedarf_Gesamt As VarDecimal
        Dim Primaerenergiebedarf As VarDecimal
        Dim Energieeffizienzklasse As VarString
        '---------------------------------------------------------
        'EnEV 2016 / GEG 2020/2023
        Dim Art_der_Nutzung_erneuerbaren_Energie_1 As VarString
        Dim Deckungsanteil_1 As VarInteger
        Dim Anteil_der_Pflichterfuellung_1 As VarInteger
        Dim Art_der_Nutzung_erneuerbaren_Energie_2 As VarString
        Dim Deckungsanteil_2 As VarInteger
        Dim Anteil_der_Pflichterfuellung_2 As VarInteger
        Dim spezifischer_Transmissionswaermetransferkoeffizient_verschaerft As VarDecimal
        '---------------------------------------------------------
        'EnEV 2016
        Dim Art_der_Nutzung_erneuerbaren_Energie_3 As VarString
        Dim Deckungsanteil_3 As VarInteger
        Dim verschaerft_nach_EEWaermeG_7_1_2_eingehalten As VarBoolean
        Dim verschaerft_nach_EEWaermeG_8 As VarInteger
        Dim Primaerenergiebedarf_Hoechstwert_verschaerft As VarDecimal
        '---------------------------------------------------------
        'GEG 2020/2023
        Dim verschaerft_nach_GEG_45_eingehalten As VarBoolean
        Dim nicht_verschaerft_nach_GEG_34 As VarBoolean
        Dim verschaerft_nach_GEG_34 As VarInteger
        Dim Anforderung_nach_GEG_16_unterschritten As VarInteger
        '---------------------------------------------------------
        'GEG 2024
        'in dem GEG 2024 gibt es keine EEWärme mehr für 4108-6
        '---------------------------------------------------------
        Dim Sommerlicher_Waermeschutz As VarBoolean
        '---------------------------------------------------------
    End Structure
    '---------------------------------------------------------
    ''' <summary>
    ''' Hier findet die Strukturdeklaration für die Bedarfswerte 18599 statt.   
    ''' </summary>
    Structure VarBedarfswerte_18599
        '---------------------------------------------------------
        Dim Wohngebaeude_Anbaugrad As VarString
        Dim Bruttovolumen As VarInteger
        Dim durchschnittliche_Geschosshoehe As VarDecimal
        '---------------------------------------------------------
        Dim Waermebrueckenzuschlag As VarDecimal
        Dim Transmissionswaermesenken As VarInteger
        Dim Luftdichtheit As VarString
        Dim Lueftungswaermesenken As VarInteger
        Dim Waermequellen_durch_solare_Einstrahlung As VarInteger
        Dim Interne_Waermequellen As VarInteger
        '---------------------------------------------------------
        Dim Pufferspeicher_Nenninhalt As VarInteger
        Dim Auslegungstemperatur As VarString
        Dim Heizsystem_innerhalb_Huelle As VarBoolean
        '---------------------------------------------------------
        Dim Trinkwarmwasserspeicher_Nenninhalt As VarInteger
        Dim Trinkwarmwasserverteilung_Zirkulation As VarBoolean
        Dim Vereinfachte_Datenaufnahme As VarBoolean
        Dim spezifischer_Transmissionswaermetransferkoeffizient_Ist As VarDecimal
        Dim spezifischer_Transmissionswaermetransferkoeffizient_Hoechstwert As VarDecimal
        Dim angerechneter_lokaler_erneuerbarer_Strom As VarInteger
        Dim Innovationsklausel As VarBoolean
        Dim Quartiersregelung As VarBoolean
        Dim Primaerenergiebedarf_Hoechstwert_Bestand As VarDecimal
        Dim Endenergiebedarf_Hoechstwert_Bestand As VarDecimal
        Dim Treibhausgasemissionen_Hoechstwert_Bestand As VarDecimal
        Dim Primaerenergiebedarf_Hoechstwert_Neubau As VarDecimal
        Dim Endenergiebedarf_Hoechstwert_Neubau As VarDecimal
        Dim Treibhausgasemissionen_Hoechstwert_Neubau As VarDecimal
        '---------------------------------------------------------
        Dim Endenergiebedarf_Waerme_AN As VarDecimal
        Dim Endenergiebedarf_Hilfsenergie_AN As VarDecimal
        Dim Endenergiebedarf_Gesamt As VarDecimal
        Dim Primaerenergiebedarf_AN As VarDecimal
        Dim Energieeffizienzklasse As VarString
        '---------------------------------------------------------
        Dim Anteil_an_Waermeenergiebedarf_Berechnung As VarBoolean
        Dim Weitere_Eintraege_und_Erlaeuterungen_in_der_Anlage As VarBoolean
        '---------------------------------------------------------
        'EnEV 2016 / GEG 2020/2023
        Dim Art_der_Nutzung_erneuerbaren_Energie_1 As VarString
        Dim Deckungsanteil_1 As VarInteger
        Dim Anteil_der_Pflichterfuellung_1 As VarInteger
        Dim Art_der_Nutzung_erneuerbaren_Energie_2 As VarString
        Dim Deckungsanteil_2 As VarInteger
        Dim Anteil_der_Pflichterfuellung_2 As VarInteger
        Dim spezifischer_Transmissionswaermetransferkoeffizient_verschaerft As VarDecimal
        '---------------------------------------------------------
        'EnEV 2016
        Dim Art_der_Nutzung_erneuerbaren_Energie_3 As VarString
        Dim Deckungsanteil_3 As VarInteger
        Dim verschaerft_nach_EEWaermeG_7_1_2_eingehalten As VarBoolean
        Dim verschaerft_nach_EEWaermeG_8 As VarInteger
        Dim Primaerenergiebedarf_Hoechstwert_verschaerft As VarDecimal
        '---------------------------------------------------------
        'GEG 2020/2023
        Dim verschaerft_nach_GEG_45_eingehalten As VarBoolean
        Dim nicht_verschaerft_nach_GEG_34 As VarBoolean
        Dim verschaerft_nach_GEG_34 As VarInteger
        Dim Anforderung_nach_GEG_16_unterschritten As VarInteger
        '---------------------------------------------------------
        'GEG 2024
        'Angaben-erneuerbare-Energien-65EE-Regel
        '---------------------------------------------------------
        Dim Sommerlicher_Waermeschutz As VarBoolean
        '---------------------------------------------------------
    End Structure
    '---------------------------------------------------------
    ''' <summary>
    ''' Hier findet die Strukturdeklaration für die Bedarfswerte NWB statt.   
    ''' </summary>
    Structure VarBedarfswerteNWG
        '---------------------------------------------------------
        Dim Bruttovolumen As VarInteger
        '---------------------------------------------------------
        Dim Waermebrueckenzuschlag As VarDecimal
        Dim mittlere_Waermedurchgangskoeffizienten As VarBoolean
        Dim Transmissionswaermesenken As VarInteger
        Dim Luftdichtheit As VarString
        '---------------------------------------------------------
        Dim Pufferspeicher_Nenninhalt As VarInteger
        Dim Auslegungstemperatur As VarString
        Dim Heizsystem_innerhalb_Huelle As VarBoolean
        '---------------------------------------------------------
        Dim Trinkwarmwasserspeicher_Nenninhalt As VarInteger
        Dim Trinkwarmwasserverteilung_Zirkulation As VarBoolean
        '---------------------------------------------------------
        Dim Deckungsanteil_RLT_Kuehlung As VarInteger
        Dim Deckungsanteil_Direkte_Raumkuehlung As VarInteger
        '---------------------------------------------------------
        Dim Automatisierungsgrad As VarString
        Dim Automatisierungsgrad_Technisches_Gebaeudemanagement As VarString
        Dim angerechneter_lokaler_erneuerbarer_Strom As VarInteger
        Dim Innovationsklausel As VarBoolean
        Dim Quartiersregelung As VarBoolean
        '---------------------------------------------------------
        Dim Primaerenergiebedarf_Hoechstwert_Bestand As VarDecimal
        Dim Endenergiebedarf_Hoechstwert_Bestand As VarDecimal
        Dim Treibhausgasemissionen_Hoechstwert_Bestand As VarDecimal
        '---------------------------------------------------------
        Dim Primaerenergiebedarf_Hoechstwert_Neubau As VarDecimal
        Dim Endenergiebedarf_Hoechstwert_Neubau As VarDecimal
        Dim Treibhausgasemissionen_Hoechstwert_Neubau As VarDecimal
        '---------------------------------------------------------
        Dim Endenergiebedarf_Waerme_NGF As VarDecimal
        Dim Endenergiebedarf_Strom_NGF As VarDecimal
        Dim Endenergiebedarf_Gesamt_NGF As VarDecimal
        Dim Primaerenergiebedarf_NGF As VarDecimal
        Dim Ein_Zonen_Modell As VarBoolean
        Dim Vereinfachte_Datenaufnahme As VarBoolean
        Dim Vereinfachungen_18599_1_D As VarBoolean
        '---------------------------------------------------------
        'EnEV 2016 / GEG 2020/2023
        Dim Art_der_Nutzung_erneuerbaren_Energie_1 As VarString
        Dim Deckungsanteil_1 As VarInteger
        Dim Anteil_der_Pflichterfuellung_1 As VarInteger
        Dim Art_der_Nutzung_erneuerbaren_Energie_2 As VarString
        Dim Deckungsanteil_2 As VarInteger
        Dim Anteil_der_Pflichterfuellung_2 As VarInteger
        Dim spezifischer_Transmissionswaermetransferkoeffizient_verschaerft As VarDecimal
        '---------------------------------------------------------
        'EnEV 2016
        Dim Art_der_Nutzung_erneuerbaren_Energie_3 As VarString
        Dim Deckungsanteil_3 As VarInteger
        Dim verschaerft_nach_EEWaermeG_7_1_2_eingehalten As VarBoolean
        Dim verschaerft_nach_EEWaermeG_8 As VarInteger
        Dim Primaerenergiebedarf_Hoechstwert_verschaerft As VarDecimal
        '---------------------------------------------------------
        'GEG 2020/2023
        Dim verschaerft_nach_GEG_45_eingehalten As VarBoolean
        Dim nicht_verschaerft_nach_GEG_34 As VarBoolean
        Dim verschaerft_nach_GEG_34 As VarInteger
        Dim Anforderung_nach_GEG_16_unterschritten As VarInteger
        Dim Anforderung_nach_GEG_19_unterschritten As VarInteger
        Dim Anforderung_nach_GEG_52_Renovierung_eingehalten As VarBoolean
        '---------------------------------------------------------
        'GEG 2024
        Dim Anteil_an_Waermeenergiebedarf_Berechnung As VarBoolean
        Dim Weitere_Eintraege_und_Erlaeuterungen_in_der_Anlage As VarBoolean
        '---------------------------------------------------------
        Dim Sommerlicher_Waermeschutz As VarBoolean
        '---------------------------------------------------------
    End Structure
    '---------------------------------------------------------
    ''' <summary>
    ''' Hier findet die Strukturdeklaration für das Bauteil Opak statt.   
    ''' </summary>
    Structure VarBauteilOpak
        Dim Flaechenbezeichnung As VarString
        Dim Flaeche As VarInteger
        Dim U_Wert As VarDecimal
        Dim Ausrichtung As VarString
        Dim grenztAn As VarString
        Dim Glasdach_Lichtband_Lichtkuppel As VarBoolean
        Dim Vorhangfassade As VarBoolean
    End Structure
    '---------------------------------------------------------
    ''' <summary>
    ''' Hier findet die Strukturdeklaration für das Bauteil Transparent statt.   
    ''' </summary>
    Structure VarBauteilTransparent
        Dim Flaechenbezeichnung As VarString
        Dim Flaeche As VarInteger
        Dim U_Wert As VarDecimal
        Dim g_Wert As VarDecimal
        Dim Ausrichtung As VarString
        Dim Glasdach_Lichtband_Lichtkuppel As VarBoolean
        Dim Vorhangfassade As VarBoolean
    End Structure
    '---------------------------------------------------------
    ''' <summary>
    ''' Hier findet die Strukturdeklaration für das Bauteil Dach statt.   
    ''' </summary>
    Structure VarBauteilDach
        Dim Flaechenbezeichnung As VarString
        Dim Flaeche As VarInteger
        Dim U_Wert As VarDecimal
    End Structure
    '---------------------------------------------------------
    ''' <summary>
    ''' Hier findet die Strukturdeklaration für die Zone statt.   
    ''' </summary>
    Structure VarZone
        Dim Zonenbezeichnung As VarString
        Dim Nutzung As VarString
        Dim Anwenderspezifische_Nutzung_Bezeichnung As VarString
        Dim Zonenbesonderheiten As VarString
        Dim Nettogrundflaeche_Zone As VarInteger
        Dim mittlere_lichte_Raumhoehe As VarDecimal
        Dim Sonnenschutz_System As VarString
        Dim Beleuchtungs_System As VarString
        Dim Beleuchtungs_Verteilung As VarString
        Dim Praesenzkontrolle_Kunstlicht As VarBoolean
        Dim Tageslichtabhaengige_Kontrollsysteme As VarBoolean
        Dim Endenergiebedarf_Heizung As VarInteger
        Dim Endenergiebedarf_Kuehlung As VarInteger
        Dim Endenergiebedarf_Befeuchtung As VarInteger
        Dim Endenergiebedarf_Trinkwarmwasser As VarInteger
        Dim Endenergiebedarf_Beleuchtung As VarInteger
        Dim Endenergiebedarf_Lufttransport As VarInteger
        Dim Endenergiebedarf_Hilfsenergie As VarInteger
    End Structure
    '---------------------------------------------------------
    ''' <summary>
    ''' Hier findet die Strukturdeklaration für die Heizungsanlage statt.   
    ''' </summary>
    Structure VarHeizungsanlage
        Dim Waermeerzeuger_Bauweise_4701 As VarString
        Dim Waermeerzeuger_Bauweise_18599 As VarString
        Dim Nennleistung As VarInteger
        Dim Waermeerzeuger_Baujahr As VarInteger
        Dim Anzahl_baugleiche As VarInteger
        Dim Energietraeger As VarString
        Dim Primaerenergiefaktor As VarDecimal
        Dim Emissionsfaktor As VarInteger
    End Structure
    '---------------------------------------------------------
    ''' <summary>
    ''' Hier findet die Strukturdeklaration für die Trinkwarwasseranlage statt.   
    ''' </summary>
    Structure VarTrinkwarmwasseranlage
        Dim Trinkwarmwassererzeuger_Bauweise_4701 As VarString
        Dim Trinkwarmwassererzeuger_Bauweise_18599 As VarString
        Dim Trinkwarmwassererzeuger_Baujahr As VarInteger
        Dim Anzahl_baugleiche As VarInteger
    End Structure
    '---------------------------------------------------------
    ''' <summary>
    ''' Hier findet die Strukturdeklaration für die Kälteanlage statt.   
    ''' </summary>
    Structure VarKaelteanlage
        Dim Kaelteerzeuger_Bauweise As VarString
        Dim Kaelteerzeuger_Regelung As VarString
        Dim Rueckkuehlung_Bauweise As VarString
        Dim Kaelteverteilung_Primaerkreis_Temperatur As VarString
        Dim Nennkaelteleistung As VarInteger
        Dim Kaelteerzeuger_Baujahr As VarInteger
        Dim Anzahl_baugleiche As VarInteger
        Dim Energietraeger As VarString
        Dim Primaerenergiefaktor As VarDecimal
    End Structure
    '---------------------------------------------------------
    ''' <summary>
    ''' Hier findet die Strukturdeklaration für das RLT System statt.   
    ''' </summary>
    Structure VarRLTSystem
        Dim Funktion_Zuluft As VarBoolean
        Dim Funktion_Abluft As VarBoolean
        Dim WRG_Rueckwaermzahl As VarInteger
        Dim Funktion_Heizregister As VarBoolean
        Dim Funktion_Kuehlregister As VarBoolean
        Dim Funktion_Dampfbefeuchter As VarBoolean
        Dim Funktion_Wasserbefeuchter As VarBoolean
        Dim Energietraeger_Befeuchtung As VarString
        Dim Anzahl_baugleiche As VarInteger
    End Structure
    '---------------------------------------------------------
    ''' <summary>
    ''' Hier findet die Strukturdeklaration für den Energieträger statt.   
    ''' </summary>
    Structure VarEnergietraeger
        '---------------------------------------------------------
        Dim Energietraegerbezeichnung As VarString
        Dim Primaerenergiefaktor As VarDecimal
        Dim Endenergiebedarf_Heizung_spezifisch As VarDecimal
        Dim Endenergiebedarf_Kuehlung_Befeuchtung_spezifisch As VarDecimal
        Dim Endenergiebedarf_Trinkwarmwasser_spezifisch As VarDecimal
        Dim Endenergiebedarf_Beleuchtung_spezifisch As VarDecimal
        Dim Endenergiebedarf_Lueftung_spezifisch As VarDecimal
        Dim Endenergiebedarf_Energietraeger_Gesamtgebaeude_spezifisch As VarDecimal
        '---------------------------------------------------------
    End Structure
    '---------------------------------------------------------
    ''' <summary>
    ''' Hier findet die Strukturdeklaration für die EE 65 mit Regeln statt.   
    ''' </summary>
    Structure VarEE_65EE_Regel
        Dim Waermeerzeuger_Bauweise_18599 As VarString
        Dim Art_der_Nutzung_erneuerbaren_Energie As VarString
        Dim Deckungsanteil As VarInteger
        Dim Anteil_der_Pflichterfuellung_Anlage As VarInteger
        Dim Anteil_der_Pflichterfuellung_Gesamt As VarInteger
    End Structure
    '---------------------------------------------------------
    ''' <summary>
    ''' Hier findet die Strukturdeklaration für die EE 65 ohne Regeln statt.   
    ''' </summary>
    Structure VarEE_65EE_keine_Regel
        Dim Waermeerzeuger_Bauweise_18599 As VarString
        Dim Art_der_Nutzung_erneuerbaren_Energie As VarString
        Dim Anteil_EE_Anlage As VarInteger
    End Structure
    '---------------------------------------------------------
    ''' <summary>
    ''' Hier findet die Strukturdeklaration für die Energieträger Daten statt.   
    ''' </summary>
    Structure VarEnergietraeger_Daten
        '---------------------------------------------------------
        Dim Energietraeger_Verbrauch As VarString
        Dim Sonstiger_Energietraeger_Verbrauch As VarString
        Dim Unterer_Heizwert As VarDecimal
        Dim Primaerenergiefaktor As VarDecimal
        Dim Emissionsfaktor As VarInteger
        '---------------------------------------------------------
        Dim Zeitraum_Daten() As VarZeitraum_Daten '1-40
        '---------------------------------------------------------
    End Structure
    '---------------------------------------------------------
    ''' <summary>
    ''' Hier findet die Strukturdeklaration für die Zeitraum Daten statt.   
    ''' </summary>
    Structure VarZeitraum_Daten
        '---------------------------------------------------------
        Dim Startdatum As VarDate
        Dim Enddatum As VarDate
        Dim Verbrauchte_Menge As VarInteger
        Dim Energieverbrauch As VarInteger
        Dim Energieverbrauchsanteil_Warmwasser_zentral As VarInteger
        Dim Warmwasserwertermittlung As VarString
        Dim Energieverbrauchsanteil_thermisch_erzeugte_Kaelte As VarInteger
        Dim Energieverbrauchsanteil_Heizung As VarInteger
        Dim Klimafaktor As VarDecimal
        '---------------------------------------------------------
    End Structure
    '---------------------------------------------------------
    ''' <summary>
    ''' Hier findet die Strukturdeklaration für die Strom Daten statt.   
    ''' </summary>
    Structure VarZeitraum_Strom_Daten
        '---------------------------------------------------------
        Dim Startdatum As VarDate
        Dim Enddatum As VarDate
        Dim Energieverbrauch_Strom As VarInteger
        Dim Energieverbrauchsanteil_elektrisch_erzeugte_Kaelte As VarInteger
        '---------------------------------------------------------
    End Structure
    '---------------------------------------------------------
    ''' <summary>
    ''' Hier findet die Strukturdeklaration für die Gebäudekategorie statt.   
    ''' </summary>
    Structure VarNutzung_Gebaeudekategorie
        '---------------------------------------------------------
        Dim Gebaeudekategorie As VarString
        Dim Flaechenanteil_Nutzung As VarInteger
        Dim Vergleichswert_Waerme As VarDecimal
        Dim Vergleichswert_Strom As VarDecimal
        '---------------------------------------------------------
    End Structure
    '---------------------------------------------------------
    ''' <summary>
    ''' Hier findet die Strukturdeklaration für den Leerstandszuschlag Heizung statt.   
    ''' </summary>
    Structure VarLeerstandszuschlag_Heizung_Daten
        Dim kein_Leerstand As VarString
        Dim Leerstandsfaktor As VarDecimal
        Dim Startdatum As VarDate
        Dim Enddatum As VarDate
        Dim Leerstandszuschlag_kWh As VarInteger
        Dim Primaerenergiefaktor As VarDecimal
        Dim Zuschlagsfaktor As VarDecimal
        Dim witterungsbereinigter_Endenergieverbrauchsanteil_fuer_Heizung As VarDecimal
    End Structure
    '---------------------------------------------------------
    ''' <summary>
    ''' Hier findet die Strukturdeklaration für den Leerstandszuschlag Warmwasser statt.   
    ''' </summary>
    Structure VarLeerstandszuschlag_Warmwasser_Daten
        '---------------------------------------------------------
        Dim keine_Nutzung_von_WW As VarBoolean
        Dim kein_Leerstand As VarString
        Dim Leerstandsfaktor As VarDecimal
        Dim Startdatum As VarDate
        Dim Enddatum As VarDate
        Dim Leerstandszuschlag_kWh As VarInteger
        Dim Primaerenergiefaktor As VarDecimal
        '---------------------------------------------------------
    End Structure
    '---------------------------------------------------------
    ''' <summary>
    ''' Hier findet die Strukturdeklaration für den Leerstandszuschlag Kälte statt.   
    ''' </summary>
    Structure VarLeerstandszuschlag_thermisch_erzeugte_Kaelte
        '---------------------------------------------------------
        Dim kein_Leerstand As VarString
        Dim Leerstandsfaktor As VarDecimal
        Dim Startdatum As VarDate
        Dim Enddatum As VarDate
        Dim Leerstandszuschlag_kWh As VarInteger
        Dim Primaerenergiefaktor As VarDecimal
        '---------------------------------------------------------
    End Structure
    '---------------------------------------------------------
    ''' <summary>
    ''' Hier findet die Strukturdeklaration für den Leerstandszuschlag Strom statt.   
    ''' </summary>
    Structure VarLeerstandszuschlag_Strom
        '---------------------------------------------------------
        Dim kein_Leerstand As VarString
        Dim Leerstandsfaktor As VarDecimal
        Dim Startdatum As VarDate
        Dim Enddatum As VarDate
        Dim Leerstandszuschlag_kWh As VarInteger
        Dim Primaerenergiefaktor As VarDecimal
        '---------------------------------------------------------
    End Structure
    '---------------------------------------------------------
    ''' <summary>
    ''' Hier findet die Strukturdeklaration für den Warmwasserzuschlag statt.   
    ''' </summary>
    Structure VarWarmwasserzuschlag_Daten
        '---------------------------------------------------------
        Dim Startdatum As VarDate
        Dim Enddatum As VarDate
        Dim Primaerenergiefaktor As VarDecimal
        Dim Warmwasserzuschlag_kWh As VarInteger
        '---------------------------------------------------------
    End Structure
    '---------------------------------------------------------
    ''' <summary>
    ''' Hier findet die Strukturdeklaration für den Kühlzuschlag statt statt.   
    ''' </summary>
    Structure VarKuehlzuschlag_Daten
        '---------------------------------------------------------
        Dim Startdatum As VarDate
        Dim Enddatum As VarDate
        Dim Gebaeudenutzflaeche_gekuehlt As VarInteger
        Dim Primaerenergiefaktor As VarDecimal
        Dim Kuehlzuschlag_kWh As VarInteger
        '---------------------------------------------------------
    End Structure
    '---------------------------------------------------------
    ''' <summary>
    ''' Hier findet die Strukturdeklaration für den Strom statt.   
    ''' </summary>
    Structure VarStrom
        '---------------------------------------------------------
        Dim Stromverbrauch_enthaelt_Zusatzheizung As VarBoolean
        Dim Stromverbrauch_enthaelt_Warmwasser As VarBoolean
        Dim Stromverbrauch_enthaelt_Lueftung As VarBoolean
        Dim Stromverbrauch_enthaelt_Beleuchtung As VarBoolean
        Dim Stromverbrauch_enthaelt_Kuehlung As VarBoolean
        Dim Stromverbrauch_enthaelt_Sonstiges As VarBoolean
        '---------------------------------------------------------
        'Zeitraum-Strom
        '---------------------------------------------------------
    End Structure
    '---------------------------------------------------------
    ''' <summary>
    ''' Hier findet die Strukturdeklaration für die Modernisierungsempfehlung statt.   
    ''' </summary>
    Structure VarModernisierungsempfehlungen
        Dim Nummer As VarInteger
        Dim Bauteil_Anlagenteil As VarString
        Dim Massnahmenbeschreibung As VarString
        Dim Modernisierungskombination As VarString
        Dim Amortisation As VarString
        Dim spezifische_Kosten As VarString
    End Structure
    '---------------------------------------------------------
#End Region
    '--------------------------------------------------
#Region "String/Integer/Decimal/Boolean"
    Structure VarString
        Dim Wert As String
    End Structure
    '---------------------------------------------------------
    Structure VarInteger
        Dim Wert As Integer
    End Structure
    '---------------------------------------------------------
    Structure VarDecimal
        Dim Wert As Decimal
    End Structure
    '---------------------------------------------------------
    Structure VarBoolean
        Dim Wert As Boolean
    End Structure
    '---------------------------------------------------------
    Structure VarDate
        Dim Wert As Date
    End Structure
    '---------------------------------------------------------
#End Region
    '--------------------------------------------------
#Region "Enum"
    Enum Ausrichtung
        '--------------------
        EinzeiligLinks
        EinzeiligMitte
        EinzeiligRechts
        '--------------------
        MehrzeiligLinksMitte
        MehrzeiligMitteMitte
        MehrzeiligRechtsMitte
        '--------------------
        MehrzeiligLinksOben
        MehrzeiligMitteOben
        MehrzeiligRechtsOben
        '--------------------
        Links_kein_Abschneiden
        Mitte_kein_Abschneiden
        Rechts_kein_Abschneiden
        '--------------------
    End Enum

    Enum Lage
        Oben
        Unten
    End Enum

#End Region
    '--------------------------------------------------
#Region "Variablen Schriften"
    '--------------------------------------------------
    ''' <summary>
    ''' Hier sind alle Variablen für Font und Zeichnen für den PDF Ausdruck hinterlegt.   
    ''' </summary>
    '--------------------------------------------------
    Public Font_Schriftgroesse_20 As New Font("Arial", 20, FontStyle.Regular)
    Public Font_Schriftgroesse_30 As New Font("Arial", 30, FontStyle.Regular)
    Public Font_Schriftgroesse_40 As New Font("Arial", 40, FontStyle.Regular)
    Public Font_Schriftgroesse_50 As New Font("Arial", 50, FontStyle.Regular)
    Public Font_Schriftgroesse_60 As New Font("Arial", 60, FontStyle.Regular)
    Public Font_Schriftgroesse_70 As New Font("Arial", 70, FontStyle.Regular)
    Public Font_Schriftgroesse_80 As New Font("Arial", 80, FontStyle.Regular)
    Public Font_Schriftgroesse_90 As New Font("Arial", 90, FontStyle.Regular)
    Public Font_Schriftgroesse_100 As New Font("Arial", 100, FontStyle.Regular)
    '--------------------------------------------------
    Public Font_Bold_Schriftgroesse_40_Narrow As New Font("Arial Narrow", 40, FontStyle.Bold)
    Public Font_Schriftgroesse_40_Narrow As New Font("Arial Narrow", 40, FontStyle.Regular)
    '--------------------------------------------------
    Public Font_Bold_Schriftgroesse_20 As New Font("Arial", 20, FontStyle.Bold)
    Public Font_Bold_Schriftgroesse_30 As New Font("Arial", 30, FontStyle.Bold)
    Public Font_Bold_Schriftgroesse_40 As New Font("Arial", 40, FontStyle.Bold)
    Public Font_Bold_Schriftgroesse_50 As New Font("Arial", 50, FontStyle.Bold)
    Public Font_Bold_Schriftgroesse_60 As New Font("Arial", 60, FontStyle.Bold)
    Public Font_Bold_Schriftgroesse_70 As New Font("Arial", 70, FontStyle.Bold)
    Public Font_Bold_Schriftgroesse_80 As New Font("Arial", 80, FontStyle.Bold)
    Public Font_Bold_Schriftgroesse_90 As New Font("Arial", 90, FontStyle.Bold)
    Public Font_Bold_Schriftgroesse_100 As New Font("Arial", 100, FontStyle.Bold)
    '--------------------------------------------------
    Public Font_Skala As New Font("Arial", 60, FontStyle.Bold)
    '--------------------------------------------------
    Public Font_Seitennummer As New Font("Arial", 70, FontStyle.Bold)
    '--------------------------------------------------
    Public Font_Druckfarbe As New SolidBrush(Color.Black)
    Public Stift_Farbe_Zeichnen_Auswahl_PDF As New Pen(Color.Black, 0.75)
    Public Stift_Farbe_Zeichnen_Auswahl_Image As New Pen(Color.Black, 4)
    '--------------------------------------------------
    Public Font_Schriftgroesse_ENTWURF As New Font("Arial", 400, FontStyle.Bold)
    Public Font_Druckfarbe_ENTWURF As New SolidBrush(Color.FromArgb(75, Color.Red))
    '--------------------------------------------------
#End Region
    '--------------------------------------------------
#Region "Variablen Ausgabeformat Bitmap/PDF"
    '--------------------------------------------------
    ''' <summary>
    ''' Hier sind alle Variablen für das Ausgabeformat Bitmap und PDF hinterlegt.   
    ''' </summary>
    '--------------------------------------------------
    Public Bitmap_PixelFormat As PixelFormat = PixelFormat.Format16bppRgb565
    Public Bitmap_CompositingQuality As Drawing2D.CompositingQuality = Drawing2D.CompositingQuality.HighSpeed
    Public Bitmap_TextRenderingHint As Drawing.Text.TextRenderingHint = Drawing.Text.TextRenderingHint.ClearTypeGridFit
    Public Bitmap_SmoothingMode As Drawing2D.SmoothingMode = Drawing2D.SmoothingMode.HighQuality
    '--------------------------------------------------
    Public PDF_CompressionLevel As TCompressionLevel = TCompressionLevel.clMax
    Public PDF_CompressionFilter As TCompressionFilter = TCompressionFilter.cfFlate
    Public PDF_PDFVersion As TPDFVersion = TPDFVersion.pvPDF_1_4
    Public PDF_JPEGQuality As Integer = 70
    Public PDF_Resolution As Integer = 300
    Public PDF_Passwort As String = My.Settings.Passwort
    '--------------------------------------------------
#End Region
    '--------------------------------------------------
End Module
