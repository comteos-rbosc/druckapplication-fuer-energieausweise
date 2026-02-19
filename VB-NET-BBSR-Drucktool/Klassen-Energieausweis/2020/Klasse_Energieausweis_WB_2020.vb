#Region "Imports"
Imports System.Drawing
Imports System.Windows.Forms
#End Region
'--------------------------------------------------
Public Class Klasse_Energieausweis_WB_2020
    '--------------------------------------------------
#Region "New"
    '--------------------------------------------------
    ''' <summary>
    ''' Beschreibung
    ''' </summary>
    Sub New()


    End Sub
    '--------------------------------------------------
#End Region
    '--------------------------------------------------
#Region "Energieausweis erstellen"
    '--------------------------------------------------
    ''' <summary>
    ''' Hier wird der gesamte Ablauf für die Erstellung der Energieausweises abgearbeitet - Wohnbau - GEG 2020
    ''' </summary>
    Sub Energieausweis_erzeugen(ByVal Oberflaeche As Boolean, ByVal PictureBox As PictureBox, ByVal Listview As ListView, ByVal Imagelist As ImageList)
        '-------------------------------------------------------
        Try
            '-------------------------------------------------------
#Region "Variablen"
            '-------------------------------------------------------
            Dim Datum_Energiegesetz As String = "08.08.2020"
            '-------------------------------------------------------
            Dim Wert_Schleife As Integer
            Dim Wert_Schleife_Offset As Integer
            Dim Wert_Schleife_Maximum As Integer
            Dim Wert_Schleife_Zeilenzahl As Integer
            '-------------------------------------------------------
            Dim Bitmap_Seite As Bitmap
            Dim Grafik_Seite As Graphics
            Dim Hintergrundgrafik As Image
            Dim Skalierungsfaktor As Single = 0.02
            '-------------------------------------------------------
            Dim Seitennummer As Integer = 0
            '-------------------------------------------------------
            Dim Anzahl_Zeilen As Integer
            Dim Zeilenzahl As Integer
            Dim Zeilenhoehe As Single
            '-------------------------------------------------------
#End Region
            '-------------------------------------------------------
#Region "Seite 1"
            '-------------------------------------------------------
            'Seite 1 -----------------------------------------------
            '-------------------------------------------------------
            Seitennummer += 1
            '-------------------------------------------------------
            Bitmap_Seite = New Bitmap(4200, 5940, Bitmap_PixelFormat)
            Grafik_Seite = Graphics.FromImage(Bitmap_Seite)
            Grafik_Seite.CompositingQuality = Bitmap_CompositingQuality
            Grafik_Seite.TextRenderingHint = Bitmap_TextRenderingHint
            Grafik_Seite.SmoothingMode = Bitmap_SmoothingMode
            Grafik_Seite.Clear(Color.White)
            '-------------------------------------------------------
            Hintergrundgrafik = Image.FromFile(Variable_Parameter.Arbeitsverzeichnis & "/Image/GEG2020_WB_Seite_1.png")
            Grafik_Seite.DrawImage(Hintergrundgrafik, 0, 0, 4200, 5940)
            '-------------------------------------------------------
            With Variable_XML_Import
                '-------------------------------------------------------
                Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, Datum_Energiegesetz, 2128, 2616, 422, 516, Font_Schriftgroesse_50, Font_Druckfarbe)
                Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, DateAdd(DateInterval.Year, 10, Variable_Steuerung.Ausstellungsdatum), 599 + 20, 2079, 722, 820, Font_Schriftgroesse_50, Font_Druckfarbe)
                Image_Text_schreiben(Grafik_Seite, Ausrichtung.Mitte_kein_Abschneiden, Variable_Steuerung.Registriernummer, 3042, 3758, 722, 822, Font_Schriftgroesse_50, Font_Druckfarbe)
                '-------------------------------------------------------
                Select Case Variable_Steuerung.Berechnungsart
                    Case "Energiebedarfsausweis"
                        '-------------------------------------------------------
                        Select Case .Berechnungsverfahren
                            Case "Bedarfsberechnung-4108-4701"
                                '-------------------------------------------------------
                                Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksOben, .Energieausweis_Daten.Wohngebaeude.Gebaeudetyp.Wert & ", " & .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_4108_4701.Wohngebaeude_Anbaugrad.Wert, 1464 + 20, 3241, 1065, 1250, Font_Schriftgroesse_40, Font_Druckfarbe)
                            '-------------------------------------------------------
                            Case "Bedarfsberechnung-18599"
                                '-------------------------------------------------------
                                Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksOben, .Energieausweis_Daten.Wohngebaeude.Gebaeudetyp.Wert & ", " & .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_18599.Wohngebaeude_Anbaugrad.Wert, 1464 + 20, 3241, 1065, 1250, Font_Schriftgroesse_40, Font_Druckfarbe)
                            '-------------------------------------------------------
                            Case "Bedarfsberechnung-easy"
                                '-------------------------------------------------------
                                Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksOben, .Energieausweis_Daten.Wohngebaeude.Gebaeudetyp.Wert & ", " & .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_easy.Wohngebaeude_Anbaugrad.Wert, 1464 + 20, 3241, 1065, 1250, Font_Schriftgroesse_40, Font_Druckfarbe)
                                '-------------------------------------------------------
                        End Select
                    '-------------------------------------------------------
                    Case "Energieverbrauchsausweis"
                        '-------------------------------------------------------
                        Select Case .Berechnungsverfahren
                            Case "Verbrauchsberechnung"
                                '-------------------------------------------------------
                                Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksOben, .Energieausweis_Daten.Wohngebaeude.Gebaeudetyp.Wert, 1464 + 20, 3241, 1065, 1250, Font_Schriftgroesse_40, Font_Druckfarbe)
                                '-------------------------------------------------------
                        End Select
                        '-------------------------------------------------------
                End Select
                '-------------------------------------------------------
                Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksMitte, .Gebaeudebezogene_Daten.Gebaeudeadresse_Strasse_Nr.Wert & ", " & .Gebaeudebezogene_Daten.Gebaeudeadresse_Postleitzahl.Wert & " " & .Gebaeudebezogene_Daten.Gebaeudeadresse_Ort.Wert, 1464 + 20, 3241, 1254, 1441, Font_Schriftgroesse_40, Font_Druckfarbe)
                Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Energieausweis_Daten.Gebaeudeteil.Wert, 1464 + 20, 3241, 1445, 1542, Font_Schriftgroesse_40, Font_Druckfarbe)
                Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Energieausweis_Daten.Baujahr_Gebaeude.Wert, 1464 + 20, 3241, 1546, 1643, Font_Schriftgroesse_40, Font_Druckfarbe)
                Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Energieausweis_Daten.Baujahr_Waermeerzeuger.Wert, 1464 + 20, 3241, 1647, 1832, Font_Schriftgroesse_40, Font_Druckfarbe)
                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Energieausweis_Daten.Wohngebaeude.Anzahl_Wohneinheiten.Wert, 0, "", 1464 + 20, 1897, 1836, 1932, Font_Schriftgroesse_40, Font_Druckfarbe)
                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Gebaeudenutzflaeche.Wert, 2, "m²", 1464 + 20, 1897, 1938, 2034, Font_Schriftgroesse_40, Font_Druckfarbe)
                '-------------------------------------------------------
                Select Case Variable_Steuerung.Berechnungsart
                    Case "Energieverbrauchsausweis"
                        Select Case .Berechnungsverfahren
                            Case "Verbrauchsberechnung"
                                '-------------------------------------------------------
                                Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Wohngebaeude.Verbrauchswerte.Flaechenermittlung_AN_aus_Wohnflaeche.Wert, 1947, 1977, 1973, 2003)
                                '-------------------------------------------------------
                        End Select
                End Select
                '-------------------------------------------------------
                Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksMitte, .Energieausweis_Daten.wesentliche_Energietraeger_Heizung.Wert, 1464 + 20, 4031, 2038, 2132, Font_Schriftgroesse_40, Font_Druckfarbe)
                Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksMitte, .Energieausweis_Daten.wesentliche_Energietraeger_Warmwasser.Wert, 1464 + 20, 4031, 2136, 2230, Font_Schriftgroesse_40, Font_Druckfarbe)
                Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksMitte, "Art: " & .Energieausweis_Daten.Erneuerbare_Art.Wert, 1464 + 20, 2483, 2234, 2420, Font_Schriftgroesse_40, Font_Druckfarbe)
                Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksMitte, "Verwendung: " & .Energieausweis_Daten.Erneuerbare_Verwendung.Wert, 2489 + 20, 4031, 2234, 2420, Font_Schriftgroesse_40, Font_Druckfarbe)
                ''--------------------------------------------------
                Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Lueftungsart_Fensterlueftung.Wert, 1509, 1539, 2472, 2502)
                Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Lueftungsart_Schachtlueftung.Wert, 1509, 1539, 2550, 2579)
                Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Lueftungsart_Anlage_m_WRG.Wert, 2511, 2541, 2472, 2502)
                Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Lueftungsart_Anlage_o_WRG.Wert, 2511, 2541, 2550, 2579)
                '-------------------------------------------------------            
                Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Kuehlungsart_passive_Kuehlung.Wert, 1509, 1539, 2674, 2704)
                Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Kuehlungsart_gelieferte_Kaelte.Wert, 1509, 1539, 2752, 2781)
                Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Kuehlungsart_Strom.Wert, 2511, 2541, 2674, 2704)
                Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Kuehlungsart_Waerme.Wert, 2511, 2541, 2752, 2781)
                '-------------------------------------------------------
                If .Energieausweis_Daten.Keine_inspektionspflichtige_Anlage.Wert = False And .Energieausweis_Daten.Anzahl_Klimanlagen.Wert > 0 Then
                    Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Energieausweis_Daten.Anzahl_Klimanlagen.Wert, 0, "Stck.", 1683 + 20, 2109, 2830 + 15, 2926, Font_Schriftgroesse_40, Font_Druckfarbe)
                    Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, FormatDateTime(.Energieausweis_Daten.Faelligkeitsdatum_Inspektion.Wert, DateFormat.ShortDate), 3206 + 20, 4031, 2830 + 15, 2926, Font_Schriftgroesse_40, Font_Druckfarbe)
                End If
                '-------------------------------------------------------
                Select Case .Energieausweis_Daten.Ausstellungsanlass.Wert
                    Case "Neubau"
                        Image_Auswahl_schreiben(Grafik_Seite, True, 1509, 1539, 2978, 3007)
                        Image_Auswahl_schreiben(Grafik_Seite, False, 2437, 2466, 2978, 3007)
                        Image_Auswahl_schreiben(Grafik_Seite, False, 1509, 1539, 3056, 3089)
                        Image_Auswahl_schreiben(Grafik_Seite, False, 3287, 3327, 2978, 3007)
                    Case "Modernisierung-Erweiterung"
                        Image_Auswahl_schreiben(Grafik_Seite, False, 1509, 1539, 2978, 3007)
                        Image_Auswahl_schreiben(Grafik_Seite, True, 2437, 2466, 2978, 3007)
                        Image_Auswahl_schreiben(Grafik_Seite, False, 1509, 1539, 3056, 3089)
                        Image_Auswahl_schreiben(Grafik_Seite, False, 3287, 3327, 2978, 3007)
                    Case "Vermietung-Verkauf"
                        Image_Auswahl_schreiben(Grafik_Seite, False, 1509, 1539, 2978, 3007)
                        Image_Auswahl_schreiben(Grafik_Seite, False, 2437, 2466, 2978, 3007)
                        Image_Auswahl_schreiben(Grafik_Seite, True, 1509, 1539, 3056, 3089)
                        Image_Auswahl_schreiben(Grafik_Seite, False, 3287, 3327, 2978, 3007)
                    Case "Sonstiges"
                        Image_Auswahl_schreiben(Grafik_Seite, False, 1509, 1539, 2978, 3007)
                        Image_Auswahl_schreiben(Grafik_Seite, False, 2437, 2466, 2978, 3007)
                        Image_Auswahl_schreiben(Grafik_Seite, False, 1509, 1539, 3056, 3089)
                        Image_Auswahl_schreiben(Grafik_Seite, True, 3287, 3327, 2978, 3007)
                End Select
                '-------------------------------------------------------
                Select Case Variable_Steuerung.Berechnungsart
                    Case "Energiebedarfsausweis"
                        Image_Auswahl_schreiben(Grafik_Seite, True, 319, 349, 3771, 3800)
                        Image_Auswahl_schreiben(Grafik_Seite, False, 319, 349, 3971, 4000)
                    Case "Energieverbrauchsausweis"
                        Image_Auswahl_schreiben(Grafik_Seite, False, 319, 349, 3771, 3800)
                        Image_Auswahl_schreiben(Grafik_Seite, True, 319, 349, 3971, 4000)
                End Select
                '-------------------------------------------------------
                Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Datenerhebung_Eigentuemer.Wert, 2118, 2148, 4176, 4206)
                Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Datenerhebung_Aussteller.Wert, 3177, 3207, 4176, 4206)
                '-------------------------------------------------------
                Image_Auswahl_schreiben(Grafik_Seite, .Gebaeudebezogene_Daten.Zusatzinfos_beigefuegt.Wert, 317, 347, 4292, 4323)
                '-------------------------------------------------------
                Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, FormatDateTime(Variable_Steuerung.Ausstellungsdatum, DateFormat.ShortDate).ToString, 3545, 3983, 5364, 5458, Font_Schriftgroesse_50, Font_Druckfarbe)
                '-------------------------------------------------------
                Image_Grafik_Projekt_schreiben(Grafik_Seite, Variable_Steuerung.Bild_Projekt, 3241, 1065, 4031 - 3241, 2038 - 1065)
                '-------------------------------------------------------
                Image_Grafik_Aussteller_schreiben(Grafik_Seite, Variable_Steuerung.Bild_Aussteller, .Gebaeudebezogene_Daten.Aussteller_Bezeichnung.Wert, .Gebaeudebezogene_Daten.Ausstellervorname.Wert, .Gebaeudebezogene_Daten.Ausstellername.Wert, .Gebaeudebezogene_Daten.Aussteller_Strasse_Nr.Wert, .Gebaeudebezogene_Daten.Aussteller_PLZ.Wert, .Gebaeudebezogene_Daten.Aussteller_Ort.Wert, 300, 5071, 2500, 300)
                '-------------------------------------------------------
                Image_Grafik_Unterschrift_schreiben(Grafik_Seite, Variable_Steuerung.Bild_Unterschrift, 3000, 5060, 1000, 300)
                '-------------------------------------------------------
            End With
            '--------------------------------------------------
            If Variable_Steuerung.Entwurf = True Then
                Image_Text_im_Winkel_schreiben(Grafik_Seite, -45, "ENTWURF", 0, 2100, 0, 2970, Font_Schriftgroesse_ENTWURF, Font_Druckfarbe_ENTWURF)
            End If
            '-------------------------------------------------------
            If Oberflaeche = True Then
                '-------------------------------------------------------
                PictureBox.Image = Bitmap_Seite
                PictureBox.Update()
                '-------------------------------------------------------
                Imagelist.Images.Add(Bitmap_skalieren(Bitmap_Seite, Skalierungsfaktor))
                Listview.Items.Add("Seite " & Seitennummer, Seitennummer - 1)
                Listview.Update()
                '-------------------------------------------------------
            End If
            '-------------------------------------------------------
            Variable_XML_Import.Bilddateien(Seitennummer) = Image_to_String(Bitmap_Seite)
            Variable_XML_Import.Bildanzahl = 1
            '-------------------------------------------------------
            Speicher_freigeben()
            '-------------------------------------------------------
#End Region
            '-------------------------------------------------------
#Region "Seite 2"
            '-------------------------------------------------------
            'Seite 2 -----------------------------------------------
            '-------------------------------------------------------
            Seitennummer += 1
            '-------------------------------------------------------
            Grafik_Seite.Clear(Color.White)
            '-------------------------------------------------------
            Hintergrundgrafik = Image.FromFile(Variable_Parameter.Arbeitsverzeichnis & "/Image/GEG2020_WB_Seite_2.png")
            Grafik_Seite.DrawImage(Hintergrundgrafik, 0, 0, 4200, 5940)
            '-------------------------------------------------------
            Dim Summe_Deckungsanteil As Integer = 0
            Dim Summe_Anteil_Pflichterfuellung As Integer = 0
            '-------------------------------------------------------
            With Variable_XML_Import
                Select Case Variable_Steuerung.Berechnungsart
                    Case "Energiebedarfsausweis"
                        '-------------------------------------------------------
                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, Datum_Energiegesetz, 2128, 2616, 422, 516, Font_Schriftgroesse_50, Font_Druckfarbe)
                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.Mitte_kein_Abschneiden, Variable_Steuerung.Registriernummer, 3042, 3758, 722, 822, Font_Schriftgroesse_50, Font_Druckfarbe)
                        '-------------------------------------------------------
                        Select Case .Berechnungsverfahren
                            Case "Bedarfsberechnung-4108-4701"
                                '-------------------------------------------------------
                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_4108_4701.Endenergiebedarf_Waerme_AN.Wert + .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_4108_4701.Endenergiebedarf_Hilfsenergie_AN.Wert, 1, "", 3103, 3643, 3013, 3184, Font_Schriftgroesse_60, Font_Druckfarbe)
                                '-------------------------------------------------------
                                Select Case .Energieausweis_Daten.Ausstellungsanlass.Wert
                                    Case "Neubau", "Modernisierung-Erweiterung"
                                        '-------------------------------------------------------
                                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Treibhausgasemissionen.Wert, 2, "", 2622, 3025, 1124, 1226, Font_Schriftgroesse_50, Font_Druckfarbe)
                                        '-------------------------------------------------------
                                        Select Case Variable_Steuerung.Neubau
                                            Case 0 'Modernisierung
                                                '-------------------------------------------------------
                                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_4108_4701.Primaerenergiebedarf_Hoechstwert_Bestand.Wert, 1, "", 1657, 1880, 2416, 2518, Font_Schriftgroesse_40, Font_Druckfarbe)
                                                '-------------------------------------------------------
                                            Case 1 'Neubau
                                                '-------------------------------------------------------
                                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_4108_4701.Primaerenergiebedarf_Hoechstwert_Neubau.Wert, 1, "", 1657, 1880, 2416, 2518, Font_Schriftgroesse_40, Font_Druckfarbe)
                                                '-------------------------------------------------------
                                        End Select
                                        '-------------------------------------------------------
                                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_4108_4701.spezifischer_Transmissionswaermeverlust_Ist.Wert, 2, "", 531, 753, 2668, 2771, Font_Schriftgroesse_40, Font_Druckfarbe)
                                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_4108_4701.spezifischer_Transmissionswaermeverlust_Hoechstwert.Wert, 2, "", 1657, 1880, 2668, 2771, Font_Schriftgroesse_40, Font_Druckfarbe)
                                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_4108_4701.Primaerenergiebedarf.Wert, 1, "", 531, 753, 2416, 2518, Font_Schriftgroesse_40, Font_Druckfarbe)
                                        '-------------------------------------------------------
                                        Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_4108_4701.Sommerlicher_Waermeschutz.Wert, 1387, 1417, 2863, 2893)
                                        '-------------------------------------------------------
                                    Case "Vermietung-Verkauf", "Sonstiges"
                                        '-------------------------------------------------------
                                End Select
                                '-------------------------------------------------------
                                Image_Auswahl_schreiben(Grafik_Seite, True, 2346, 2376, 2311, 2341)
                                Image_Auswahl_schreiben(Grafik_Seite, False, 2346, 2376, 2396, 2426)
                                Image_Auswahl_schreiben(Grafik_Seite, False, 2346, 2376, 2480, 2510)
                                '-------------------------------------------------------
                                Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_4108_4701.Vereinfachte_Datenaufnahme.Wert, 2346, 2376, 2565, 2595)
                                '-------------------------------------------------------
                                Image_Pfeil_WB(Grafik_Seite, Lage.Oben, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_4108_4701.Endenergiebedarf_Waerme_AN.Wert + .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_4108_4701.Endenergiebedarf_Hilfsenergie_AN.Wert, 1, "Endenergiebedarf dieses Gebäudes", 543, 3623, 1535, 1874)
                                '-------------------------------------------------------
                                Image_Pfeil_WB(Grafik_Seite, Lage.Unten, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_4108_4701.Primaerenergiebedarf.Wert, 1, "Primärenergiebedarf dieses Gebäudes", 543, 3623, 1535, 1874)
                                '-------------------------------------------------------
                                'Erneuerbare Energien ----------------------------------
                                '-------------------------------------------------------
                                Select Case .Energieausweis_Daten.Ausstellungsanlass.Wert
                                    Case "Neubau"
                                        '-------------------------------------------------------
                                        If .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_4108_4701.Art_der_Nutzung_erneuerbaren_Energie_1.Wert <> "" Then
                                            '-------------------------------------------------------
                                            Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_4108_4701.Art_der_Nutzung_erneuerbaren_Energie_1.Wert, 281, 1463, 3831, 3934, Font_Schriftgroesse_40, Font_Druckfarbe)
                                            Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_4108_4701.Deckungsanteil_1.Wert, 0, "", 1478, 1682, 3831, 3934, Font_Schriftgroesse_40, Font_Druckfarbe)
                                            Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_4108_4701.Anteil_der_Pflichterfuellung_1.Wert, 0, "", 1781, 1988, 3831, 3934, Font_Schriftgroesse_40, Font_Druckfarbe)
                                            '-------------------------------------------------------
                                            Summe_Deckungsanteil += .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_4108_4701.Deckungsanteil_1.Wert
                                            Summe_Anteil_Pflichterfuellung += .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_4108_4701.Anteil_der_Pflichterfuellung_1.Wert
                                            '-------------------------------------------------------
                                        End If
                                        If .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_4108_4701.Art_der_Nutzung_erneuerbaren_Energie_2.Wert <> "" Then
                                            '-------------------------------------------------------
                                            Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_4108_4701.Art_der_Nutzung_erneuerbaren_Energie_2.Wert, 281, 1463, 3943, 4040, Font_Schriftgroesse_40, Font_Druckfarbe)
                                            Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_4108_4701.Deckungsanteil_2.Wert, 0, "", 1478, 1682, 3943, 4040, Font_Schriftgroesse_40, Font_Druckfarbe)
                                            Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_4108_4701.Anteil_der_Pflichterfuellung_2.Wert, 0, "", 1781, 1988, 3943, 4040, Font_Schriftgroesse_40, Font_Druckfarbe)
                                            '-------------------------------------------------------
                                            Summe_Deckungsanteil += .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_4108_4701.Deckungsanteil_2.Wert
                                            Summe_Anteil_Pflichterfuellung += .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_4108_4701.Anteil_der_Pflichterfuellung_2.Wert
                                            '-------------------------------------------------------
                                        End If
                                        '-------------------------------------------------------
                                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, Summe_Deckungsanteil, 0, "", 1478, 1682, 4050, 4153, Font_Schriftgroesse_40, Font_Druckfarbe)
                                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, Summe_Anteil_Pflichterfuellung, 0, "", 1781, 1988, 4050, 4153, Font_Schriftgroesse_40, Font_Druckfarbe)
                                        '-------------------------------------------------------
                                        Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_4108_4701.verschaerft_nach_GEG_45_eingehalten.Wert, 306, 347, 4688, 4730)
                                        If .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_4108_4701.nicht_verschaerft_nach_GEG_34.Wert = False Then
                                            Image_Auswahl_schreiben(Grafik_Seite, True, 306, 347, 4873, 4915)
                                            Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_4108_4701.Anforderung_nach_GEG_16_unterschritten.Wert, 0, "", 1720, 1836, 4934, 4995, Font_Schriftgroesse_40, Font_Druckfarbe)
                                            Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_4108_4701.verschaerft_nach_GEG_34.Wert, 0, "", 1428, 1541, 5005, 5065, Font_Schriftgroesse_40, Font_Druckfarbe)
                                        Else
                                            Image_Auswahl_schreiben(Grafik_Seite, False, 306, 347, 4873, 4915)
                                        End If
                                        '-------------------------------------------------------
                                    Case "Modernisierung-Erweiterung", "Vermietung-Verkauf", "Sonstiges"
                                        '-------------------------------------------------------
                                End Select
                                '-------------------------------------------------------
                            Case "Bedarfsberechnung-18599"
                                '-------------------------------------------------------
                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_18599.Endenergiebedarf_Gesamt.Wert, 1, "", 3103, 3643, 3013, 3184, Font_Schriftgroesse_60, Font_Druckfarbe)
                                '-------------------------------------------------------
                                Select Case .Energieausweis_Daten.Ausstellungsanlass.Wert
                                    Case "Neubau", "Modernisierung-Erweiterung"
                                        '-------------------------------------------------------
                                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Treibhausgasemissionen.Wert, 2, "", 2622, 3025, 1124, 1226, Font_Schriftgroesse_50, Font_Druckfarbe)
                                        '-------------------------------------------------------
                                        Select Case Variable_Steuerung.Neubau
                                            Case 0 'Modernisierung
                                                '-------------------------------------------------------
                                                'Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_18599.Treibhausgasemissionen_Hoechstwert_Bestand.Wert, 2, "", 2622, 3025, 1124, 1226, Font_Schriftgroesse_50, Font_Druckfarbe)
                                                '-------------------------------------------------------
                                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_18599.Primaerenergiebedarf_Hoechstwert_Bestand.Wert, 1, "", 1657, 1880, 2416, 2518, Font_Schriftgroesse_40, Font_Druckfarbe)
                                                '-------------------------------------------------------
                                            Case 1 'Neubau
                                                '-------------------------------------------------------
                                                'Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_18599.Treibhausgasemissionen_Hoechstwert_Neubau.Wert, 2, "", 2622, 3025, 1124, 1226, Font_Schriftgroesse_50, Font_Druckfarbe)
                                                '-------------------------------------------------------
                                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_18599.Primaerenergiebedarf_Hoechstwert_Neubau.Wert, 1, "", 1657, 1880, 2416, 2518, Font_Schriftgroesse_40, Font_Druckfarbe)
                                                '-------------------------------------------------------
                                        End Select
                                        '-------------------------------------------------------
                                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_18599.spezifischer_Transmissionswaermetransferkoeffizient_Ist.Wert, 2, "", 531, 753, 2668, 2771, Font_Schriftgroesse_40, Font_Druckfarbe)
                                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_18599.spezifischer_Transmissionswaermetransferkoeffizient_Hoechstwert.Wert, 2, "", 1657, 1880, 2668, 2771, Font_Schriftgroesse_40, Font_Druckfarbe)
                                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_18599.Primaerenergiebedarf_AN.Wert, 1, "", 531, 753, 2416, 2518, Font_Schriftgroesse_40, Font_Druckfarbe)
                                        '-------------------------------------------------------
                                        Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_18599.Sommerlicher_Waermeschutz.Wert, 1387, 1417, 2863, 2893)
                                        '-------------------------------------------------------
                                    Case "Vermietung-Verkauf", "Sonstiges"
                                        '-------------------------------------------------------
                                End Select
                                '-------------------------------------------------------
                                Image_Auswahl_schreiben(Grafik_Seite, False, 2346, 2376, 2311, 2341)
                                Image_Auswahl_schreiben(Grafik_Seite, True, 2346, 2376, 2396, 2426)
                                Image_Auswahl_schreiben(Grafik_Seite, False, 2346, 2376, 2480, 2510)
                                '-------------------------------------------------------
                                Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_18599.Vereinfachte_Datenaufnahme.Wert, 2346, 2376, 2565, 2595)
                                '-------------------------------------------------------
                                Image_Pfeil_WB(Grafik_Seite, Lage.Oben, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_18599.Endenergiebedarf_Gesamt.Wert, 1, "Endenergiebedarf dieses Gebäudes", 543, 3623, 1535, 1874)
                                '-------------------------------------------------------
                                Image_Pfeil_WB(Grafik_Seite, Lage.Unten, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_18599.Primaerenergiebedarf_AN.Wert, 1, "Primärenergiebedarf dieses Gebäudes", 543, 3623, 1535, 1874)
                                '-------------------------------------------------------
                                'Erneuerbare Energien ----------------------------------
                                '-------------------------------------------------------
                                Select Case .Energieausweis_Daten.Ausstellungsanlass.Wert
                                    Case "Neubau"
                                        '-------------------------------------------------------
                                        If .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_18599.Art_der_Nutzung_erneuerbaren_Energie_1.Wert <> "" Then
                                            '-------------------------------------------------------
                                            Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_18599.Art_der_Nutzung_erneuerbaren_Energie_1.Wert, 281, 1463, 3831, 3934, Font_Schriftgroesse_40, Font_Druckfarbe)
                                            Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_18599.Deckungsanteil_1.Wert, 0, "", 1478, 1682, 3831, 3934, Font_Schriftgroesse_40, Font_Druckfarbe)
                                            Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_18599.Anteil_der_Pflichterfuellung_1.Wert, 0, "", 1781, 1988, 3831, 3934, Font_Schriftgroesse_40, Font_Druckfarbe)
                                            '-------------------------------------------------------
                                            Summe_Deckungsanteil += .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_18599.Deckungsanteil_1.Wert
                                            Summe_Anteil_Pflichterfuellung += .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_18599.Anteil_der_Pflichterfuellung_1.Wert
                                            '-------------------------------------------------------
                                        End If
                                        If .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_18599.Art_der_Nutzung_erneuerbaren_Energie_2.Wert <> "" Then
                                            '-------------------------------------------------------
                                            Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_18599.Art_der_Nutzung_erneuerbaren_Energie_2.Wert, 281, 1463, 3943, 4040, Font_Schriftgroesse_40, Font_Druckfarbe)
                                            Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_18599.Deckungsanteil_2.Wert, 0, "", 1478, 1682, 3943, 4040, Font_Schriftgroesse_40, Font_Druckfarbe)
                                            Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_18599.Anteil_der_Pflichterfuellung_2.Wert, 0, "", 1781, 1988, 3943, 4040, Font_Schriftgroesse_40, Font_Druckfarbe)
                                            '-------------------------------------------------------
                                            Summe_Deckungsanteil += .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_18599.Deckungsanteil_2.Wert
                                            Summe_Anteil_Pflichterfuellung += .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_18599.Anteil_der_Pflichterfuellung_2.Wert
                                            '-------------------------------------------------------
                                        End If
                                        '-------------------------------------------------------
                                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, Summe_Deckungsanteil, 0, "", 1478, 1682, 4050, 4153, Font_Schriftgroesse_40, Font_Druckfarbe)
                                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, Summe_Anteil_Pflichterfuellung, 0, "", 1781, 1988, 4050, 4153, Font_Schriftgroesse_40, Font_Druckfarbe)
                                        '-------------------------------------------------------
                                        Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_18599.verschaerft_nach_GEG_45_eingehalten.Wert, 306, 347, 4688, 4730)
                                        If .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_18599.nicht_verschaerft_nach_GEG_34.Wert = False Then
                                            Image_Auswahl_schreiben(Grafik_Seite, True, 306, 347, 4873, 4915)
                                            Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_18599.Anforderung_nach_GEG_16_unterschritten.Wert, 0, "", 1720, 1836, 4934, 4995, Font_Schriftgroesse_40, Font_Druckfarbe)
                                            Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_18599.verschaerft_nach_GEG_34.Wert, 0, "", 1428, 1541, 5005, 5065, Font_Schriftgroesse_40, Font_Druckfarbe)
                                        Else
                                            Image_Auswahl_schreiben(Grafik_Seite, False, 306, 347, 4873, 4915)
                                        End If
                                        '-------------------------------------------------------
                                    Case "Modernisierung-Erweiterung", "Vermietung-Verkauf", "Sonstiges"
                                        '-------------------------------------------------------
                                End Select
                                '-------------------------------------------------------
                            Case "Bedarfsberechnung-easy"
                                '-------------------------------------------------------
                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Treibhausgasemissionen.Wert, 2, "", 2622, 3025, 1124, 1226, Font_Schriftgroesse_50, Font_Druckfarbe)
                                '-------------------------------------------------------
                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_easy.Primaerenergiebedarf_Anforderungswert.Wert, 1, "", 1657, 1880, 2416, 2518, Font_Schriftgroesse_40, Font_Druckfarbe)
                                '-------------------------------------------------------
                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_easy.Primaerenergiebedarf_Ist_Wert.Wert, 1, "", 531, 753, 2416, 2518, Font_Schriftgroesse_40, Font_Druckfarbe)
                                '-------------------------------------------------------
                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_easy.Energetische_Qualitaet_Ist_Wert.Wert, 2, "", 531, 753, 2668, 2771, Font_Schriftgroesse_40, Font_Druckfarbe)
                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_easy.Energetische_Qualitaet_Anforderungs_Wert.Wert, 2, "", 1657, 1880, 2668, 2771, Font_Schriftgroesse_40, Font_Druckfarbe)
                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_easy.Endenergiebedarf.Wert, 1, "", 3103, 3643, 3013, 3184, Font_Schriftgroesse_60, Font_Druckfarbe)
                                '-------------------------------------------------------
                                Image_Auswahl_schreiben(Grafik_Seite, False, 1387, 1417, 2863, 2893)
                                '-------------------------------------------------------
                                Image_Auswahl_schreiben(Grafik_Seite, False, 2346, 2376, 2311, 2341)
                                Image_Auswahl_schreiben(Grafik_Seite, False, 2346, 2376, 2396, 2426)
                                Image_Auswahl_schreiben(Grafik_Seite, True, 2346, 2376, 2480, 2510)
                                '-------------------------------------------------------
                                Image_Auswahl_schreiben(Grafik_Seite, False, 2346, 2376, 2565, 2595)
                                '-------------------------------------------------------
                                Image_Pfeil_WB(Grafik_Seite, Lage.Oben, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_easy.Endenergiebedarf.Wert, 1, "Endenergiebedarf dieses Gebäudes", 543, 3623, 1535, 1874)
                                '-------------------------------------------------------
                                Image_Pfeil_WB(Grafik_Seite, Lage.Unten, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_easy.Primaerenergiebedarf_Ist_Wert.Wert, 1, "Primärenergiebedarf dieses Gebäudes", 543, 3623, 1535, 1874)
                                '-------------------------------------------------------
                                'Erneuerbare Energien
                                '-------------------------------------------------------
                                If .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_easy.Art_der_Nutzung_erneuerbaren_Energie_1.Wert <> "" Then
                                    '-------------------------------------------------------
                                    Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_easy.Art_der_Nutzung_erneuerbaren_Energie_1.Wert, 281, 1463, 3831, 3934, Font_Schriftgroesse_40, Font_Druckfarbe)
                                    Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_easy.Deckungsanteil_1.Wert, 0, "", 1478, 1682, 3831, 3934, Font_Schriftgroesse_40, Font_Druckfarbe)
                                    Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_easy.Anteil_der_Pflichterfuellung_1.Wert, 0, "", 1781, 1988, 3831, 3934, Font_Schriftgroesse_40, Font_Druckfarbe)
                                    '-------------------------------------------------------
                                    Summe_Deckungsanteil += .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_easy.Deckungsanteil_1.Wert
                                    Summe_Anteil_Pflichterfuellung += .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_easy.Anteil_der_Pflichterfuellung_1.Wert
                                    '-------------------------------------------------------
                                End If
                                If .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_easy.Art_der_Nutzung_erneuerbaren_Energie_2.Wert <> "" Then
                                    '-------------------------------------------------------
                                    Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_easy.Art_der_Nutzung_erneuerbaren_Energie_2.Wert, 281, 1463, 3943, 4040, Font_Schriftgroesse_40, Font_Druckfarbe)
                                    Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_easy.Deckungsanteil_2.Wert, 0, "", 1478, 1682, 3943, 4040, Font_Schriftgroesse_40, Font_Druckfarbe)
                                    Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_easy.Anteil_der_Pflichterfuellung_2.Wert, 0, "", 1781, 1988, 3943, 4040, Font_Schriftgroesse_40, Font_Druckfarbe)
                                    '-------------------------------------------------------
                                    Summe_Deckungsanteil += .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_easy.Deckungsanteil_2.Wert
                                    Summe_Anteil_Pflichterfuellung += .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_easy.Anteil_der_Pflichterfuellung_2.Wert
                                    '-------------------------------------------------------
                                End If
                                '-------------------------------------------------------
                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, Summe_Deckungsanteil, 0, "", 1478, 1682, 4050, 4153, Font_Schriftgroesse_40, Font_Druckfarbe)
                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, Summe_Anteil_Pflichterfuellung, 0, "", 1781, 1988, 4050, 4153, Font_Schriftgroesse_40, Font_Druckfarbe)
                                '-------------------------------------------------------
                        End Select
                        '-------------------------------------------------------
                End Select
                '--------------------------------------------------
            End With
            '--------------------------------------------------
            If Variable_Steuerung.Entwurf = True Then
                Image_Text_im_Winkel_schreiben(Grafik_Seite, -45, "ENTWURF", 0, 2100, 0, 2970, Font_Schriftgroesse_ENTWURF, Font_Druckfarbe_ENTWURF)
            End If
            '-------------------------------------------------------
            If Oberflaeche = True Then
                '-------------------------------------------------------
                Imagelist.Images.Add(Bitmap_skalieren(Bitmap_Seite, Skalierungsfaktor))
                Listview.Items.Add("Seite " & Seitennummer, Seitennummer - 1)
                Listview.Update()
                '-------------------------------------------------------
            End If
            '-------------------------------------------------------
            Variable_XML_Import.Bilddateien(Seitennummer) = Image_to_String(Bitmap_Seite)
            Variable_XML_Import.Bildanzahl += 1
            '-------------------------------------------------------
            Speicher_freigeben()
            '-------------------------------------------------------
#End Region
            '-------------------------------------------------------
#Region "Seite 3"
            '-------------------------------------------------------
            'Seite 3 -----------------------------------------------
            '-------------------------------------------------------
            Seitennummer += 1
            '-------------------------------------------------------
            Grafik_Seite.Clear(Color.White)
            '-------------------------------------------------------
            Hintergrundgrafik = Image.FromFile(Variable_Parameter.Arbeitsverzeichnis & "/Image/GEG2020_WB_Seite_3.png")
            Grafik_Seite.DrawImage(Hintergrundgrafik, 0, 0, 4200, 5940)
            '-------------------------------------------------------
            With Variable_XML_Import
                '-------------------------------------------------------
                If Variable_Steuerung.Berechnungsart = "Energieverbrauchsausweis" Or Variable_XML_Import.Zusaetzliche_Verbrauchsberechnung = True Then
                    '-------------------------------------------------------
                    Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, Datum_Energiegesetz, 2128, 2616, 422, 516, Font_Schriftgroesse_50, Font_Druckfarbe)
                    Image_Text_schreiben(Grafik_Seite, Ausrichtung.Mitte_kein_Abschneiden, Variable_Steuerung.Registriernummer, 3042, 3758, 722, 822, Font_Schriftgroesse_50, Font_Druckfarbe)
                    '-------------------------------------------------------
                    If .Berechnungsverfahren = "Verbrauchsberechnung" Or Variable_XML_Import.Zusaetzliche_Verbrauchsberechnung = True Then
                        '-------------------------------------------------------
                        If Variable_XML_Import.Zusaetzliche_Verbrauchsberechnung = False Then
                            Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Treibhausgasemissionen.Wert, 2, "", 2649, 2952, 1105, 1198, Font_Schriftgroesse_50, Font_Druckfarbe)
                        Else
                            Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Treibhausgasemissionen_Zusaetzliche_Verbrauchsdaten.Wert, 2, "", 2649, 2952, 1105, 1198, Font_Schriftgroesse_50, Font_Druckfarbe)
                        End If
                        '-------------------------------------------------------
                        Image_Pfeil_WB(Grafik_Seite, Lage.Oben, .Energieausweis_Daten.Wohngebaeude.Verbrauchswerte.Mittlerer_Endenergieverbrauch.Wert, 1, "Endenergieverbrauch dieses Gebäudes", 550, 3630, 1517, 1856)
                        '-------------------------------------------------------
                        Image_Pfeil_WB(Grafik_Seite, Lage.Unten, .Energieausweis_Daten.Wohngebaeude.Verbrauchswerte.Mittlerer_Primaerenergieverbrauch.Wert, 1, "Primärenergieverbrauch dieses Gebäudes", 550, 3630, 1517, 1856)
                        '-------------------------------------------------------
                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Verbrauchswerte.Mittlerer_Endenergieverbrauch.Wert, 1, "", 3159, 3661, 2195, 2357, Font_Schriftgroesse_60, Font_Druckfarbe)
                        '-------------------------------------------------------
                        Zusammenstellung_WB_Verbrauchserfassung()
                        '-------------------------------------------------------
                        Zeilenzahl = 6
                        Zeilenhoehe = (3414 - 2815) / 6
                        '-------------------------------------------------------
                        If .Verbrauchserfassung_Anzahl < 6 Then
                            Zeilenzahl = .Verbrauchserfassung_Anzahl
                        End If
                        '-------------------------------------------------------
                        For Anzahl_Zeilen = 1 To Zeilenzahl
                            '-------------------------------------------------------
                            Image_Datum_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Verbrauchserfassung(Anzahl_Zeilen).Startdatum.Wert, 245, 572, 2815 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 2815 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                            '--------------------------------------------------
                            Image_Datum_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Verbrauchserfassung(Anzahl_Zeilen).Enddatum.Wert, 576, 902, 2815 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 2815 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                            '--------------------------------------------------
                            Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Verbrauchserfassung(Anzahl_Zeilen).Energietraeger_Verbrauch.Wert, 906, 2469, 2815 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 2815 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                            '--------------------------------------------------
                            Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Verbrauchserfassung(Anzahl_Zeilen).Primaerenergiefaktor.Wert, 1, "", 2473, 2709, 2815 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 2815 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                            '--------------------------------------------------
                            If .Verbrauchserfassung(Anzahl_Zeilen).Energieverbrauch.Wert <> 0 Then
                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Verbrauchserfassung(Anzahl_Zeilen).Energieverbrauch.Wert, 0, "", 2713, 3069, 2815 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 2815 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                            End If
                            '--------------------------------------------------
                            If .Verbrauchserfassung(Anzahl_Zeilen).Energieverbrauchsanteil_Warmwasser_zentral.Wert <> 0 Then
                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Verbrauchserfassung(Anzahl_Zeilen).Energieverbrauchsanteil_Warmwasser_zentral.Wert, 0, "", 3073, 3429, 2815 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 2815 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                            End If
                            '--------------------------------------------------
                            If .Verbrauchserfassung(Anzahl_Zeilen).Energieverbrauchsanteil_Heizung.Wert <> 0 Then
                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Verbrauchserfassung(Anzahl_Zeilen).Energieverbrauchsanteil_Heizung.Wert, 0, "", 3433, 3789, 2815 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 2815 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                            End If
                            '--------------------------------------------------
                            If .Verbrauchserfassung(Anzahl_Zeilen).Klimafaktor.Wert <> 0 Then
                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Verbrauchserfassung(Anzahl_Zeilen).Klimafaktor.Wert, 2, "", 3793, 4032, 2815 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 2815 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                            End If
                            '-------------------------------------------------------
                        Next
                        '-------------------------------------------------------
                        For Anzahl_Zeilen = 1 To 6
                            '-------------------------------------------------------
                            Image_Linie(Grafik_Seite, 245, 4032, 2810 + (Zeilenhoehe * Anzahl_Zeilen), 2810 + (Zeilenhoehe * Anzahl_Zeilen))
                            '-------------------------------------------------------
                        Next
                        '-------------------------------------------------------
                        If .Verbrauchserfassung_Anzahl > 6 Then
                            '-------------------------------------------------------
                            Image_Auswahl_schreiben(Grafik_Seite, True, 300, 330, 3462, 3492)
                            '-------------------------------------------------------
                        Else
                            '-------------------------------------------------------
                            Image_Auswahl_schreiben(Grafik_Seite, False, 300, 330, 3462, 3492)
                            '-------------------------------------------------------
                        End If

                        '-------------------------------------------------------
                    End If
                    '-------------------------------------------------------
                End If
                '--------------------------------------------------
            End With
            '--------------------------------------------------
            If Variable_Steuerung.Entwurf = True Then
                Image_Text_im_Winkel_schreiben(Grafik_Seite, -45, "ENTWURF", 0, 2100, 0, 2970, Font_Schriftgroesse_ENTWURF, Font_Druckfarbe_ENTWURF)
            End If
            '-------------------------------------------------------
            If Oberflaeche = True Then
                '-------------------------------------------------------
                Imagelist.Images.Add(Bitmap_skalieren(Bitmap_Seite, Skalierungsfaktor))
                Listview.Items.Add("Seite " & Seitennummer, Seitennummer - 1)
                Listview.Update()
                '-------------------------------------------------------
            End If
            '-------------------------------------------------------
            Variable_XML_Import.Bilddateien(Seitennummer) = Image_to_String(Bitmap_Seite)
            Variable_XML_Import.Bildanzahl += 1
            '-------------------------------------------------------
            Speicher_freigeben()
            '-------------------------------------------------------
#End Region
            '-------------------------------------------------------
#Region "Seite 4"
            '-------------------------------------------------------
            'Seite 4 -----------------------------------------------
            '-------------------------------------------------------
            Seitennummer += 1
            '-------------------------------------------------------
            Grafik_Seite.Clear(Color.White)
            '-------------------------------------------------------
            Hintergrundgrafik = Image.FromFile(Variable_Parameter.Arbeitsverzeichnis & "/Image/GEG2020_WB_Seite_4.png")
            Grafik_Seite.DrawImage(Hintergrundgrafik, 0, 0, 4200, 5940)
            '-------------------------------------------------------
            With Variable_XML_Import
                '-------------------------------------------------------
                Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, Datum_Energiegesetz, 2128, 2616, 422, 516, Font_Schriftgroesse_50, Font_Druckfarbe)
                Image_Text_schreiben(Grafik_Seite, Ausrichtung.Mitte_kein_Abschneiden, Variable_Steuerung.Registriernummer, 3042, 3758, 722, 822, Font_Schriftgroesse_50, Font_Druckfarbe)
                '-------------------------------------------------------
                .Modernisierungsempfehlung_Anzahl = Anzahl_Modernisierungsempfehlung()
                '--------------------------------------------------
                If .Energieausweis_Daten.Empfehlungen_moeglich.Wert = False Then
                    '--------------------------------------------------
                    Image_Auswahl_schreiben(Grafik_Seite, False, 2425, 2472, 1134, 1180)
                    Image_Auswahl_schreiben(Grafik_Seite, True, 2833, 2879, 1134, 1180)
                    '--------------------------------------------------
                Else
                    '--------------------------------------------------
                    Image_Auswahl_schreiben(Grafik_Seite, True, 2425, 2472, 1134, 1180)
                    Image_Auswahl_schreiben(Grafik_Seite, False, 2833, 2879, 1134, 1180)
                    '--------------------------------------------------
                    Zeilenzahl = 5
                    Zeilenhoehe = 290 + 4
                    '-------------------------------------------------------
                    If .Modernisierungsempfehlung_Anzahl < 5 Then
                        Zeilenzahl = .Modernisierungsempfehlung_Anzahl
                    End If
                    '-------------------------------------------------------
                    For Anzahl_Zeilen = 1 To Zeilenzahl
                        '-------------------------------------------------------
                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Modernisierungsempfehlungen(Anzahl_Zeilen).Nummer.Wert, 0, "", 245, 416, 1882 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1882 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                        '-------------------------------------------------------
                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksMitte, .Modernisierungsempfehlungen(Anzahl_Zeilen).Bauteil_Anlagenteil.Wert, 422 + 20, 925, 1882 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1882 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                        '-------------------------------------------------------
                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksMitte, .Modernisierungsempfehlungen(Anzahl_Zeilen).Massnahmenbeschreibung.Wert, 931 + 20, 2585, 1882 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1882 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                        '-------------------------------------------------------
                        Select Case .Modernisierungsempfehlungen(Anzahl_Zeilen).Modernisierungskombination.Wert
                            Case "in Zusammenhang mit größerer Modernisierung"
                                Image_Auswahl_schreiben(Grafik_Seite, True, 2704, 2750, 1922 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1969 + (Zeilenhoehe * (Anzahl_Zeilen - 1)))
                                Image_Auswahl_schreiben(Grafik_Seite, False, 2951, 2998, 1922 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1969 + (Zeilenhoehe * (Anzahl_Zeilen - 1)))
                            Case "als Einzelmaßnahme"
                                Image_Auswahl_schreiben(Grafik_Seite, False, 2704, 2750, 1922 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1969 + (Zeilenhoehe * (Anzahl_Zeilen - 1)))
                                Image_Auswahl_schreiben(Grafik_Seite, True, 2951, 2998, 1922 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1969 + (Zeilenhoehe * (Anzahl_Zeilen - 1)))
                        End Select
                        '-------------------------------------------------------
                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksMitte, .Modernisierungsempfehlungen(Anzahl_Zeilen).Amortisation.Wert, 3086 + 20, 3412, 1882 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1882 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                        '-------------------------------------------------------
                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksMitte, .Modernisierungsempfehlungen(Anzahl_Zeilen).spezifische_Kosten.Wert, 3418 + 20, 4032, 1882 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1882 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                        '-------------------------------------------------------
                    Next
                    '--------------------------------------------------
                    If .Modernisierungsempfehlung_Anzahl > 5 Then
                        '-------------------------------------------------------
                        Image_Auswahl_schreiben(Grafik_Seite, True, 317, 359, 3381, 3423)
                        '-------------------------------------------------------
                    Else
                        '-------------------------------------------------------
                        Image_Auswahl_schreiben(Grafik_Seite, False, 317, 359, 3381, 3423)
                        '-------------------------------------------------------
                    End If
                    '--------------------------------------------------
                End If
                '--------------------------------------------------
                Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksMitte, .Gebaeudebezogene_Daten.Angaben_erhaeltlich.Wert, 1672 + 20, 4032, 3637, 3884, Font_Schriftgroesse_40, Font_Druckfarbe)
                '--------------------------------------------------
                Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksOben, .Gebaeudebezogene_Daten.Ergaenzdende_Erlaeuterungen.Wert, 245 + 20, 4032, 4165, 5569, Font_Schriftgroesse_40, Font_Druckfarbe)
                '--------------------------------------------------
            End With
            '--------------------------------------------------
            If Variable_Steuerung.Entwurf = True Then
                Image_Text_im_Winkel_schreiben(Grafik_Seite, -45, "ENTWURF", 0, 2100, 0, 2970, Font_Schriftgroesse_ENTWURF, Font_Druckfarbe_ENTWURF)
            End If
            '-------------------------------------------------------
            If Oberflaeche = True Then
                '-------------------------------------------------------
                Imagelist.Images.Add(Bitmap_skalieren(Bitmap_Seite, Skalierungsfaktor))
                Listview.Items.Add("Seite " & Seitennummer, Seitennummer - 1)
                Listview.Update()
                '-------------------------------------------------------
            End If
            '-------------------------------------------------------
            Variable_XML_Import.Bilddateien(Seitennummer) = Image_to_String(Bitmap_Seite)
            Variable_XML_Import.Bildanzahl += 1
            '-------------------------------------------------------
            Speicher_freigeben()
            '-------------------------------------------------------
#End Region
            '-------------------------------------------------------
#Region "Seite 5"
            '-------------------------------------------------------
            'Seite 5 -----------------------------------------------
            '-------------------------------------------------------
            Seitennummer += 1
            '-------------------------------------------------------
            Grafik_Seite.Clear(Color.White)
            '-------------------------------------------------------
            Hintergrundgrafik = Image.FromFile(Variable_Parameter.Arbeitsverzeichnis & "/Image/GEG2020_WB_Seite_5.png")
            Grafik_Seite.DrawImage(Hintergrundgrafik, 0, 0, 4200, 5940)
            '-------------------------------------------------------
            With Variable_XML_Import
                '-------------------------------------------------------
                Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, Datum_Energiegesetz, 2128, 2616, 422, 516, Font_Schriftgroesse_50, Font_Druckfarbe)
                '-------------------------------------------------------
            End With
            '--------------------------------------------------
            If Variable_Steuerung.Entwurf = True Then
                Image_Text_im_Winkel_schreiben(Grafik_Seite, -45, "ENTWURF", 0, 2100, 0, 2970, Font_Schriftgroesse_ENTWURF, Font_Druckfarbe_ENTWURF)
            End If
            '-------------------------------------------------------
            If Oberflaeche = True Then
                '-------------------------------------------------------
                Imagelist.Images.Add(Bitmap_skalieren(Bitmap_Seite, Skalierungsfaktor))
                Listview.Items.Add("Seite " & Seitennummer, Seitennummer - 1)
                Listview.Update()
                '-------------------------------------------------------
            End If
            '-------------------------------------------------------
            Variable_XML_Import.Bilddateien(Seitennummer) = Image_to_String(Bitmap_Seite)
            Variable_XML_Import.Bildanzahl += 1
            '-------------------------------------------------------
            Speicher_freigeben()
            '-------------------------------------------------------
#End Region
            '-------------------------------------------------------
#Region "Zusatzseite Verbrauchserfassung"
            '-------------------------------------------------------
            If Variable_Steuerung.Berechnungsart = "Energieverbrauchsausweis" Or Variable_XML_Import.Zusaetzliche_Verbrauchsberechnung = True Then
                '-------------------------------------------------------
                If Variable_XML_Import.Berechnungsverfahren = "Verbrauchsberechnung" Or Variable_XML_Import.Zusaetzliche_Verbrauchsberechnung = True Then
                    '-------------------------------------------------------
                    Wert_Schleife_Offset = 6
                    Wert_Schleife_Maximum = 324
                    Wert_Schleife_Zeilenzahl = 50
                    '-------------------------------------------------------
                    For Wert_Schleife = 0 To (Wert_Schleife_Maximum) Step Wert_Schleife_Zeilenzahl
                        '-------------------------------------------------------
                        With Variable_XML_Import
                            '-------------------------------------------------------
                            If .Verbrauchserfassung_Anzahl > (Wert_Schleife + Wert_Schleife_Offset) Then
                                '-------------------------------------------------------
                                Seitennummer += 1
                                '-------------------------------------------------------
                                Grafik_Seite.Clear(Color.White)
                                '-------------------------------------------------------
                                Hintergrundgrafik = Image.FromFile(Variable_Parameter.Arbeitsverzeichnis & "/Image/GEG2020_WB_Ueberlauf_Verbrauch.png")
                                Grafik_Seite.DrawImage(Hintergrundgrafik, 0, 0, 4200, 5940)
                                '-------------------------------------------------------
                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, Seitennummer, 0, "", 3893, 4040, 688, 840, Font_Seitennummer, Font_Druckfarbe)
                                '-------------------------------------------------------
                                Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, Datum_Energiegesetz, 2128, 2616, 422, 516, Font_Schriftgroesse_50, Font_Druckfarbe)
                                Image_Text_schreiben(Grafik_Seite, Ausrichtung.Mitte_kein_Abschneiden, Variable_Steuerung.Registriernummer, 3053, 3768, 722, 820, Font_Schriftgroesse_50, Font_Druckfarbe)
                                '-------------------------------------------------------
                                Zeilenhoehe = (5242 - 1366) / Wert_Schleife_Zeilenzahl
                                '-------------------------------------------------------
                                If .Verbrauchserfassung_Anzahl < (Wert_Schleife + Wert_Schleife_Zeilenzahl + Wert_Schleife_Offset) Then
                                    '-------------------------------------------------------
                                    Wert_Schleife_Zeilenzahl = .Verbrauchserfassung_Anzahl - Wert_Schleife_Offset - Wert_Schleife
                                    '-------------------------------------------------------
                                End If
                                '-------------------------------------------------------
                                For Anzahl_Zeilen = 1 To Wert_Schleife_Zeilenzahl
                                    '-------------------------------------------------------
                                    Image_Datum_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Verbrauchserfassung(Anzahl_Zeilen + Wert_Schleife + Wert_Schleife_Offset).Startdatum.Wert, 245, 572, 1366 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1366 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                    '--------------------------------------------------
                                    Image_Datum_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Verbrauchserfassung(Anzahl_Zeilen + Wert_Schleife + Wert_Schleife_Offset).Enddatum.Wert, 576, 902, 1366 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1366 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                    '--------------------------------------------------
                                    Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Verbrauchserfassung(Anzahl_Zeilen + Wert_Schleife + Wert_Schleife_Offset).Energietraeger_Verbrauch.Wert, 906, 2469, 1366 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1366 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                    '--------------------------------------------------
                                    Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Verbrauchserfassung(Anzahl_Zeilen + Wert_Schleife + Wert_Schleife_Offset).Primaerenergiefaktor.Wert, 1, "", 2473, 2709, 1366 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1366 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                    '--------------------------------------------------
                                    If .Verbrauchserfassung(Anzahl_Zeilen + Wert_Schleife + Wert_Schleife_Offset).Energieverbrauch.Wert <> 0 Then
                                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Verbrauchserfassung(Anzahl_Zeilen + Wert_Schleife + Wert_Schleife_Offset).Energieverbrauch.Wert, 0, "", 2713, 3069, 1366 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1366 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                    End If
                                    '--------------------------------------------------
                                    If .Verbrauchserfassung(Anzahl_Zeilen + Wert_Schleife + Wert_Schleife_Offset).Energieverbrauchsanteil_Warmwasser_zentral.Wert <> 0 Then
                                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Verbrauchserfassung(Anzahl_Zeilen + Wert_Schleife + Wert_Schleife_Offset).Energieverbrauchsanteil_Warmwasser_zentral.Wert, 0, "", 3073, 3429, 1366 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1366 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                    End If
                                    '--------------------------------------------------
                                    If .Verbrauchserfassung(Anzahl_Zeilen + Wert_Schleife + Wert_Schleife_Offset).Energieverbrauchsanteil_Heizung.Wert <> 0 Then
                                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Verbrauchserfassung(Anzahl_Zeilen + Wert_Schleife + Wert_Schleife_Offset).Energieverbrauchsanteil_Heizung.Wert, 0, "", 3433, 3789, 1366 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1366 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                    End If
                                    '--------------------------------------------------
                                    If .Verbrauchserfassung(Anzahl_Zeilen + Wert_Schleife + Wert_Schleife_Offset).Klimafaktor.Wert <> 0 Then
                                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Verbrauchserfassung(Anzahl_Zeilen + Wert_Schleife + Wert_Schleife_Offset).Klimafaktor.Wert, 2, "", 3793, 4032, 1366 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1366 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                    End If
                                    '-------------------------------------------------------
                                Next
                                '-------------------------------------------------------
                                For Anzahl_Zeilen = 1 To Wert_Schleife_Zeilenzahl 'Linien
                                    '-------------------------------------------------------
                                    Image_Linie(Grafik_Seite, 245, 4032, 1366 + (Zeilenhoehe * Anzahl_Zeilen), 1366 + (Zeilenhoehe * Anzahl_Zeilen))
                                    '-------------------------------------------------------
                                Next
                                '-------------------------------------------------------
                                If .Verbrauchserfassung_Anzahl > (Wert_Schleife + Wert_Schleife_Zeilenzahl + Wert_Schleife_Offset) Then 'Auswahl
                                    '-------------------------------------------------------
                                    Image_Auswahl_schreiben(Grafik_Seite, True, 300, 330, 5462, 5492)
                                    '-------------------------------------------------------
                                End If
                                '--------------------------------------------------
                                If Variable_Steuerung.Entwurf = True Then
                                    Image_Text_im_Winkel_schreiben(Grafik_Seite, -45, "ENTWURF", 0, 2100, 0, 2970, Font_Schriftgroesse_ENTWURF, Font_Druckfarbe_ENTWURF)
                                End If
                                '-------------------------------------------------------
                                If Oberflaeche = True Then
                                    '-------------------------------------------------------
                                    Imagelist.Images.Add(Bitmap_skalieren(Bitmap_Seite, Skalierungsfaktor))
                                    Listview.Items.Add("Verbrauchserfassung", Seitennummer - 1)
                                    Listview.Update()
                                    '-------------------------------------------------------
                                End If
                                '-------------------------------------------------------
                                Variable_XML_Import.Bilddateien(Seitennummer) = Image_to_String(Bitmap_Seite)
                                Variable_XML_Import.Bildanzahl += 1
                                '-------------------------------------------------------
                                Speicher_freigeben()
                                '-------------------------------------------------------
                            End If
                            '-------------------------------------------------------
                        End With
                        '-------------------------------------------------------
                    Next
                    '-------------------------------------------------------
                End If
                '-------------------------------------------------------
            End If
            '-------------------------------------------------------
#End Region
            '-------------------------------------------------------
#Region "Zusatzseite Modernisierungsempfehlungen"
            '-------------------------------------------------------
            Wert_Schleife_Offset = 5
            Wert_Schleife_Maximum = 30
            Wert_Schleife_Zeilenzahl = 12
            '-------------------------------------------------------
            For Wert_Schleife = 0 To (Wert_Schleife_Maximum) Step Wert_Schleife_Zeilenzahl
                '-------------------------------------------------------
                With Variable_XML_Import
                    '-------------------------------------------------------
                    If .Modernisierungsempfehlung_Anzahl > (Wert_Schleife + Wert_Schleife_Offset) Then
                        '-------------------------------------------------------
                        Seitennummer += 1
                        '-------------------------------------------------------
                        Grafik_Seite.Clear(Color.White)
                        '-------------------------------------------------------
                        Hintergrundgrafik = Image.FromFile(Variable_Parameter.Arbeitsverzeichnis & "/Image/GEG2020_WB_Ueberlauf_Modernisierung.png")
                        Grafik_Seite.DrawImage(Hintergrundgrafik, 0, 0, 4200, 5940)
                        '-------------------------------------------------------
                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, Seitennummer, 0, "", 3893, 4040, 688, 840, Font_Seitennummer, Font_Druckfarbe)
                        '-------------------------------------------------------
                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, Datum_Energiegesetz, 2128, 2616, 422, 516, Font_Schriftgroesse_50, Font_Druckfarbe)
                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.Mitte_kein_Abschneiden, Variable_Steuerung.Registriernummer, 3053, 3768, 722, 820, Font_Schriftgroesse_50, Font_Druckfarbe)
                        '-------------------------------------------------------
                        If .Energieausweis_Daten.Empfehlungen_moeglich.Wert = False Then
                            '-------------------------------------------------------
                            Image_Auswahl_schreiben(Grafik_Seite, True, 2833, 2880, 1143, 1176)
                            '-------------------------------------------------------
                        Else
                            '-------------------------------------------------------
                            Image_Auswahl_schreiben(Grafik_Seite, True, 2425, 2471, 1143, 1176)
                            '-------------------------------------------------------
                            Zeilenhoehe = 294
                            '-------------------------------------------------------
                            If .Modernisierungsempfehlung_Anzahl < (Wert_Schleife + Wert_Schleife_Zeilenzahl + Wert_Schleife_Offset) Then
                                '-------------------------------------------------------
                                Wert_Schleife_Zeilenzahl = .Modernisierungsempfehlung_Anzahl - Wert_Schleife_Offset - Wert_Schleife
                                '-------------------------------------------------------
                            End If
                            '-------------------------------------------------------
                            For Anzahl_Zeilen = 1 To Wert_Schleife_Zeilenzahl
                                '-------------------------------------------------------
                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Modernisierungsempfehlungen(Anzahl_Zeilen + Wert_Schleife + Wert_Schleife_Offset).Nummer.Wert, 0, "", 245, 417, 1882 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1882 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                '-------------------------------------------------------
                                Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksMitte, .Modernisierungsempfehlungen(Anzahl_Zeilen + Wert_Schleife + Wert_Schleife_Offset).Bauteil_Anlagenteil.Wert, 422 + 20, 925, 1882 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1882 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                '-------------------------------------------------------
                                Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksMitte, .Modernisierungsempfehlungen(Anzahl_Zeilen + Wert_Schleife + Wert_Schleife_Offset).Massnahmenbeschreibung.Wert, 931 + 20, 2585, 1882 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1882 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                '-------------------------------------------------------
                                Select Case .Modernisierungsempfehlungen(Anzahl_Zeilen + Wert_Schleife + Wert_Schleife_Offset).Modernisierungskombination.Wert
                                    Case "in Zusammenhang mit größerer Modernisierung"
                                        Image_Auswahl_schreiben(Grafik_Seite, True, 2703, 2750, 1922 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1969 + (Zeilenhoehe * (Anzahl_Zeilen - 1)))
                                    Case "als Einzelmaßnahme"
                                        Image_Auswahl_schreiben(Grafik_Seite, True, 2951, 2998, 1922 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1969 + (Zeilenhoehe * (Anzahl_Zeilen - 1)))
                                End Select
                                '-------------------------------------------------------
                                Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksMitte, .Modernisierungsempfehlungen(Anzahl_Zeilen + Wert_Schleife + Wert_Schleife_Offset).Amortisation.Wert, 3086 + 20, 3412, 1882 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1882 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                '-------------------------------------------------------
                                Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksMitte, .Modernisierungsempfehlungen(Anzahl_Zeilen + Wert_Schleife + Wert_Schleife_Offset).spezifische_Kosten.Wert, 3419 + 20, 4032, 1882 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1882 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                '-------------------------------------------------------
                            Next
                            '-------------------------------------------------------
                            If .Modernisierungsempfehlung_Anzahl > (Wert_Schleife + Wert_Schleife_Zeilenzahl + Wert_Schleife_Offset) Then 'Auswahl
                                '-------------------------------------------------------
                                Image_Auswahl_schreiben(Grafik_Seite, True, 317, 359, 5459, 5500)
                                '-------------------------------------------------------
                            End If
                            '-------------------------------------------------------
                        End If
                        '--------------------------------------------------
                        If Variable_Steuerung.Entwurf = True Then
                            Image_Text_im_Winkel_schreiben(Grafik_Seite, -45, "ENTWURF", 0, 2100, 0, 2970, Font_Schriftgroesse_ENTWURF, Font_Druckfarbe_ENTWURF)
                        End If
                        '-------------------------------------------------------
                        If Oberflaeche = True Then
                            '-------------------------------------------------------
                            Imagelist.Images.Add(Bitmap_skalieren(Bitmap_Seite, Skalierungsfaktor))
                            Listview.Items.Add("Modernisierungsempfehlungen", Seitennummer - 1)
                            Listview.Update()
                            '-------------------------------------------------------
                        End If
                        '-------------------------------------------------------
                        Variable_XML_Import.Bilddateien(Seitennummer) = Image_to_String(Bitmap_Seite)
                        Variable_XML_Import.Bildanzahl += 1
                        '-------------------------------------------------------
                        Speicher_freigeben()
                        '-------------------------------------------------------
                    End If
                    '-------------------------------------------------------
                End With
                '-------------------------------------------------------
            Next
            '-------------------------------------------------------
#End Region
            '-------------------------------------------------------
        Catch ex As Exception
            Fehlerfenster(ex)
        End Try
        '--------------------------------------------------
    End Sub
    '--------------------------------------------------
#End Region
    '--------------------------------------------------
End Class
