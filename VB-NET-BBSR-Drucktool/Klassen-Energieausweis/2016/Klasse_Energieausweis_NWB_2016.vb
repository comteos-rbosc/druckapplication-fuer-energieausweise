#Region "Imports"
Imports System.Drawing
Imports System.Windows.Forms
#End Region
'--------------------------------------------------
Public Class Klasse_Energieausweis_NWB_2016
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
    ''' Hier wird der gesamte Ablauf für die Erstellung der Energieausweises abgearbeitet - Nichtwohnbau - EnEV 2016
    ''' </summary>
    Sub Energieausweis_erzeugen(ByVal Oberflaeche As Boolean, ByVal PictureBox As PictureBox, ByVal Listview As ListView, ByVal Imagelist As ImageList)
        '--------------------------------------------------
        Try
            '-------------------------------------------------------
#Region "Variablen"
            '-------------------------------------------------------
            Dim Datum_Energiegesetz As String = "18.11.2013"
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
            Hintergrundgrafik = Image.FromFile(Variable_Parameter.Arbeitsverzeichnis & "/Image/ENEV2016_NWB_Seite_1.png")
            Grafik_Seite.DrawImage(Hintergrundgrafik, 0, 0, 4200, 5940)
            '-------------------------------------------------------
            With Variable_XML_Import
                '-------------------------------------------------------
                Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, Datum_Energiegesetz, 2458, 2919, 448, 545, Font_Schriftgroesse_50, Font_Druckfarbe)
                Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, DateAdd(DateInterval.Year, 10, Variable_Steuerung.Ausstellungsdatum), 549 + 20, 1000, 771, 821, Font_Schriftgroesse_50, Font_Druckfarbe)
                Image_Text_schreiben(Grafik_Seite, Ausrichtung.Mitte_kein_Abschneiden, Variable_Steuerung.Registriernummer, 3238, 3699, 646, 745, Font_Bold_Schriftgroesse_40_Narrow, Font_Druckfarbe)
                '-------------------------------------------------------
                Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksOben, .Energieausweis_Daten.Nichtwohngebaeude.Hauptnutzung_Gebaeudekategorie.Wert, 1290 + 20, 3130, 1103, 1311, Font_Schriftgroesse_50, Font_Druckfarbe)
                Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Gebaeudebezogene_Daten.Gebaeudeadresse_Strasse_Nr.Wert & ", " & .Gebaeudebezogene_Daten.Gebaeudeadresse_Postleitzahl.Wert & " " & .Gebaeudebezogene_Daten.Gebaeudeadresse_Ort.Wert, 1290 + 20, 3130, 1320, 1441, Font_Schriftgroesse_50, Font_Druckfarbe)
                Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Energieausweis_Daten.Gebaeudeteil.Wert, 1290 + 20, 3130, 1450, 1571, Font_Schriftgroesse_50, Font_Druckfarbe)
                Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Energieausweis_Daten.Baujahr_Gebaeude.Wert, 1290 + 20, 3130, 1580, 1701, Font_Schriftgroesse_50, Font_Druckfarbe)
                Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Energieausweis_Daten.Baujahr_Waermeerzeuger.Wert, 1290 + 20, 3130, 1710, 1830, Font_Schriftgroesse_50, Font_Druckfarbe)
                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Energieausweis_Daten.Nichtwohngebaeude.Nettogrundflaeche.Wert, 2, "m²", 1290 + 20, 3130, 1839, 1960, Font_Schriftgroesse_50, Font_Druckfarbe)
                Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksMitte, .Energieausweis_Daten.wesentliche_Energietraeger_Heizung.Wert, 1290 + 20, 3130, 1969, 2178, Font_Schriftgroesse_50, Font_Druckfarbe)
                Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Energieausweis_Daten.Erneuerbare_Art.Wert, 1416 + 20, 2490, 2187 + 15, 2308, Font_Schriftgroesse_50, Font_Druckfarbe)
                Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Energieausweis_Daten.Erneuerbare_Verwendung.Wert, 3119 + 20, 4030, 2187 + 15, 2308, Font_Schriftgroesse_50, Font_Druckfarbe)
                '--------------------------------------------------
                Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Lueftungsart_Fensterlueftung.Wert, 1313, 1352, 2365, 2404)
                Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Lueftungsart_Schachtlueftung.Wert, 1313, 1352, 2453, 2492)
                Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Lueftungsart_Anlage_m_WRG.Wert, 1926, 1965, 2365, 2404)
                Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Lueftungsart_Anlage_o_WRG.Wert, 1926, 1965, 2453, 2492)
                Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Anlage_zur_Kuehlung.Wert, 3471, 3510, 2365, 2404)
                '-------------------------------------------------------
                Select Case .Energieausweis_Daten.Ausstellungsanlass.Wert
                    Case "Neubau"
                        Image_Auswahl_schreiben(Grafik_Seite, True, 1313, 1352, 2583, 2622)
                        Image_Auswahl_schreiben(Grafik_Seite, False, 2124, 2164, 2583, 2622)
                        Image_Auswahl_schreiben(Grafik_Seite, False, 1313, 1352, 2671, 2711)
                        Image_Auswahl_schreiben(Grafik_Seite, False, 3080, 3120, 2671, 2711)
                        Image_Auswahl_schreiben(Grafik_Seite, False, 3080, 3120, 2583, 2622)
                    Case "Modernisierung-Erweiterung"
                        Image_Auswahl_schreiben(Grafik_Seite, False, 1313, 1352, 2583, 2622)
                        Image_Auswahl_schreiben(Grafik_Seite, True, 2124, 2164, 2583, 2622)
                        Image_Auswahl_schreiben(Grafik_Seite, False, 1313, 1352, 2671, 2711)
                        Image_Auswahl_schreiben(Grafik_Seite, False, 3080, 3120, 2671, 2711)
                        Image_Auswahl_schreiben(Grafik_Seite, False, 3080, 3120, 2583, 2622)
                    Case "Vermietung-Verkauf"
                        Image_Auswahl_schreiben(Grafik_Seite, False, 1313, 1352, 2583, 2622)
                        Image_Auswahl_schreiben(Grafik_Seite, False, 2124, 2164, 2583, 2622)
                        Image_Auswahl_schreiben(Grafik_Seite, True, 1313, 1352, 2671, 2711)
                        Image_Auswahl_schreiben(Grafik_Seite, False, 3080, 3120, 2671, 2711)
                        Image_Auswahl_schreiben(Grafik_Seite, False, 3080, 3120, 2583, 2622)
                    Case "Sonstiges"
                        Image_Auswahl_schreiben(Grafik_Seite, False, 1313, 1352, 2583, 2622)
                        Image_Auswahl_schreiben(Grafik_Seite, False, 2124, 2164, 2583, 2622)
                        Image_Auswahl_schreiben(Grafik_Seite, False, 1313, 1352, 2671, 2711)
                        Image_Auswahl_schreiben(Grafik_Seite, True, 3080, 3120, 2671, 2711)
                        Image_Auswahl_schreiben(Grafik_Seite, False, 3080, 3120, 2583, 2622)
                    Case "Aushangpflicht"
                        Image_Auswahl_schreiben(Grafik_Seite, False, 1313, 1352, 2583, 2622)
                        Image_Auswahl_schreiben(Grafik_Seite, False, 2124, 2164, 2583, 2622)
                        Image_Auswahl_schreiben(Grafik_Seite, False, 1313, 1352, 2671, 2711)
                        Image_Auswahl_schreiben(Grafik_Seite, False, 3080, 3120, 2671, 2711)
                        Image_Auswahl_schreiben(Grafik_Seite, True, 3080, 3120, 2583, 2622)
                End Select
                '-------------------------------------------------------
                Select Case Variable_Steuerung.Berechnungsart
                    Case "Energiebedarfsausweis"
                        Image_Auswahl_schreiben(Grafik_Seite, True, 262, 301, 3402, 3441)
                        Image_Auswahl_schreiben(Grafik_Seite, False, 262, 301, 3864, 3903)
                    Case "Energieverbrauchsausweis"
                        Image_Auswahl_schreiben(Grafik_Seite, False, 262, 301, 3402, 3441)
                        Image_Auswahl_schreiben(Grafik_Seite, True, 262, 301, 3864, 3903)
                End Select
                '-------------------------------------------------------
                Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Datenerhebung_Eigentuemer.Wert, 2089, 2128, 4158, 4198)
                Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Datenerhebung_Aussteller.Wert, 2941, 2981, 4158, 4198)
                '-------------------------------------------------------
                Image_Auswahl_schreiben(Grafik_Seite, .Gebaeudebezogene_Daten.Zusatzinfos_beigefuegt.Wert, 262, 301, 4283, 4323)
                '-------------------------------------------------------
                Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, FormatDateTime(Variable_Steuerung.Ausstellungsdatum, DateFormat.ShortDate).ToString, 2240, 2750, 5284, 5356, Font_Schriftgroesse_50, Font_Druckfarbe)
                '-------------------------------------------------------
                Image_Grafik_Projekt_schreiben(Grafik_Seite, Variable_Steuerung.Bild_Projekt, 3138, 1103, 4031 - 3138, 2178 - 1103)
                '-------------------------------------------------------
                Image_Grafik_Aussteller_schreiben(Grafik_Seite, Variable_Steuerung.Bild_Aussteller, .Gebaeudebezogene_Daten.Aussteller_Bezeichnung.Wert, .Gebaeudebezogene_Daten.Ausstellervorname.Wert, .Gebaeudebezogene_Daten.Ausstellername.Wert, .Gebaeudebezogene_Daten.Aussteller_Strasse_Nr.Wert, .Gebaeudebezogene_Daten.Aussteller_PLZ.Wert, .Gebaeudebezogene_Daten.Aussteller_Ort.Wert, 290, 5075, 1730, 300)
                '-------------------------------------------------------
                Image_Grafik_Unterschrift_schreiben(Grafik_Seite, Variable_Steuerung.Bild_Unterschrift, 3000, 5050, 1000, 300)
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
            Hintergrundgrafik = Image.FromFile(Variable_Parameter.Arbeitsverzeichnis & "/Image/ENEV2016_NWB_Seite_2.png")
            Grafik_Seite.DrawImage(Hintergrundgrafik, 0, 0, 4200, 5940)
            '-------------------------------------------------------
            With Variable_XML_Import
                Select Case Variable_Steuerung.Berechnungsart
                    Case "Energiebedarfsausweis"
                        '-------------------------------------------------------
                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, Datum_Energiegesetz, 2459, 2919, 444, 543, Font_Schriftgroesse_50, Font_Druckfarbe)
                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.Mitte_kein_Abschneiden, Variable_Steuerung.Registriernummer, 3239, 3699, 643, 741, Font_Bold_Schriftgroesse_40_Narrow, Font_Druckfarbe)
                        '-------------------------------------------------------
                        Select Case .Berechnungsverfahren
                            Case "Bedarfsberechnung-18599"
                                '-------------------------------------------------------
                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Gebaeudebezogene_Daten.Treibhausgasemissionen.Wert, 1, "", 3469, 3676, 1118, 1243, Font_Schriftgroesse_50, Font_Druckfarbe)
                                '-------------------------------------------------------
                                Select Case .Energieausweis_Daten.Ausstellungsanlass.Wert
                                    Case "Neubau", "Modernisierung-Erweiterung", "Aushangpflicht"
                                        '-------------------------------------------------------
                                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Nichtwohngebaeude.Bedarfswerte_NWG.Primaerenergiebedarf_NGF.Wert, 1, "", 489, 717, 2271, 2351, Font_Schriftgroesse_40, Font_Druckfarbe)
                                        '-------------------------------------------------------
                                        Select Case Variable_Steuerung.Neubau
                                            Case 0 'Modernisierung
                                                '-------------------------------------------------------
                                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Nichtwohngebaeude.Bedarfswerte_NWG.Primaerenergiebedarf_Hoechstwert_Bestand.Wert, 1, "", 1515, 1742, 2271, 2351, Font_Schriftgroesse_40, Font_Druckfarbe)
                                                '-------------------------------------------------------
                                            Case 1 'Neubau
                                                '-------------------------------------------------------
                                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Nichtwohngebaeude.Bedarfswerte_NWG.Primaerenergiebedarf_Hoechstwert_Neubau.Wert, 1, "", 1515, 1742, 2271, 2351, Font_Schriftgroesse_40, Font_Druckfarbe)
                                                '-------------------------------------------------------
                                        End Select
                                        '-------------------------------------------------------
                                        Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Nichtwohngebaeude.Bedarfswerte_NWG.mittlere_Waermedurchgangskoeffizienten.Wert, 1534, 1565, 2395, 2425)
                                        Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Nichtwohngebaeude.Bedarfswerte_NWG.Sommerlicher_Waermeschutz.Wert, 1534, 1565, 2483, 2513)
                                        '-------------------------------------------------------
                                    Case "Vermietung-Verkauf", "Sonstiges"
                                        '-------------------------------------------------------
                                End Select
                                '-------------------------------------------------------
                                Select Case .Energieausweis_Daten.Nichtwohngebaeude.Bedarfswerte_NWG.Ein_Zonen_Modell.Wert
                                    Case False 'Mehr-Zonen-Modell
                                        Image_Auswahl_schreiben(Grafik_Seite, True, 2206, 2237, 2217, 2247)
                                        Image_Auswahl_schreiben(Grafik_Seite, False, 2206, 2237, 2306, 2337)
                                    Case True 'Ein-Zonen-Modell
                                        Image_Auswahl_schreiben(Grafik_Seite, False, 2206, 2237, 2217, 2247)
                                        Image_Auswahl_schreiben(Grafik_Seite, True, 2206, 2237, 2306, 2337)
                                End Select
                                '-------------------------------------------------------
                                Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Nichtwohngebaeude.Bedarfswerte_NWG.Vereinfachte_Datenaufnahme.Wert, 2206, 2237, 2395, 2425) 'Vereinfachung nach §50 Absatz 4
                                '-------------------------------------------------------
                                Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Nichtwohngebaeude.Bedarfswerte_NWG.Vereinfachungen_18599_1_D.Wert, 2206, 2237, 2483, 2513) 'Vereinfachung nach §21 Absatz 2 Satz 2
                                '-------------------------------------------------------
                                Image_Pfeil_NWB(Grafik_Seite, Lage.Oben, .Energieausweis_Daten.Nichtwohngebaeude.Bedarfswerte_NWG.Primaerenergiebedarf_NGF.Wert, 1, .Energieausweis_Daten.Nichtwohngebaeude.Bedarfswerte_NWG.Primaerenergiebedarf_Hoechstwert_Bestand.Wert, 3, "Primärenergiebedarf dieses Gebäudes", "", 651, 3671, 1528, 1882)
                                '-------------------------------------------------------
                                'Tabelle Energiebedarf ---------------------------------
                                '-------------------------------------------------------
                                .Energiebedarf_Anzahl = Anzahl_Energiebedarf()
                                '-------------------------------------------------------
                                Zeilenzahl = 4
                                Zeilenhoehe = 3088 - 2995
                                '-------------------------------------------------------
                                If .Energiebedarf_Anzahl < 4 Then
                                    Zeilenzahl = .Energiebedarf_Anzahl
                                End If
                                '-------------------------------------------------------
                                For Anzahl_Zeilen = 1 To Zeilenzahl
                                    '-------------------------------------------------------
                                    Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Energietraeger(Anzahl_Zeilen).Energietraegerbezeichnung.Wert, 241, 970, 2995 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 2995 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                    '--------------------------------------------------
                                    If .Energietraeger(Anzahl_Zeilen).Endenergiebedarf_Heizung_spezifisch.Wert <> 0 Then
                                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energietraeger(Anzahl_Zeilen).Endenergiebedarf_Heizung_spezifisch.Wert, 1, "", 978, 1469, 2995 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 2995 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                    End If
                                    '--------------------------------------------------
                                    If .Energietraeger(Anzahl_Zeilen).Endenergiebedarf_Trinkwarmwasser_spezifisch.Wert <> 0 Then
                                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energietraeger(Anzahl_Zeilen).Endenergiebedarf_Trinkwarmwasser_spezifisch.Wert, 1, "", 1477, 1967, 2995 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 2995 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                    End If
                                    '--------------------------------------------------
                                    If .Energietraeger(Anzahl_Zeilen).Endenergiebedarf_Beleuchtung_spezifisch.Wert <> 0 Then
                                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energietraeger(Anzahl_Zeilen).Endenergiebedarf_Beleuchtung_spezifisch.Wert, 1, "", 1975, 2540, 2995 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 2995 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                    End If
                                    '--------------------------------------------------
                                    If .Energietraeger(Anzahl_Zeilen).Endenergiebedarf_Lueftung_spezifisch.Wert <> 0 Then
                                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energietraeger(Anzahl_Zeilen).Endenergiebedarf_Lueftung_spezifisch.Wert, 1, "", 2548, 2965, 2995 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 2995 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                    End If
                                    '--------------------------------------------------
                                    If .Energietraeger(Anzahl_Zeilen).Endenergiebedarf_Kuehlung_Befeuchtung_spezifisch.Wert <> 0 Then
                                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energietraeger(Anzahl_Zeilen).Endenergiebedarf_Kuehlung_Befeuchtung_spezifisch.Wert, 1, "", 2973, 3517, 2995 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 2995 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                    End If
                                    '--------------------------------------------------
                                    If .Energietraeger(Anzahl_Zeilen).Endenergiebedarf_Energietraeger_Gesamtgebaeude_spezifisch.Wert <> 0 Then
                                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energietraeger(Anzahl_Zeilen).Endenergiebedarf_Energietraeger_Gesamtgebaeude_spezifisch.Wert, 1, "", 3525, 4032, 2995 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 2995 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                    End If
                                    '--------------------------------------------------
                                Next
                                '-------------------------------------------------------
                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Nichtwohngebaeude.Bedarfswerte_NWG.Endenergiebedarf_Waerme_NGF.Wert, 1, "", 2967, 3627, 3374, 3506, Font_Schriftgroesse_60, Font_Druckfarbe)
                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Nichtwohngebaeude.Bedarfswerte_NWG.Endenergiebedarf_Strom_NGF.Wert, 1, "", 2967, 3627, 3514, 3653, Font_Schriftgroesse_60, Font_Druckfarbe)
                                '-------------------------------------------------------
                                'Tabelle Zonen -----------------------------------------
                                '-------------------------------------------------------
                                .Zonen_Anzahl = Anzahl_Zonen()
                                '-------------------------------------------------------
                                Zeilenzahl = 7
                                Zeilenhoehe = 4087 - 3996
                                '-------------------------------------------------------
                                If .Zonen_Anzahl < 7 Then
                                    Zeilenzahl = .Zonen_Anzahl
                                End If
                                '-------------------------------------------------------
                                For Anzahl_Zeilen = 1 To Zeilenzahl
                                    '-------------------------------------------------------
                                    Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Energieausweis_Daten.Nichtwohngebaeude.Zone(Anzahl_Zeilen).Zonenbezeichnung.Wert, 2288, 3314, 3996 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 3996 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                    '--------------------------------------------------
                                    If .Energieausweis_Daten.Nichtwohngebaeude.Zone(Anzahl_Zeilen).Nettogrundflaeche_Zone.Wert <> 0 Then
                                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Nichtwohngebaeude.Zone(Anzahl_Zeilen).Nettogrundflaeche_Zone.Wert, 1, "", 3324, 3701, 3996 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 3996 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                    End If
                                    '--------------------------------------------------
                                    If .Energieausweis_Daten.Nichtwohngebaeude.Zone(Anzahl_Zeilen).Nettogrundflaeche_Zone.Wert <> 0 Then
                                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, (.Energieausweis_Daten.Nichtwohngebaeude.Zone(Anzahl_Zeilen).Nettogrundflaeche_Zone.Wert / .Energieausweis_Daten.Nichtwohngebaeude.Nettogrundflaeche.Wert * 100), 2, "", 3710, 4030, 3996 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 3996 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                    End If
                                    '--------------------------------------------------
                                Next
                                '-------------------------------------------------------
                                If .Zonen_Anzahl > 7 Then
                                    '-------------------------------------------------------
                                    Image_Auswahl_schreiben(Grafik_Seite, True, 2164, 2210, 4650, 4697)
                                    '-------------------------------------------------------
                                End If
                                '-------------------------------------------------------
                                'Erneuerbare Energien ----------------------------------
                                '-------------------------------------------------------
                                Select Case .Energieausweis_Daten.Ausstellungsanlass.Wert
                                    Case "Neubau"
                                        '-------------------------------------------------------
                                        If .Energieausweis_Daten.Nichtwohngebaeude.Bedarfswerte_NWG.Art_der_Nutzung_erneuerbaren_Energie_1.Wert <> "" Then
                                            Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Energieausweis_Daten.Nichtwohngebaeude.Bedarfswerte_NWG.Art_der_Nutzung_erneuerbaren_Energie_1.Wert, 419, 1034, 4117, 4214, Font_Schriftgroesse_40, Font_Druckfarbe)
                                            Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Nichtwohngebaeude.Bedarfswerte_NWG.Deckungsanteil_1.Wert, 1, "", 1653, 1886, 4117, 4214, Font_Schriftgroesse_40, Font_Druckfarbe)
                                        End If
                                        If .Energieausweis_Daten.Nichtwohngebaeude.Bedarfswerte_NWG.Art_der_Nutzung_erneuerbaren_Energie_2.Wert <> "" Then
                                            Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Energieausweis_Daten.Nichtwohngebaeude.Bedarfswerte_NWG.Art_der_Nutzung_erneuerbaren_Energie_2.Wert, 419, 1034, 4223, 4315, Font_Schriftgroesse_40, Font_Druckfarbe)
                                            Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Nichtwohngebaeude.Bedarfswerte_NWG.Deckungsanteil_2.Wert, 1, "", 1653, 1886, 4223, 4315, Font_Schriftgroesse_40, Font_Druckfarbe)
                                        End If
                                        If .Energieausweis_Daten.Nichtwohngebaeude.Bedarfswerte_NWG.Art_der_Nutzung_erneuerbaren_Energie_3.Wert <> "" Then
                                            Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Energieausweis_Daten.Nichtwohngebaeude.Bedarfswerte_NWG.Art_der_Nutzung_erneuerbaren_Energie_3.Wert, 419, 1034, 4325, 4410, Font_Schriftgroesse_40, Font_Druckfarbe)
                                            Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Nichtwohngebaeude.Bedarfswerte_NWG.Deckungsanteil_3.Wert, 1, "", 1653, 1886, 4325, 4410, Font_Schriftgroesse_40, Font_Druckfarbe)
                                        End If
                                        '-------------------------------------------------------
                                        If .Energieausweis_Daten.Nichtwohngebaeude.Bedarfswerte_NWG.verschaerft_nach_EEWaermeG_7_1_2_eingehalten.Wert = True Then
                                            Image_Auswahl_schreiben(Grafik_Seite, True, 255, 287, 4825, 4857)
                                            If .Energieausweis_Daten.Nichtwohngebaeude.Bedarfswerte_NWG.Primaerenergiebedarf_Hoechstwert_verschaerft.Wert > 0 Then
                                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Nichtwohngebaeude.Bedarfswerte_NWG.Primaerenergiebedarf_Hoechstwert_verschaerft.Wert, 1, "", 1344, 1652, 5032, 5124, Font_Schriftgroesse_40, Font_Druckfarbe)
                                            End If
                                        Else
                                            Image_Auswahl_schreiben(Grafik_Seite, False, 255, 287, 4825, 4857)
                                        End If
                                        If .Energieausweis_Daten.Nichtwohngebaeude.Bedarfswerte_NWG.verschaerft_nach_EEWaermeG_8.Wert > 0 Then
                                            Image_Auswahl_schreiben(Grafik_Seite, True, 255, 287, 5191, 5223)
                                            Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Nichtwohngebaeude.Bedarfswerte_NWG.verschaerft_nach_EEWaermeG_8.Wert, 0, "", 1455, 1645, 5164, 5242, Font_Schriftgroesse_40, Font_Druckfarbe)
                                            If .Energieausweis_Daten.Nichtwohngebaeude.Bedarfswerte_NWG.Primaerenergiebedarf_Hoechstwert_verschaerft.Wert > 0 Then
                                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Nichtwohngebaeude.Bedarfswerte_NWG.Primaerenergiebedarf_Hoechstwert_verschaerft.Wert, 1, "", 1344, 1652, 5397, 5489, Font_Schriftgroesse_40, Font_Druckfarbe)
                                            End If
                                        Else
                                            Image_Auswahl_schreiben(Grafik_Seite, False, 255, 287, 5191, 5223)
                                        End If
                                        '-------------------------------------------------------
                                    Case "Modernisierung-Erweiterung", "Aushangpflicht", "Vermietung-Verkauf", "Sonstiges"
                                        '-------------------------------------------------------
                                End Select
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
            '--------------------------------------------------
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
            Hintergrundgrafik = Image.FromFile(Variable_Parameter.Arbeitsverzeichnis & "/Image/ENEV2016_NWB_Seite_3.png")
            Grafik_Seite.DrawImage(Hintergrundgrafik, 0, 0, 4200, 5940)
            '-------------------------------------------------------
            With Variable_XML_Import
                '-------------------------------------------------------
                If Variable_Steuerung.Berechnungsart = "Energieverbrauchsausweis" Or Variable_XML_Import.Zusaetzliche_Verbrauchsberechnung = True Then
                    '-------------------------------------------------------
                    Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, Datum_Energiegesetz, 2458, 2920, 447, 546, Font_Schriftgroesse_50, Font_Druckfarbe)
                    Image_Text_schreiben(Grafik_Seite, Ausrichtung.Mitte_kein_Abschneiden, Variable_Steuerung.Registriernummer, 3240, 3699, 647, 744, Font_Bold_Schriftgroesse_40_Narrow, Font_Druckfarbe)
                    '-------------------------------------------------------
                    If .Berechnungsverfahren = "Verbrauchsberechnung" Or Variable_XML_Import.Zusaetzliche_Verbrauchsberechnung = True Then
                        '-------------------------------------------------------
                        Image_Pfeil_NWB(Grafik_Seite, Lage.Oben, .Energieausweis_Daten.Nichtwohngebaeude.Verbrauchswerte_NWG.Endenergieverbrauch_Waerme.Wert, 1, .Energieausweis_Daten.Nichtwohngebaeude.Verbrauchswerte_NWG.Endenergieverbrauch_Waerme_Vergleichswert.Wert, 2, "Endenergieverbrauch Wärme", "[Pflichtangabe in Immobilienanzeigen]", 657, 3676, 1520, 1856)
                        '-------------------------------------------------------
                        Image_Pfeil_NWB(Grafik_Seite, Lage.Oben, .Energieausweis_Daten.Nichtwohngebaeude.Verbrauchswerte_NWG.Endenergieverbrauch_Strom.Wert, 1, .Energieausweis_Daten.Nichtwohngebaeude.Verbrauchswerte_NWG.Endenergieverbrauch_Strom_Vergleichswert.Wert, 2, "Endenergieverbrauch Strom", "[Pflichtangabe in Immobilienanzeigen]", 633, 3655, 2682, 3019)
                        '-------------------------------------------------------
                        Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Nichtwohngebaeude.Verbrauchswerte_NWG.Warmwasser_enthalten.Wert, 274, 312, 2160, 2198)
                        '-------------------------------------------------------
                        Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Nichtwohngebaeude.Verbrauchswerte_NWG.Strom_Daten.Stromverbrauch_enthaelt_Zusatzheizung.Wert, 275, 311, 3416, 3454)
                        Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Nichtwohngebaeude.Verbrauchswerte_NWG.Strom_Daten.Stromverbrauch_enthaelt_Warmwasser.Wert, 984, 1020, 3416, 3454)
                        Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Nichtwohngebaeude.Verbrauchswerte_NWG.Strom_Daten.Stromverbrauch_enthaelt_Lueftung.Wert, 1644, 1682, 3416, 3454)
                        Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Nichtwohngebaeude.Verbrauchswerte_NWG.Strom_Daten.Stromverbrauch_enthaelt_Beleuchtung.Wert, 2126, 2163, 3416, 3454)
                        Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Nichtwohngebaeude.Verbrauchswerte_NWG.Strom_Daten.Stromverbrauch_enthaelt_Kuehlung.Wert, 3110, 3147, 3416, 3454)
                        Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Nichtwohngebaeude.Verbrauchswerte_NWG.Strom_Daten.Stromverbrauch_enthaelt_Sonstiges.Wert, 3580, 3618, 3416, 3454)
                        '-------------------------------------------------------
                        'Tabelle Verbrauchserfassung ---------------------------
                        '-------------------------------------------------------
                        Zusammenstellung_NWB_Verbrauchserfassung()
                        '-------------------------------------------------------
                        Zeilenzahl = 5
                        Zeilenhoehe = 4116 - 4004
                        '-------------------------------------------------------
                        If .Verbrauchserfassung_Anzahl < 5 Then
                            Zeilenzahl = .Verbrauchserfassung_Anzahl
                        End If
                        '-------------------------------------------------------
                        For Anzahl_Zeilen = 1 To Zeilenzahl
                            '-------------------------------------------------------
                            Image_Datum_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Verbrauchserfassung(Anzahl_Zeilen).Startdatum.Wert, 242, 591, 4004 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 4004 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                            '-------------------------------------------------------
                            Image_Datum_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Verbrauchserfassung(Anzahl_Zeilen).Enddatum.Wert, 599, 986, 4004 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 4004 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                            '-------------------------------------------------------
                            Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Verbrauchserfassung(Anzahl_Zeilen).Energietraeger_Verbrauch.Wert, 995, 1530, 4004 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 4004 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40_Narrow, Font_Druckfarbe)
                            '-------------------------------------------------------
                            If .Verbrauchserfassung(Anzahl_Zeilen).Primaerenergiefaktor.Wert <> 0 Then
                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Verbrauchserfassung(Anzahl_Zeilen).Primaerenergiefaktor.Wert, 1, "", 1538, 1848, 4004 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 4004 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                            End If
                            '-------------------------------------------------------
                            If .Verbrauchserfassung(Anzahl_Zeilen).Energieverbrauch.Wert <> 0 Then
                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Verbrauchserfassung(Anzahl_Zeilen).Energieverbrauch.Wert, 0, "", 1857, 2486, 4004 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 4004 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                            End If
                            '-------------------------------------------------------
                            If .Verbrauchserfassung(Anzahl_Zeilen).Energieverbrauchsanteil_Warmwasser_zentral.Wert <> 0 Then
                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Verbrauchserfassung(Anzahl_Zeilen).Energieverbrauchsanteil_Warmwasser_zentral.Wert, 0, "", 2495, 2886, 4004 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 4004 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                            End If
                            '-------------------------------------------------------
                            If .Verbrauchserfassung(Anzahl_Zeilen).Energieverbrauch.Wert - .Verbrauchserfassung(Anzahl_Zeilen).Energieverbrauchsanteil_Warmwasser_zentral.Wert <> 0 Then
                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Verbrauchserfassung(Anzahl_Zeilen).Energieverbrauch.Wert - .Verbrauchserfassung(Anzahl_Zeilen).Energieverbrauchsanteil_Warmwasser_zentral.Wert, 0, "", 2895, 3247, 4004 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 4004 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                            End If
                            '-------------------------------------------------------
                            If .Verbrauchserfassung(Anzahl_Zeilen).Klimafaktor.Wert <> 0 Then
                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Verbrauchserfassung(Anzahl_Zeilen).Klimafaktor.Wert, 2, "", 3256, 3583, 4004 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 4004 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                            End If
                            '-------------------------------------------------------
                            If .Verbrauchserfassung(Anzahl_Zeilen).Energieverbrauch_Strom.Wert <> 0 Then
                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Verbrauchserfassung(Anzahl_Zeilen).Energieverbrauch_Strom.Wert, 0, "", 3591, 4031, 4004 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 4004 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                            End If
                            '-------------------------------------------------------
                        Next
                        '-------------------------------------------------------
                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Nichtwohngebaeude.Verbrauchswerte_NWG.Primaerenergieverbrauch.Wert, 1, "", 3251, 3663, 4565, 4672, Font_Schriftgroesse_60, Font_Druckfarbe)
                        '-------------------------------------------------------
                        'Tabelle Gebäudenutzung --------------------------------
                        '-------------------------------------------------------
                        .Gebaeudenutzung_Anzahl = Anzahl_Gebaeudenutzung()
                        '-------------------------------------------------------
                        Zeilenzahl = 3
                        Zeilenhoehe = 5299 - 5191
                        '-------------------------------------------------------
                        If .Gebaeudenutzung_Anzahl < 3 Then
                            Zeilenzahl = .Gebaeudenutzung_Anzahl
                        End If
                        '-------------------------------------------------------
                        For Anzahl_Zeilen = 1 To Zeilenzahl
                            '-------------------------------------------------------
                            Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Nutzung_Gebaeudekategorie(Anzahl_Zeilen).Gebaeudekategorie.Wert, 245, 902, 5191 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 5191 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40_Narrow, Font_Druckfarbe)
                            '--------------------------------------------------
                            If .Nutzung_Gebaeudekategorie(Anzahl_Zeilen).Flaechenanteil_Nutzung.Wert <> 0 Then
                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligRechts, .Nutzung_Gebaeudekategorie(Anzahl_Zeilen).Flaechenanteil_Nutzung.Wert, 1, "", 911, 1285, 5191 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 5191 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                            End If
                            '--------------------------------------------------
                            If .Nutzung_Gebaeudekategorie(Anzahl_Zeilen).Vergleichswert_Waerme.Wert <> 0 Then
                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Nutzung_Gebaeudekategorie(Anzahl_Zeilen).Vergleichswert_Waerme.Wert, 1, "", 1387, 1838, 5191 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 5191 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                            End If
                            '--------------------------------------------------
                            If .Nutzung_Gebaeudekategorie(Anzahl_Zeilen).Vergleichswert_Strom.Wert <> 0 Then
                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Nutzung_Gebaeudekategorie(Anzahl_Zeilen).Vergleichswert_Strom.Wert, 1, "", 1847, 2285, 5191 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 5191 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                            End If
                            '--------------------------------------------------
                        Next
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
            '--------------------------------------------------
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
            Hintergrundgrafik = Image.FromFile(Variable_Parameter.Arbeitsverzeichnis & "/Image/ENEV2016_NWB_Seite_4.png")
            Grafik_Seite.DrawImage(Hintergrundgrafik, 0, 0, 4200, 5940)
            '-------------------------------------------------------
            With Variable_XML_Import
                '-------------------------------------------------------
                Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, Datum_Energiegesetz, 2458, 2920, 449, 546, Font_Schriftgroesse_50, Font_Druckfarbe)
                Image_Text_schreiben(Grafik_Seite, Ausrichtung.Mitte_kein_Abschneiden, Variable_Steuerung.Registriernummer, 3237, 3699, 647, 745, Font_Bold_Schriftgroesse_40_Narrow, Font_Druckfarbe)
                '-------------------------------------------------------
                .Modernisierungsempfehlung_Anzahl = Anzahl_Modernisierungsempfehlung()
                '--------------------------------------------------
                If .Energieausweis_Daten.Empfehlungen_moeglich.Wert = False Then
                    '--------------------------------------------------
                    Image_Auswahl_schreiben(Grafik_Seite, False, 2828, 2882, 1195, 1249)
                    Image_Auswahl_schreiben(Grafik_Seite, True, 3446, 3499, 1195, 1249)
                    '--------------------------------------------------
                Else
                    '--------------------------------------------------
                    Image_Auswahl_schreiben(Grafik_Seite, True, 2828, 2882, 1195, 1249)
                    Image_Auswahl_schreiben(Grafik_Seite, False, 3446, 3499, 1195, 1249)
                    '--------------------------------------------------
                    Zeilenzahl = 10
                    Zeilenhoehe = 156
                    '-------------------------------------------------------
                    If .Modernisierungsempfehlung_Anzahl < 10 Then
                        Zeilenzahl = .Modernisierungsempfehlung_Anzahl
                    End If
                    '-------------------------------------------------------
                    For Anzahl_Zeilen = 1 To Zeilenzahl
                        '-------------------------------------------------------
                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Modernisierungsempfehlungen(Anzahl_Zeilen).Nummer.Wert, 0, "", 251, 428, 2035 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 2035 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                        '-------------------------------------------------------
                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksMitte, .Modernisierungsempfehlungen(Anzahl_Zeilen).Bauteil_Anlagenteil.Wert, 437 + 20, 1098, 2035 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 2035 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                        '-------------------------------------------------------
                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksMitte, .Modernisierungsempfehlungen(Anzahl_Zeilen).Massnahmenbeschreibung.Wert, 1108 + 20, 2380, 2035 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 2035 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                        '-------------------------------------------------------
                        Select Case .Modernisierungsempfehlungen(Anzahl_Zeilen).Modernisierungskombination.Wert
                            Case "in Zusammenhang mit größerer Modernisierung"
                                Image_Auswahl_schreiben(Grafik_Seite, True, 2641, 2695, 2083 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 2136 + (Zeilenhoehe * (Anzahl_Zeilen - 1)))
                                Image_Auswahl_schreiben(Grafik_Seite, False, 3090, 3143, 2083 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 2136 + (Zeilenhoehe * (Anzahl_Zeilen - 1)))
                            Case "als Einzelmaßnahme"
                                Image_Auswahl_schreiben(Grafik_Seite, False, 2641, 2695, 2083 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 2136 + (Zeilenhoehe * (Anzahl_Zeilen - 1)))
                                Image_Auswahl_schreiben(Grafik_Seite, True, 3090, 3143, 2083 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 2136 + (Zeilenhoehe * (Anzahl_Zeilen - 1)))
                        End Select
                        '-------------------------------------------------------
                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksMitte, .Modernisierungsempfehlungen(Anzahl_Zeilen).Amortisation.Wert, 3284 + 20, 3645, 2035 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 2035 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                        '-------------------------------------------------------
                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksMitte, .Modernisierungsempfehlungen(Anzahl_Zeilen).spezifische_Kosten.Wert, 3654 + 20, 4040, 2035 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 2035 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                        '-------------------------------------------------------
                    Next
                    '--------------------------------------------------
                    If .Modernisierungsempfehlung_Anzahl > 10 Then
                        '-------------------------------------------------------
                        Image_Auswahl_schreiben(Grafik_Seite, True, 312, 366, 3642, 3696)
                        '-------------------------------------------------------
                    Else
                        '-------------------------------------------------------
                        Image_Auswahl_schreiben(Grafik_Seite, False, 312, 366, 3642, 3696)
                        '-------------------------------------------------------
                    End If
                    '--------------------------------------------------
                End If
                '--------------------------------------------------
                Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksMitte, .Gebaeudebezogene_Daten.Angaben_erhaeltlich.Wert, 1840 + 20, 4040, 3989, 4293, Font_Schriftgroesse_40, Font_Druckfarbe)
                '--------------------------------------------------
                Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksOben, .Gebaeudebezogene_Daten.Ergaenzdende_Erlaeuterungen.Wert, 244 + 20, 4040, 4595, 5620, Font_Schriftgroesse_40, Font_Druckfarbe)
                '--------------------------------------------------
            End With
            '--------------------------------------------------
            If Variable_Steuerung.Entwurf = True Then
                Image_Text_im_Winkel_schreiben(Grafik_Seite, -45, "ENTWURF", 0, 2100, 0, 2970, Font_Schriftgroesse_ENTWURF, Font_Druckfarbe_ENTWURF)
            End If
            '--------------------------------------------------
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
            Hintergrundgrafik = Image.FromFile(Variable_Parameter.Arbeitsverzeichnis & "/Image/ENEV2016_NWB_Seite_5.png")
            Grafik_Seite.DrawImage(Hintergrundgrafik, 0, 0, 4200, 5940)
            '-------------------------------------------------------
            With Variable_XML_Import
                '-------------------------------------------------------
                Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, Datum_Energiegesetz, 2458, 2919, 446, 544, Font_Schriftgroesse_50, Font_Druckfarbe)
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
#Region "Zusatzseite Energiebedarf"
            '-------------------------------------------------------
            Select Case Variable_Steuerung.Berechnungsart
                Case "Energiebedarfsausweis"
                    '-------------------------------------------------------
                    Select Case Variable_XML_Import.Berechnungsverfahren
                        Case "Bedarfsberechnung-18599"
                            '-------------------------------------------------------
                            Wert_Schleife_Offset = 3
                            Wert_Schleife_Maximum = 200
                            Wert_Schleife_Zeilenzahl = 50
                            '-------------------------------------------------------
                            For Wert_Schleife = 0 To (Wert_Schleife_Maximum) Step Wert_Schleife_Zeilenzahl
                                '-------------------------------------------------------
                                With Variable_XML_Import
                                    '-------------------------------------------------------
                                    If .Energiebedarf_Anzahl > (Wert_Schleife + Wert_Schleife_Offset) Then
                                        '-------------------------------------------------------
                                        Seitennummer += 1
                                        '-------------------------------------------------------
                                        Grafik_Seite.Clear(Color.White)
                                        '-------------------------------------------------------
                                        Hintergrundgrafik = Image.FromFile(Variable_Parameter.Arbeitsverzeichnis & "/Image/ENEV2016_NWB_Ueberlauf_Bedarf.png")
                                        Grafik_Seite.DrawImage(Hintergrundgrafik, 0, 0, 4200, 5940)
                                        '-------------------------------------------------------
                                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, Seitennummer, 0, "", 3893, 4040, 688, 840, Font_Seitennummer, Font_Druckfarbe)
                                        '-------------------------------------------------------
                                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, Datum_Energiegesetz, 2458, 2919, 446, 544, Font_Schriftgroesse_50, Font_Druckfarbe)
                                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.Mitte_kein_Abschneiden, Variable_Steuerung.Registriernummer, 3049, 3764, 722, 820, Font_Bold_Schriftgroesse_40_Narrow, Font_Druckfarbe)
                                        '-------------------------------------------------------
                                        Zeilenhoehe = (5408 - 1310) / Wert_Schleife_Zeilenzahl
                                        '-------------------------------------------------------
                                        If .Energiebedarf_Anzahl < (Wert_Schleife + Wert_Schleife_Zeilenzahl + Wert_Schleife_Offset) Then
                                            '-------------------------------------------------------
                                            Wert_Schleife_Zeilenzahl = .Energiebedarf_Anzahl - Wert_Schleife_Offset - Wert_Schleife
                                            '-------------------------------------------------------
                                        End If
                                        '-------------------------------------------------------
                                        For Anzahl_Zeilen = 1 To Wert_Schleife_Zeilenzahl
                                            '-------------------------------------------------------
                                            Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Energietraeger(Anzahl_Zeilen + Wert_Schleife + Wert_Schleife_Offset).Energietraegerbezeichnung.Wert, 244, 1760, 1310 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1310 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                            '-------------------------------------------------------
                                            If .Energietraeger(Anzahl_Zeilen + Wert_Schleife + Wert_Schleife_Offset).Endenergiebedarf_Heizung_spezifisch.Wert <> 0 Then
                                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energietraeger(Anzahl_Zeilen + Wert_Schleife + Wert_Schleife_Offset).Endenergiebedarf_Heizung_spezifisch.Wert, 1, "", 1764, 2054, 1310 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1310 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                            End If
                                            '-------------------------------------------------------
                                            If .Energietraeger(Anzahl_Zeilen + Wert_Schleife + Wert_Schleife_Offset).Endenergiebedarf_Trinkwarmwasser_spezifisch.Wert <> 0 Then
                                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energietraeger(Anzahl_Zeilen + Wert_Schleife + Wert_Schleife_Offset).Endenergiebedarf_Trinkwarmwasser_spezifisch.Wert, 1, "", 2058, 2455, 1310 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1310 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                            End If
                                            '-------------------------------------------------------
                                            If .Energietraeger(Anzahl_Zeilen + Wert_Schleife + Wert_Schleife_Offset).Endenergiebedarf_Beleuchtung_spezifisch.Wert <> 0 Then
                                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energietraeger(Anzahl_Zeilen + Wert_Schleife + Wert_Schleife_Offset).Endenergiebedarf_Beleuchtung_spezifisch.Wert, 1, "", 2459, 2859, 1310 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1310 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                            End If
                                            '-------------------------------------------------------
                                            If .Energietraeger(Anzahl_Zeilen + Wert_Schleife + Wert_Schleife_Offset).Endenergiebedarf_Lueftung_spezifisch.Wert <> 0 Then
                                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energietraeger(Anzahl_Zeilen + Wert_Schleife + Wert_Schleife_Offset).Endenergiebedarf_Lueftung_spezifisch.Wert, 1, "", 2863, 3183, 1310 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1310 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                            End If
                                            '-------------------------------------------------------
                                            If .Energietraeger(Anzahl_Zeilen + Wert_Schleife + Wert_Schleife_Offset).Endenergiebedarf_Kuehlung_Befeuchtung_spezifisch.Wert <> 0 Then
                                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energietraeger(Anzahl_Zeilen + Wert_Schleife + Wert_Schleife_Offset).Endenergiebedarf_Kuehlung_Befeuchtung_spezifisch.Wert, 1, "", 3187, 3651, 1310 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1310 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                            End If
                                            '-------------------------------------------------------
                                            If .Energietraeger(Anzahl_Zeilen + Wert_Schleife + Wert_Schleife_Offset).Endenergiebedarf_Energietraeger_Gesamtgebaeude_spezifisch.Wert <> 0 Then
                                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energietraeger(Anzahl_Zeilen + Wert_Schleife + Wert_Schleife_Offset).Endenergiebedarf_Energietraeger_Gesamtgebaeude_spezifisch.Wert, 1, "", 3654, 4033, 1310 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1310 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                            End If
                                            '-------------------------------------------------------
                                        Next
                                        '-------------------------------------------------------
                                        For Anzahl_Zeilen = 1 To Wert_Schleife_Zeilenzahl
                                            '-------------------------------------------------------
                                            Image_Linie(Grafik_Seite, 244, 4033, 1310 + (Zeilenhoehe * Anzahl_Zeilen), 1310 + (Zeilenhoehe * Anzahl_Zeilen))
                                            '-------------------------------------------------------
                                        Next
                                        '-------------------------------------------------------
                                        If .Energiebedarf_Anzahl > (Wert_Schleife + Wert_Schleife_Zeilenzahl + Wert_Schleife_Offset) Then
                                            '-------------------------------------------------------
                                            Image_Auswahl_schreiben(Grafik_Seite, True, 300, 330, 5495, 5525)
                                            '-------------------------------------------------------
                                        Else
                                            '-------------------------------------------------------
                                            Image_Auswahl_schreiben(Grafik_Seite, False, 300, 330, 5495, 5525)
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
                                            Listview.Items.Add("Energiebedarf", Seitennummer - 1)
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
                    End Select
                    '-------------------------------------------------------
            End Select
            '-------------------------------------------------------
#End Region
            '-------------------------------------------------------
#Region "Zusatzseite Zonen"
            '-------------------------------------------------------
            Select Case Variable_Steuerung.Berechnungsart
                Case "Energiebedarfsausweis"
                    '-------------------------------------------------------
                    Select Case Variable_XML_Import.Berechnungsverfahren
                        Case "Bedarfsberechnung-18599"
                            '-------------------------------------------------------
                            Wert_Schleife_Offset = 7
                            Wert_Schleife_Maximum = 200
                            Wert_Schleife_Zeilenzahl = 50
                            '-------------------------------------------------------
                            For Wert_Schleife = 0 To (Wert_Schleife_Maximum) Step Wert_Schleife_Zeilenzahl
                                '-------------------------------------------------------
                                With Variable_XML_Import
                                    '-------------------------------------------------------
                                    If .Zonen_Anzahl > (Wert_Schleife + Wert_Schleife_Offset) Then
                                        '-------------------------------------------------------
                                        Seitennummer += 1
                                        '-------------------------------------------------------
                                        Grafik_Seite.Clear(Color.White)
                                        '-------------------------------------------------------
                                        Hintergrundgrafik = Image.FromFile(Variable_Parameter.Arbeitsverzeichnis & "/Image/ENEV2016_NWB_Ueberlauf_Gebaeudezonen.png")
                                        Grafik_Seite.DrawImage(Hintergrundgrafik, 0, 0, 4200, 5940)
                                        '-------------------------------------------------------
                                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, Seitennummer, 0, "", 3893, 4040, 688, 840, Font_Seitennummer, Font_Druckfarbe)
                                        '-------------------------------------------------------
                                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, Datum_Energiegesetz, 2458, 2919, 446, 544, Font_Schriftgroesse_50, Font_Druckfarbe)
                                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.Mitte_kein_Abschneiden, Variable_Steuerung.Registriernummer, 3032, 3748, 718, 818, Font_Bold_Schriftgroesse_40_Narrow, Font_Druckfarbe)
                                        '-------------------------------------------------------
                                        Zeilenhoehe = (5389 - 1189) / Wert_Schleife_Zeilenzahl
                                        '-------------------------------------------------------
                                        If .Zonen_Anzahl < (Wert_Schleife + Wert_Schleife_Zeilenzahl + Wert_Schleife_Offset) Then
                                            '-------------------------------------------------------
                                            Wert_Schleife_Zeilenzahl = .Zonen_Anzahl - Wert_Schleife_Offset - Wert_Schleife
                                            '-------------------------------------------------------
                                        End If
                                        '-------------------------------------------------------
                                        For Anzahl_Zeilen = 1 To Wert_Schleife_Zeilenzahl
                                            '-------------------------------------------------------
                                            Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, Anzahl_Zeilen + Wert_Schleife + Wert_Schleife_Offset, 0, "", 244, 438, 1189 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1189 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                            '-------------------------------------------------------
                                            Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Energieausweis_Daten.Nichtwohngebaeude.Zone(Anzahl_Zeilen + Wert_Schleife + Wert_Schleife_Offset).Zonenbezeichnung.Wert, 442, 3318, 1189 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1189 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                            '-------------------------------------------------------
                                            If .Energieausweis_Daten.Nichtwohngebaeude.Zone(Anzahl_Zeilen + Wert_Schleife + Wert_Schleife_Offset).Nettogrundflaeche_Zone.Wert <> 0 Then
                                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Nichtwohngebaeude.Zone(Anzahl_Zeilen + Wert_Schleife + Wert_Schleife_Offset).Nettogrundflaeche_Zone.Wert, 1, "", 3322, 3668, 1189 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1189 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                            End If
                                            '-------------------------------------------------------
                                            If .Energieausweis_Daten.Nichtwohngebaeude.Zone(Anzahl_Zeilen + Wert_Schleife + Wert_Schleife_Offset).Nettogrundflaeche_Zone.Wert <> 0 Then
                                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, (.Energieausweis_Daten.Nichtwohngebaeude.Zone(Anzahl_Zeilen + Wert_Schleife + Wert_Schleife_Offset).Nettogrundflaeche_Zone.Wert / .Energieausweis_Daten.Nichtwohngebaeude.Nettogrundflaeche.Wert * 100), 2, "", 3672, 4032, 1189 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1189 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                            End If
                                            '-------------------------------------------------------
                                        Next
                                        '-------------------------------------------------------
                                        For Anzahl_Zeilen = 1 To Wert_Schleife_Zeilenzahl
                                            '-------------------------------------------------------
                                            Image_Linie(Grafik_Seite, 244, 4033, 1189 + (Zeilenhoehe * Anzahl_Zeilen), 1189 + (Zeilenhoehe * Anzahl_Zeilen))
                                            '-------------------------------------------------------
                                        Next
                                        '-------------------------------------------------------
                                        If .Zonen_Anzahl > (Wert_Schleife + Wert_Schleife_Zeilenzahl + Wert_Schleife_Offset) Then
                                            '-------------------------------------------------------
                                            Image_Auswahl_schreiben(Grafik_Seite, True, 308, 338, 5459, 5488)
                                            '-------------------------------------------------------
                                        Else
                                            '-------------------------------------------------------
                                            Image_Auswahl_schreiben(Grafik_Seite, False, 308, 338, 5459, 5488)
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
                                            Listview.Items.Add("Zonen", Seitennummer - 1)
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
                    End Select
                    '-------------------------------------------------------
            End Select
            '-------------------------------------------------------
#End Region
            '-------------------------------------------------------
#Region "Zusatzseite Verbrauchserfassung"
            '-------------------------------------------------------
            If Variable_Steuerung.Berechnungsart = "Energieverbrauchsausweis" Or Variable_XML_Import.Zusaetzliche_Verbrauchsberechnung = True Then
                '-------------------------------------------------------
                If Variable_XML_Import.Berechnungsverfahren = "Verbrauchsberechnung" Or Variable_XML_Import.Zusaetzliche_Verbrauchsberechnung = True Then
                    '-------------------------------------------------------
                    Wert_Schleife_Offset = 3
                    Wert_Schleife_Maximum = 364
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
                                Hintergrundgrafik = Image.FromFile(Variable_Parameter.Arbeitsverzeichnis & "/Image/ENEV2016_NWB_Ueberlauf_Verbrauch.png")
                                Grafik_Seite.DrawImage(Hintergrundgrafik, 0, 0, 4200, 5940)
                                '-------------------------------------------------------
                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, Seitennummer, 0, "", 3893, 4040, 688, 840, Font_Seitennummer, Font_Druckfarbe)
                                '-------------------------------------------------------
                                Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, Datum_Energiegesetz, 2458, 2919, 446, 544, Font_Schriftgroesse_50, Font_Druckfarbe)
                                Image_Text_schreiben(Grafik_Seite, Ausrichtung.Mitte_kein_Abschneiden, Variable_Steuerung.Registriernummer, 3049, 3764, 722, 820, Font_Bold_Schriftgroesse_40_Narrow, Font_Druckfarbe)
                                '-------------------------------------------------------
                                Zeilenhoehe = (5384 - 1389) / Wert_Schleife_Zeilenzahl
                                '-------------------------------------------------------
                                If .Verbrauchserfassung_Anzahl < (Wert_Schleife + Wert_Schleife_Zeilenzahl + Wert_Schleife_Offset) Then
                                    '-------------------------------------------------------
                                    Wert_Schleife_Zeilenzahl = .Verbrauchserfassung_Anzahl - Wert_Schleife_Offset - Wert_Schleife
                                    '-------------------------------------------------------
                                End If
                                '-------------------------------------------------------
                                For Anzahl_Zeilen = 1 To Wert_Schleife_Zeilenzahl
                                    '-------------------------------------------------------
                                    Image_Datum_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Verbrauchserfassung(Anzahl_Zeilen + Wert_Schleife + Wert_Schleife_Offset).Startdatum.Wert, 249, 573, 1389 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1389 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                    '--------------------------------------------------
                                    Image_Datum_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Verbrauchserfassung(Anzahl_Zeilen + Wert_Schleife + Wert_Schleife_Offset).Enddatum.Wert, 577, 902, 1389 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1389 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                    '--------------------------------------------------
                                    Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Verbrauchserfassung(Anzahl_Zeilen + Wert_Schleife + Wert_Schleife_Offset).Energietraeger_Verbrauch.Wert, 906, 1793, 1389 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1389 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                    '--------------------------------------------------
                                    If .Verbrauchserfassung(Anzahl_Zeilen + Wert_Schleife + Wert_Schleife_Offset).Primaerenergiefaktor.Wert <> 0 Then
                                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Verbrauchserfassung(Anzahl_Zeilen + Wert_Schleife + Wert_Schleife_Offset).Primaerenergiefaktor.Wert, 1, "", 1797, 2034, 1389 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1389 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                    End If
                                    '--------------------------------------------------
                                    If .Verbrauchserfassung(Anzahl_Zeilen + Wert_Schleife + Wert_Schleife_Offset).Energieverbrauch.Wert <> 0 Then
                                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Verbrauchserfassung(Anzahl_Zeilen + Wert_Schleife + Wert_Schleife_Offset).Energieverbrauch.Wert, 0, "", 2037, 2393, 1389 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1389 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                    End If
                                    '--------------------------------------------------
                                    If .Verbrauchserfassung(Anzahl_Zeilen + Wert_Schleife + Wert_Schleife_Offset).Energieverbrauchsanteil_Warmwasser_zentral.Wert <> 0 Then
                                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Verbrauchserfassung(Anzahl_Zeilen + Wert_Schleife + Wert_Schleife_Offset).Energieverbrauchsanteil_Warmwasser_zentral.Wert, 0, "", 2397, 2753, 1389 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1389 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                    End If
                                    '--------------------------------------------------
                                    If .Verbrauchserfassung(Anzahl_Zeilen + Wert_Schleife + Wert_Schleife_Offset).Energieverbrauchsanteil_thermisch_erzeugte_Kaelte.Wert <> 0 Then
                                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Verbrauchserfassung(Anzahl_Zeilen + Wert_Schleife + Wert_Schleife_Offset).Energieverbrauchsanteil_thermisch_erzeugte_Kaelte.Wert, 0, "", 2757, 3113, 1389 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1389 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                    End If
                                    '--------------------------------------------------
                                    If .Verbrauchserfassung(Anzahl_Zeilen + Wert_Schleife + Wert_Schleife_Offset).Energieverbrauchsanteil_Heizung.Wert <> 0 Then
                                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Verbrauchserfassung(Anzahl_Zeilen + Wert_Schleife + Wert_Schleife_Offset).Energieverbrauchsanteil_Heizung.Wert, 0, "", 3117, 3474, 1389 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1389 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                    End If
                                    '--------------------------------------------------
                                    If .Verbrauchserfassung(Anzahl_Zeilen + Wert_Schleife + Wert_Schleife_Offset).Klimafaktor.Wert <> 0 Then
                                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Verbrauchserfassung(Anzahl_Zeilen + Wert_Schleife + Wert_Schleife_Offset).Klimafaktor.Wert, 2, "", 3478, 3678, 1389 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1389 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                    End If
                                    '-------------------------------------------------------
                                    If .Verbrauchserfassung(Anzahl_Zeilen + Wert_Schleife + Wert_Schleife_Offset).Energieverbrauch_Strom.Wert <> 0 Then
                                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Verbrauchserfassung(Anzahl_Zeilen + Wert_Schleife + Wert_Schleife_Offset).Energieverbrauch_Strom.Wert, 0, "", 3682, 4032, 1389 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1389 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                    End If
                                    '-------------------------------------------------------
                                Next
                                '-------------------------------------------------------
                                For Anzahl_Zeilen = 1 To Wert_Schleife_Zeilenzahl
                                    '-------------------------------------------------------
                                    Image_Linie(Grafik_Seite, 249, 4032, 1389 + (Zeilenhoehe * Anzahl_Zeilen), 1389 + (Zeilenhoehe * Anzahl_Zeilen))
                                    '-------------------------------------------------------
                                Next
                                '-------------------------------------------------------
                                If .Verbrauchserfassung_Anzahl > (Wert_Schleife + Wert_Schleife_Zeilenzahl + Wert_Schleife_Offset) Then
                                    '-------------------------------------------------------
                                    Image_Auswahl_schreiben(Grafik_Seite, True, 300, 330, 5462, 5492)
                                    '-------------------------------------------------------
                                Else
                                    '-------------------------------------------------------
                                    Image_Auswahl_schreiben(Grafik_Seite, False, 300, 330, 5462, 5492)
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
            Wert_Schleife_Offset = 10
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
                        Hintergrundgrafik = Image.FromFile(Variable_Parameter.Arbeitsverzeichnis & "/Image/ENEV2016_NWB_Ueberlauf_Modernisierung.png")
                        Grafik_Seite.DrawImage(Hintergrundgrafik, 0, 0, 4200, 5940)
                        '-------------------------------------------------------
                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, Seitennummer, 0, "", 3893, 4040, 688, 840, Font_Seitennummer, Font_Druckfarbe)
                        '-------------------------------------------------------
                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, Datum_Energiegesetz, 2458, 2919, 446, 544, Font_Schriftgroesse_50, Font_Druckfarbe)
                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.Mitte_kein_Abschneiden, Variable_Steuerung.Registriernummer, 3053, 3768, 722, 820, Font_Bold_Schriftgroesse_40_Narrow, Font_Druckfarbe)
                        '-------------------------------------------------------
                        If .Energieausweis_Daten.Empfehlungen_moeglich.Wert = False Then
                            '-------------------------------------------------------
                            Image_Auswahl_schreiben(Grafik_Seite, False, 2425, 2471, 1143, 1176)
                            Image_Auswahl_schreiben(Grafik_Seite, True, 2833, 2880, 1143, 1176)
                            '-------------------------------------------------------
                        Else
                            '-------------------------------------------------------
                            Image_Auswahl_schreiben(Grafik_Seite, True, 2425, 2471, 1143, 1176)
                            Image_Auswahl_schreiben(Grafik_Seite, False, 2833, 2880, 1143, 1176)
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
                                        '-------------------------------------------------------
                                        Image_Auswahl_schreiben(Grafik_Seite, True, 2703, 2750, 1922 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1969 + (Zeilenhoehe * (Anzahl_Zeilen - 1)))
                                        Image_Auswahl_schreiben(Grafik_Seite, False, 2951, 2998, 1922 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1969 + (Zeilenhoehe * (Anzahl_Zeilen - 1)))
                                        '-------------------------------------------------------
                                    Case "als Einzelmaßnahme"
                                        '-------------------------------------------------------
                                        Image_Auswahl_schreiben(Grafik_Seite, False, 2703, 2750, 1922 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1969 + (Zeilenhoehe * (Anzahl_Zeilen - 1)))
                                        Image_Auswahl_schreiben(Grafik_Seite, True, 2951, 2998, 1922 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1969 + (Zeilenhoehe * (Anzahl_Zeilen - 1)))
                                        '-------------------------------------------------------
                                End Select
                                '-------------------------------------------------------
                                Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksMitte, .Modernisierungsempfehlungen(Anzahl_Zeilen + Wert_Schleife + Wert_Schleife_Offset).Amortisation.Wert, 3086 + 20, 3412, 1882 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1882 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                '-------------------------------------------------------
                                Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksMitte, .Modernisierungsempfehlungen(Anzahl_Zeilen + Wert_Schleife + Wert_Schleife_Offset).spezifische_Kosten.Wert, 3419 + 20, 4032, 1882 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1882 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                '-------------------------------------------------------
                            Next
                            '-------------------------------------------------------
                            If .Modernisierungsempfehlung_Anzahl > (Wert_Schleife + Wert_Schleife_Zeilenzahl + Wert_Schleife_Offset) Then
                                '-------------------------------------------------------
                                Image_Auswahl_schreiben(Grafik_Seite, True, 317, 359, 5459, 5500)
                                '-------------------------------------------------------
                            Else
                                '-------------------------------------------------------
                                Image_Auswahl_schreiben(Grafik_Seite, False, 317, 359, 5459, 5500)
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
#Region "Zusatzseite Aushang - Bedarf"
            '-------------------------------------------------------
            Select Case Variable_Steuerung.Berechnungsart
                Case "Energiebedarfsausweis"
                    '-------------------------------------------------------
                    Select Case Variable_XML_Import.Berechnungsverfahren
                        Case "Bedarfsberechnung-18599"
                            '-------------------------------------------------------
                            Seitennummer += 1
                            '-------------------------------------------------------
                            Grafik_Seite.Clear(Color.White)
                            '-------------------------------------------------------
                            Hintergrundgrafik = Image.FromFile(Variable_Parameter.Arbeitsverzeichnis & "/Image/ENEV2016_NWB_Bedarf_Aushang.png")
                            Grafik_Seite.DrawImage(Hintergrundgrafik, 0, 0, 4200, 5940)
                            '-------------------------------------------------------
                            With Variable_XML_Import
                                '-------------------------------------------------------
                                Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, Datum_Energiegesetz, 2477, 2940, 450, 547, Font_Schriftgroesse_50, Font_Druckfarbe)
                                Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, DateAdd(DateInterval.Year, 10, Variable_Steuerung.Ausstellungsdatum), 549 + 20, 1000, 771, 821, Font_Schriftgroesse_50, Font_Druckfarbe)
                                Image_Text_schreiben(Grafik_Seite, Ausrichtung.Mitte_kein_Abschneiden, Variable_Steuerung.Registriernummer, 2818, 3281, 645, 743, Font_Bold_Schriftgroesse_40_Narrow, Font_Druckfarbe)
                                '-------------------------------------------------------
                                Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksOben, .Energieausweis_Daten.Nichtwohngebaeude.Hauptnutzung_Gebaeudekategorie.Wert, 1273 + 20, 3266, 1163, 1372, Font_Schriftgroesse_50, Font_Druckfarbe)
                                Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksMitte, .Gebaeudebezogene_Daten.Gebaeudeadresse_Strasse_Nr.Wert & ", " & .Gebaeudebezogene_Daten.Gebaeudeadresse_Postleitzahl.Wert & " " & .Gebaeudebezogene_Daten.Gebaeudeadresse_Ort.Wert, 1273 + 20, 3266, 1381, 1502, Font_Schriftgroesse_50, Font_Druckfarbe)
                                Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Energieausweis_Daten.Gebaeudeteil.Wert, 1273 + 20, 3266, 1511, 1633, Font_Schriftgroesse_50, Font_Druckfarbe)
                                Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Energieausweis_Daten.Baujahr_Gebaeude.Wert, 1273 + 20, 3266, 1641, 1761, Font_Schriftgroesse_50, Font_Druckfarbe)
                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Energieausweis_Daten.Nichtwohngebaeude.Nettogrundflaeche.Wert, 2, "m²", 1273 + 20, 3266, 1770, 1891, Font_Schriftgroesse_50, Font_Druckfarbe)
                                Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksMitte, .Energieausweis_Daten.wesentliche_Energietraeger_Heizung.Wert, 1273 + 20, 3266, 1900, 2110, Font_Schriftgroesse_50, Font_Druckfarbe)
                                '-------------------------------------------------------
                                Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Energieausweis_Daten.Erneuerbare_Art.Wert, 1425 + 20, 2653, 2118, 2226, Font_Schriftgroesse_50, Font_Druckfarbe)
                                Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Energieausweis_Daten.Erneuerbare_Verwendung.Wert, 3130 + 20, 4030, 2118, 2226, Font_Schriftgroesse_50, Font_Druckfarbe)
                                '-------------------------------------------------------
                                Image_Pfeil_NWB(Grafik_Seite, Lage.Oben, .Energieausweis_Daten.Nichtwohngebaeude.Bedarfswerte_NWG.Primaerenergiebedarf_NGF.Wert, 1, .Energieausweis_Daten.Nichtwohngebaeude.Bedarfswerte_NWG.Primaerenergiebedarf_Hoechstwert_Bestand.Wert, 3, "Primärenergiebedarf dieses Gebäudes", "", 635, 3654, 2799, 3151)
                                '-------------------------------------------------------
                                Image_Diagram_NWB_Energieaufteilung(Grafik_Seite, 412, 2346, 3842, 4624)
                                '-------------------------------------------------------
                                Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, FormatDateTime(Variable_Steuerung.Ausstellungsdatum, DateFormat.ShortDate).ToString, 2252, 2778, 5370, 5452, Font_Schriftgroesse_50, Font_Druckfarbe)
                                '-------------------------------------------------------
                                Image_Grafik_Projekt_schreiben(Grafik_Seite, Variable_Steuerung.Bild_Projekt, 3275, 1163, 4030 - 3275, 2110 - 1163)
                                '-------------------------------------------------------
                                Image_Grafik_Aussteller_schreiben(Grafik_Seite, Variable_Steuerung.Bild_Aussteller, .Gebaeudebezogene_Daten.Aussteller_Bezeichnung.Wert, .Gebaeudebezogene_Daten.Ausstellervorname.Wert, .Gebaeudebezogene_Daten.Ausstellername.Wert, .Gebaeudebezogene_Daten.Aussteller_Strasse_Nr.Wert, .Gebaeudebezogene_Daten.Aussteller_PLZ.Wert, .Gebaeudebezogene_Daten.Aussteller_Ort.Wert, 270, 5076, 1750, 400)
                                '-------------------------------------------------------
                                Image_Grafik_Unterschrift_schreiben(Grafik_Seite, Variable_Steuerung.Bild_Unterschrift, 3000, 5150, 1000, 300)
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
                                Listview.Items.Add("Aushang - Bedarf", Seitennummer - 1)
                                Listview.Update()
                                '-------------------------------------------------------
                            End If
                            '-------------------------------------------------------
                            Variable_XML_Import.Bilddateien(Seitennummer) = Image_to_String(Bitmap_Seite)
                            Variable_XML_Import.Bildanzahl += 1
                            '-------------------------------------------------------
                            Speicher_freigeben()
                            '-------------------------------------------------------
                    End Select
                    '-------------------------------------------------------
            End Select
            '-------------------------------------------------------
#End Region
            '-------------------------------------------------------
#Region "Zusatzseite Aushang - Verbrauch"
            '-------------------------------------------------------
            If Variable_Steuerung.Berechnungsart = "Energieverbrauchsausweis" Or Variable_XML_Import.Zusaetzliche_Verbrauchsberechnung = True Then
                '-------------------------------------------------------
                If Variable_XML_Import.Berechnungsverfahren = "Verbrauchsberechnung" Or Variable_XML_Import.Zusaetzliche_Verbrauchsberechnung = True Then
                    '-------------------------------------------------------
                    Seitennummer += 1
                    '-------------------------------------------------------
                    Grafik_Seite.Clear(Color.White)
                    '-------------------------------------------------------
                    Hintergrundgrafik = Image.FromFile(Variable_Parameter.Arbeitsverzeichnis & "/Image/ENEV2016_NWB_Verbrauch_Aushang.png")
                    Grafik_Seite.DrawImage(Hintergrundgrafik, 0, 0, 4200, 5940)
                    '-------------------------------------------------------
                    With Variable_XML_Import
                        '-------------------------------------------------------
                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, Datum_Energiegesetz, 2477, 2940, 450, 547, Font_Schriftgroesse_50, Font_Druckfarbe)
                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, DateAdd(DateInterval.Year, 10, Variable_Steuerung.Ausstellungsdatum), 549 + 20, 1000, 771, 821, Font_Schriftgroesse_50, Font_Druckfarbe)
                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.Mitte_kein_Abschneiden, Variable_Steuerung.Registriernummer, 2818, 3281, 645, 743, Font_Bold_Schriftgroesse_40_Narrow, Font_Druckfarbe)
                        '-------------------------------------------------------
                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksOben, .Energieausweis_Daten.Nichtwohngebaeude.Hauptnutzung_Gebaeudekategorie.Wert, 1277 + 20, 3311, 1122, 1332, Font_Schriftgroesse_50, Font_Druckfarbe)
                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksMitte, .Gebaeudebezogene_Daten.Gebaeudeadresse_Strasse_Nr.Wert & ", " & .Gebaeudebezogene_Daten.Gebaeudeadresse_Postleitzahl.Wert & " " & .Gebaeudebezogene_Daten.Gebaeudeadresse_Ort.Wert, 1277 + 20, 3311, 1340, 1461, Font_Schriftgroesse_50, Font_Druckfarbe)
                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Energieausweis_Daten.Gebaeudeteil.Wert, 1277 + 20, 3311, 1469, 1591, Font_Schriftgroesse_50, Font_Druckfarbe)
                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Energieausweis_Daten.Baujahr_Gebaeude.Wert, 1277 + 20, 3311, 1600, 1721, Font_Schriftgroesse_50, Font_Druckfarbe)
                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Energieausweis_Daten.Nichtwohngebaeude.Nettogrundflaeche.Wert, 2, "m²", 1277 + 20, 3311, 1729, 1852, Font_Schriftgroesse_50, Font_Druckfarbe)
                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksMitte, .Energieausweis_Daten.wesentliche_Energietraeger_Heizung.Wert, 1277 + 20, 3311, 1860, 2068, Font_Schriftgroesse_50, Font_Druckfarbe)
                        '-------------------------------------------------------
                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Energieausweis_Daten.Erneuerbare_Art.Wert, 1425 + 20, 2653, 2077, 2197, Font_Schriftgroesse_50, Font_Druckfarbe)
                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Energieausweis_Daten.Erneuerbare_Verwendung.Wert, 3130 + 20, 4030, 2077, 2197, Font_Schriftgroesse_50, Font_Druckfarbe)
                        '-------------------------------------------------------
                        Image_Pfeil_NWB(Grafik_Seite, Lage.Oben, .Energieausweis_Daten.Nichtwohngebaeude.Verbrauchswerte_NWG.Endenergieverbrauch_Waerme.Wert, 1, .Energieausweis_Daten.Nichtwohngebaeude.Verbrauchswerte_NWG.Endenergieverbrauch_Waerme_Vergleichswert.Wert, 2, "Endenergieverbrauch Wärme", "[Pflichtangabe in Immobilienanzeigen]", 644, 3664, 2815, 3152)
                        '-------------------------------------------------------
                        Image_Pfeil_NWB(Grafik_Seite, Lage.Oben, .Energieausweis_Daten.Nichtwohngebaeude.Verbrauchswerte_NWG.Endenergieverbrauch_Strom.Wert, 1, .Energieausweis_Daten.Nichtwohngebaeude.Verbrauchswerte_NWG.Endenergieverbrauch_Strom_Vergleichswert.Wert, 2, "Endenergieverbrauch Strom", "[Pflichtangabe in Immobilienanzeigen]", 638, 3658, 3831, 4168)
                        '-------------------------------------------------------
                        Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Nichtwohngebaeude.Verbrauchswerte_NWG.Warmwasser_enthalten.Wert, 273, 312, 3397, 3435)
                        '-------------------------------------------------------
                        Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Nichtwohngebaeude.Verbrauchswerte_NWG.Strom_Daten.Stromverbrauch_enthaelt_Zusatzheizung.Wert, 273, 312, 4548, 4587)
                        Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Nichtwohngebaeude.Verbrauchswerte_NWG.Strom_Daten.Stromverbrauch_enthaelt_Warmwasser.Wert, 972, 1010, 4548, 4587)
                        Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Nichtwohngebaeude.Verbrauchswerte_NWG.Strom_Daten.Stromverbrauch_enthaelt_Lueftung.Wert, 1643, 1683, 4548, 4587)
                        Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Nichtwohngebaeude.Verbrauchswerte_NWG.Strom_Daten.Stromverbrauch_enthaelt_Beleuchtung.Wert, 2116, 2155, 4548, 4587)
                        Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Nichtwohngebaeude.Verbrauchswerte_NWG.Strom_Daten.Stromverbrauch_enthaelt_Kuehlung.Wert, 3092, 3131, 4548, 4587)
                        Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Nichtwohngebaeude.Verbrauchswerte_NWG.Strom_Daten.Stromverbrauch_enthaelt_Sonstiges.Wert, 3563, 3603, 4548, 4587)
                        '-------------------------------------------------------
                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Nichtwohngebaeude.Verbrauchswerte_NWG.Primaerenergieverbrauch.Wert, 1, "", 3313, 3689, 4672, 4823, Font_Schriftgroesse_60, Font_Druckfarbe)
                        '-------------------------------------------------------
                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, FormatDateTime(Variable_Steuerung.Ausstellungsdatum, DateFormat.ShortDate).ToString, 2252, 2778, 5350, 5432, Font_Schriftgroesse_50, Font_Druckfarbe)
                        '-------------------------------------------------------
                        Image_Grafik_Projekt_schreiben(Grafik_Seite, Variable_Steuerung.Bild_Projekt, 3319, 1122, 4033 - 3319, 2068 - 1122)
                        '-------------------------------------------------------
                        Image_Grafik_Aussteller_schreiben(Grafik_Seite, Variable_Steuerung.Bild_Aussteller, .Gebaeudebezogene_Daten.Aussteller_Bezeichnung.Wert, .Gebaeudebezogene_Daten.Ausstellervorname.Wert, .Gebaeudebezogene_Daten.Ausstellername.Wert, .Gebaeudebezogene_Daten.Aussteller_Strasse_Nr.Wert, .Gebaeudebezogene_Daten.Aussteller_PLZ.Wert, .Gebaeudebezogene_Daten.Aussteller_Ort.Wert, 270, 5076, 1750, 400)
                        '-------------------------------------------------------
                        Image_Grafik_Unterschrift_schreiben(Grafik_Seite, Variable_Steuerung.Bild_Unterschrift, 3000, 5100, 1000, 300)
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
                        Listview.Items.Add("Aushang - Verbrauch", Seitennummer - 1)
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
            End If
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
