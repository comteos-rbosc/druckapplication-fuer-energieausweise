#Region "Imports"
Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms
#End Region
'--------------------------------------------------
Public Class Klasse_Energieausweis_NWB_2024
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
    ''' Hier wird der gesamte Ablauf für die Erstellung der Energieausweises abgearbeitet - Nichtwohnbau - GEG 2024
    ''' </summary>
    Sub Energieausweis_erzeugen(ByVal Oberflaeche As Boolean, ByVal PictureBox As PictureBox, ByVal Listview As ListView, ByVal Imagelist As ImageList)
        '--------------------------------------------------
        Try
            '-------------------------------------------------------
#Region "Variablen"
            '-------------------------------------------------------
            Dim Datum_Energiegesetz As String = "16.10.2023"
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
            '-------------------------------------------------------
            Grafik_Seite.CompositingQuality = Bitmap_CompositingQuality
            Grafik_Seite.TextRenderingHint = Bitmap_TextRenderingHint
            Grafik_Seite.SmoothingMode = Bitmap_SmoothingMode
            '-------------------------------------------------------
            Grafik_Seite.Clear(Color.White)
            '-------------------------------------------------------
            Hintergrundgrafik = Image.FromFile(Variable_Parameter.Arbeitsverzeichnis & "/Image/GEG2024_NWB_Seite_1.png")
            Grafik_Seite.DrawImage(Hintergrundgrafik, 0, 0, 4200, 5940)
            '-------------------------------------------------------
            With Variable_XML_Import
                '-------------------------------------------------------
                Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, Datum_Energiegesetz, 2129, 2617, 422, 517, Font_Schriftgroesse_50, Font_Druckfarbe)
                Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, DateAdd(DateInterval.Year, 10, Variable_Steuerung.Ausstellungsdatum), 621 + 20, 1339, 718, 818, Font_Schriftgroesse_50, Font_Druckfarbe)
                Image_Text_schreiben(Grafik_Seite, Ausrichtung.Mitte_kein_Abschneiden, Variable_Steuerung.Registriernummer, 3032, 3748, 718, 818, Font_Schriftgroesse_50, Font_Druckfarbe)
                '-------------------------------------------------------
                If .Energieausweis_Daten.Nichtwohngebaeude.Hauptnutzung_Gebaeudekategorie_Sonstiges_Beschreibung.Wert = "" Then
                    Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksOben, .Energieausweis_Daten.Nichtwohngebaeude.Hauptnutzung_Gebaeudekategorie.Wert, 1470 + 20, 3248, 1067, 1252, Font_Schriftgroesse_40, Font_Druckfarbe)
                Else
                    Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksOben, .Energieausweis_Daten.Nichtwohngebaeude.Hauptnutzung_Gebaeudekategorie_Sonstiges_Beschreibung.Wert, 1470 + 20, 3248, 1067, 1252, Font_Schriftgroesse_40, Font_Druckfarbe)
                End If
                Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksMitte, .Gebaeudebezogene_Daten.Gebaeudeadresse_Strasse_Nr.Wert & ", " & .Gebaeudebezogene_Daten.Gebaeudeadresse_Postleitzahl.Wert & " " & .Gebaeudebezogene_Daten.Gebaeudeadresse_Ort.Wert, 1470 + 20, 3248, 1256, 1443, Font_Schriftgroesse_40, Font_Druckfarbe)
                Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Energieausweis_Daten.Gebaeudeteil.Wert, 1470 + 20, 3248, 1447, 1544, Font_Schriftgroesse_40, Font_Druckfarbe)
                Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Energieausweis_Daten.Baujahr_Gebaeude.Wert, 1470 + 20, 3248, 1548, 1645, Font_Schriftgroesse_40, Font_Druckfarbe)
                Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Energieausweis_Daten.Baujahr_Waermeerzeuger.Wert, 1470 + 20, 3248, 1649, 1834, Font_Schriftgroesse_40, Font_Druckfarbe)
                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Energieausweis_Daten.Nichtwohngebaeude.Nettogrundflaeche.Wert, 2, "m²", 1470 + 20, 3248, 1838, 1935, Font_Schriftgroesse_40, Font_Druckfarbe)
                Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksMitte, .Energieausweis_Daten.wesentliche_Energietraeger_Heizung.Wert, 1470 + 20, 4030, 1939, 2033, Font_Schriftgroesse_40, Font_Druckfarbe)
                Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksMitte, .Energieausweis_Daten.wesentliche_Energietraeger_Warmwasser.Wert, 1470 + 20, 4030, 2037, 2132, Font_Schriftgroesse_40, Font_Druckfarbe)
                Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksMitte, "Art: " & .Energieausweis_Daten.Erneuerbare_Art.Wert, 1470 + 20, 2490, 2136, 2321, Font_Schriftgroesse_40, Font_Druckfarbe)
                Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksMitte, "Verwendung: " & .Energieausweis_Daten.Erneuerbare_Verwendung.Wert, 2494 + 20, 4030, 2136, 2321, Font_Schriftgroesse_40, Font_Druckfarbe)
                '--------------------------------------------------
                Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Lueftungsart_Fensterlueftung.Wert, 1515, 1544, 2373, 2403)
                Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Lueftungsart_Schachtlueftung.Wert, 1515, 1544, 2450, 2480)
                Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Lueftungsart_Anlage_m_WRG.Wert, 2517, 2547, 2373, 2403)
                Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Lueftungsart_Anlage_o_WRG.Wert, 2517, 2547, 2450, 2480)
                '-------------------------------------------------------            
                Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Kuehlungsart_passive_Kuehlung.Wert, 1515, 1544, 2574, 2605)
                Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Kuehlungsart_gelieferte_Kaelte.Wert, 1515, 1544, 2652, 2683)
                Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Kuehlungsart_Strom.Wert, 2517, 2547, 2574, 2605)
                Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Kuehlungsart_Waerme.Wert, 2517, 2547, 2652, 2683)
                '-------------------------------------------------------
                If .Energieausweis_Daten.Keine_inspektionspflichtige_Anlage.Wert = False And .Energieausweis_Daten.Anzahl_Klimanlagen.Wert > 0 Then
                    Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Energieausweis_Daten.Anzahl_Klimanlagen.Wert, 0, "Stck.", 1683 + 20, 2115, 2729 + 15, 2826, Font_Schriftgroesse_40, Font_Druckfarbe)
                    Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, FormatDateTime(.Energieausweis_Daten.Faelligkeitsdatum_Inspektion.Wert, DateFormat.ShortDate), 3206 + 20, 4031, 2729 + 15, 2826, Font_Schriftgroesse_40, Font_Druckfarbe)
                End If
                '-------------------------------------------------------
                Select Case .Energieausweis_Daten.Ausstellungsanlass.Wert
                    Case "Neubau"
                        Image_Auswahl_schreiben(Grafik_Seite, True, 1515, 1544, 2884, 2915)
                        Image_Auswahl_schreiben(Grafik_Seite, False, 2441, 2472, 2884, 2915)
                        Image_Auswahl_schreiben(Grafik_Seite, False, 1515, 1544, 2962, 2993)
                        Image_Auswahl_schreiben(Grafik_Seite, False, 3278, 3309, 2962, 2993)
                        Image_Auswahl_schreiben(Grafik_Seite, False, 3278, 3309, 2884, 2915)
                    Case "Modernisierung-Erweiterung"
                        Image_Auswahl_schreiben(Grafik_Seite, False, 1515, 1544, 2884, 2915)
                        Image_Auswahl_schreiben(Grafik_Seite, True, 2441, 2472, 2884, 2915)
                        Image_Auswahl_schreiben(Grafik_Seite, False, 1515, 1544, 2962, 2993)
                        Image_Auswahl_schreiben(Grafik_Seite, False, 3278, 3309, 2962, 2993)
                        Image_Auswahl_schreiben(Grafik_Seite, False, 3278, 3309, 2884, 2915)
                    Case "Vermietung-Verkauf"
                        Image_Auswahl_schreiben(Grafik_Seite, False, 1515, 1544, 2884, 2915)
                        Image_Auswahl_schreiben(Grafik_Seite, False, 2441, 2472, 2884, 2915)
                        Image_Auswahl_schreiben(Grafik_Seite, True, 1515, 1544, 2962, 2993)
                        Image_Auswahl_schreiben(Grafik_Seite, False, 3278, 3309, 2962, 2993)
                        Image_Auswahl_schreiben(Grafik_Seite, False, 3278, 3309, 2884, 2915)
                    Case "Sonstiges"
                        Image_Auswahl_schreiben(Grafik_Seite, False, 1515, 1544, 2884, 2915)
                        Image_Auswahl_schreiben(Grafik_Seite, False, 2441, 2472, 2884, 2915)
                        Image_Auswahl_schreiben(Grafik_Seite, False, 1515, 1544, 2962, 2993)
                        Image_Auswahl_schreiben(Grafik_Seite, True, 3278, 3309, 2962, 2993)
                        Image_Auswahl_schreiben(Grafik_Seite, False, 3278, 3309, 2884, 2915)
                    Case "Aushangpflicht"
                        Image_Auswahl_schreiben(Grafik_Seite, False, 1515, 1544, 2884, 2915)
                        Image_Auswahl_schreiben(Grafik_Seite, False, 2441, 2472, 2884, 2915)
                        Image_Auswahl_schreiben(Grafik_Seite, False, 1515, 1544, 2962, 2993)
                        Image_Auswahl_schreiben(Grafik_Seite, False, 3278, 3309, 2962, 2993)
                        Image_Auswahl_schreiben(Grafik_Seite, True, 3278, 3309, 2884, 2915)
                End Select
                '-------------------------------------------------------
                Select Case Variable_Steuerung.Berechnungsart
                    Case "Energiebedarfsausweis"
                        Image_Auswahl_schreiben(Grafik_Seite, True, 293, 324, 3598, 3628)
                        Image_Auswahl_schreiben(Grafik_Seite, False, 293, 324, 3894, 3925)
                    Case "Energieverbrauchsausweis"
                        Image_Auswahl_schreiben(Grafik_Seite, False, 293, 324, 3598, 3628)
                        Image_Auswahl_schreiben(Grafik_Seite, True, 293, 324, 3894, 3925)
                End Select
                '-------------------------------------------------------
                Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Datenerhebung_Eigentuemer.Wert, 2101, 2131, 4098, 4129)
                Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Datenerhebung_Aussteller.Wert, 3161, 3191, 4098, 4129)
                '-------------------------------------------------------
                Image_Auswahl_schreiben(Grafik_Seite, .Gebaeudebezogene_Daten.Zusatzinfos_beigefuegt.Wert, 293, 324, 4209, 4239)
                '-------------------------------------------------------
                Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, FormatDateTime(Variable_Steuerung.Ausstellungsdatum, DateFormat.ShortDate).ToString, 3549, 3983, 5230, 5324, Font_Schriftgroesse_50, Font_Druckfarbe)
                '-------------------------------------------------------
                Image_Grafik_Projekt_schreiben(Grafik_Seite, Variable_Steuerung.Bild_Projekt, 3253, 1067, 4030 - 3253, 1935 - 1067)
                '-------------------------------------------------------
                Image_Grafik_Aussteller_schreiben(Grafik_Seite, Variable_Steuerung.Bild_Aussteller, .Gebaeudebezogene_Daten.Aussteller_Bezeichnung.Wert, .Gebaeudebezogene_Daten.Ausstellervorname.Wert, .Gebaeudebezogene_Daten.Ausstellername.Wert, .Gebaeudebezogene_Daten.Aussteller_Strasse_Nr.Wert, .Gebaeudebezogene_Daten.Aussteller_PLZ.Wert, .Gebaeudebezogene_Daten.Aussteller_Ort.Wert, 277, 4925, 2500, 300)
                '-------------------------------------------------------
                Image_Grafik_Unterschrift_schreiben(Grafik_Seite, Variable_Steuerung.Bild_Unterschrift, 3014, 4925, 1000, 300)
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
            Hintergrundgrafik = Image.FromFile(Variable_Parameter.Arbeitsverzeichnis & "/Image/GEG2024_NWB_Seite_2.png")
            Grafik_Seite.DrawImage(Hintergrundgrafik, 0, 0, 4200, 5940)
            '-------------------------------------------------------
            With Variable_XML_Import
                Select Case Variable_Steuerung.Berechnungsart
                    Case "Energiebedarfsausweis"
                        '-------------------------------------------------------
                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, Datum_Energiegesetz, 2129, 2617, 422, 517, Font_Schriftgroesse_50, Font_Druckfarbe)
                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.Mitte_kein_Abschneiden, Variable_Steuerung.Registriernummer, 3032, 3748, 718, 818, Font_Schriftgroesse_50, Font_Druckfarbe)
                        '-------------------------------------------------------
                        Select Case .Berechnungsverfahren
                            Case "Bedarfsberechnung-18599"
                                '-------------------------------------------------------
                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Treibhausgasemissionen.Wert, 2, "", 2649, 3053, 1109, 1212, Font_Schriftgroesse_50, Font_Druckfarbe)
                                '-------------------------------------------------------
                                Select Case .Energieausweis_Daten.Ausstellungsanlass.Wert
                                    Case "Neubau", "Modernisierung-Erweiterung", "Aushangpflicht"
                                        '-------------------------------------------------------
                                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Nichtwohngebaeude.Bedarfswerte_NWG.Primaerenergiebedarf_NGF.Wert, 1, "", 516, 759, 2247, 2339, Font_Schriftgroesse_40, Font_Druckfarbe)
                                        '-------------------------------------------------------
                                        Select Case Variable_Steuerung.Neubau
                                            Case 0 'Modernisierung
                                                '-------------------------------------------------------
                                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Nichtwohngebaeude.Bedarfswerte_NWG.Primaerenergiebedarf_Hoechstwert_Bestand.Wert, 1, "", 1661, 1904, 2247, 2339, Font_Schriftgroesse_40, Font_Druckfarbe)
                                                '-------------------------------------------------------
                                            Case 1 'Neubau
                                                '-------------------------------------------------------
                                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Nichtwohngebaeude.Bedarfswerte_NWG.Primaerenergiebedarf_Hoechstwert_Neubau.Wert, 1, "", 1661, 1904, 2247, 2339, Font_Schriftgroesse_40, Font_Druckfarbe)
                                                '-------------------------------------------------------
                                        End Select
                                        '-------------------------------------------------------
                                        Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Nichtwohngebaeude.Bedarfswerte_NWG.mittlere_Waermedurchgangskoeffizienten.Wert, 1388, 1429, 2383, 2425)
                                        Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Nichtwohngebaeude.Bedarfswerte_NWG.Sommerlicher_Waermeschutz.Wert, 1388, 1429, 2468, 2510)
                                        '-------------------------------------------------------
                                    Case "Vermietung-Verkauf", "Sonstiges"
                                        '-------------------------------------------------------
                                End Select
                                '-------------------------------------------------------
                                Select Case .Energieausweis_Daten.Nichtwohngebaeude.Bedarfswerte_NWG.Ein_Zonen_Modell.Wert
                                    Case False 'Mehr-Zonen-Modell
                                        Image_Auswahl_schreiben(Grafik_Seite, True, 2348, 2390, 2209, 2250)
                                        Image_Auswahl_schreiben(Grafik_Seite, False, 2348, 2390, 2293, 2335)
                                    Case True 'Ein-Zonen-Modell
                                        Image_Auswahl_schreiben(Grafik_Seite, False, 2348, 2390, 2209, 2250)
                                        Image_Auswahl_schreiben(Grafik_Seite, True, 2348, 2390, 2293, 2335)
                                End Select
                                '-------------------------------------------------------
                                Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Nichtwohngebaeude.Bedarfswerte_NWG.Vereinfachte_Datenaufnahme.Wert, 2348, 2390, 2377, 2419) 'Vereinfachung nach §50 Absatz 4
                                '-------------------------------------------------------
                                Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Nichtwohngebaeude.Bedarfswerte_NWG.Vereinfachungen_18599_1_D.Wert, 2348, 2390, 2462, 2504) 'Vereinfachung nach §21 Absatz 2 Satz 2
                                '-------------------------------------------------------
                                Image_Pfeil_NWB(Grafik_Seite, Lage.Oben, .Energieausweis_Daten.Nichtwohngebaeude.Bedarfswerte_NWG.Primaerenergiebedarf_NGF.Wert, 1, .Energieausweis_Daten.Nichtwohngebaeude.Bedarfswerte_NWG.Primaerenergiebedarf_Hoechstwert_Bestand.Wert, 3, "Primärenergiebedarf dieses Gebäudes", "", 552, 3633, 1493, 1844)
                                '-------------------------------------------------------
                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Nichtwohngebaeude.Bedarfswerte_NWG.Endenergiebedarf_Waerme_NGF.Wert, 1, "", 3019, 3660, 3264, 3379, Font_Schriftgroesse_60, Font_Druckfarbe)
                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Nichtwohngebaeude.Bedarfswerte_NWG.Endenergiebedarf_Strom_NGF.Wert, 1, "", 3019, 3660, 3385, 3487, Font_Schriftgroesse_60, Font_Druckfarbe)
                                '-------------------------------------------------------
                                'Tabelle Energiebedarf ---------------------------------
                                '-------------------------------------------------------
                                .Energiebedarf_Anzahl = Anzahl_Energiebedarf()
                                '-------------------------------------------------------
                                Zeilenzahl = 2
                                Zeilenhoehe = (3156 - 2955) / 2
                                '-------------------------------------------------------
                                If .Energiebedarf_Anzahl < 2 Then
                                    Zeilenzahl = .Energiebedarf_Anzahl
                                End If
                                '-------------------------------------------------------
                                For Anzahl_Zeilen = 1 To Zeilenzahl
                                    '-------------------------------------------------------
                                    Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Energietraeger(Anzahl_Zeilen).Energietraegerbezeichnung.Wert, 245, 1761, 2955 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 2955 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                    '--------------------------------------------------
                                    If .Energietraeger(Anzahl_Zeilen).Endenergiebedarf_Heizung_spezifisch.Wert <> 0 Then
                                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energietraeger(Anzahl_Zeilen).Endenergiebedarf_Heizung_spezifisch.Wert, 1, "", 1765, 2054, 2955 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 2955 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                    End If
                                    '--------------------------------------------------
                                    If .Energietraeger(Anzahl_Zeilen).Endenergiebedarf_Trinkwarmwasser_spezifisch.Wert <> 0 Then
                                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energietraeger(Anzahl_Zeilen).Endenergiebedarf_Trinkwarmwasser_spezifisch.Wert, 1, "", 2058, 2456, 2955 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 2955 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                    End If
                                    '--------------------------------------------------
                                    If .Energietraeger(Anzahl_Zeilen).Endenergiebedarf_Beleuchtung_spezifisch.Wert <> 0 Then
                                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energietraeger(Anzahl_Zeilen).Endenergiebedarf_Beleuchtung_spezifisch.Wert, 1, "", 2460, 2859, 2955 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 2955 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                    End If
                                    '--------------------------------------------------
                                    If .Energietraeger(Anzahl_Zeilen).Endenergiebedarf_Lueftung_spezifisch.Wert <> 0 Then
                                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energietraeger(Anzahl_Zeilen).Endenergiebedarf_Lueftung_spezifisch.Wert, 1, "", 2863, 3184, 2955 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 2955 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                    End If
                                    '--------------------------------------------------
                                    If .Energietraeger(Anzahl_Zeilen).Endenergiebedarf_Kuehlung_Befeuchtung_spezifisch.Wert <> 0 Then
                                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energietraeger(Anzahl_Zeilen).Endenergiebedarf_Kuehlung_Befeuchtung_spezifisch.Wert, 1, "", 3188, 3652, 2955 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 2955 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                    End If
                                    '--------------------------------------------------
                                    If .Energietraeger(Anzahl_Zeilen).Endenergiebedarf_Energietraeger_Gesamtgebaeude_spezifisch.Wert <> 0 Then
                                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energietraeger(Anzahl_Zeilen).Endenergiebedarf_Energietraeger_Gesamtgebaeude_spezifisch.Wert, 1, "", 3656, 4032, 2955 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 2955 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                    End If
                                    '--------------------------------------------------
                                Next
                                '-------------------------------------------------------
                                For Anzahl_Zeilen = 1 To 2
                                    '-------------------------------------------------------
                                    Image_Linie(Grafik_Seite, 245, 4030, 2955 + (Zeilenhoehe * Anzahl_Zeilen), 2955 + (Zeilenhoehe * Anzahl_Zeilen))
                                    '-------------------------------------------------------
                                Next
                                '-------------------------------------------------------
                                If .Energiebedarf_Anzahl > 2 Then
                                    '-------------------------------------------------------
                                    Image_Auswahl_schreiben(Grafik_Seite, True, 294, 335, 3176, 3217)
                                    '-------------------------------------------------------
                                End If
                                '-------------------------------------------------------
                                'Tabelle Zonen -----------------------------------------
                                '-------------------------------------------------------
                                .Zonen_Anzahl = Anzahl_Zonen()
                                '-------------------------------------------------------
                                Zeilenzahl = 5
                                Zeilenhoehe = (4315 - 3816) / 5
                                '-------------------------------------------------------
                                If .Zonen_Anzahl < 5 Then
                                    Zeilenzahl = .Zonen_Anzahl
                                End If
                                '-------------------------------------------------------
                                For Anzahl_Zeilen = 1 To Zeilenzahl
                                    '-------------------------------------------------------
                                    Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Energieausweis_Daten.Nichtwohngebaeude.Zone(Anzahl_Zeilen).Zonenbezeichnung.Wert, 2303, 3318, 3816 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 3816 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                    '-------------------------------------------------------
                                    If .Energieausweis_Daten.Nichtwohngebaeude.Zone(Anzahl_Zeilen).Nettogrundflaeche_Zone.Wert <> 0 Then
                                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Nichtwohngebaeude.Zone(Anzahl_Zeilen).Nettogrundflaeche_Zone.Wert, 1, "", 3322, 3669, 3816 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 3816 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                    End If
                                    '-------------------------------------------------------
                                    If .Energieausweis_Daten.Nichtwohngebaeude.Zone(Anzahl_Zeilen).Nettogrundflaeche_Zone.Wert <> 0 Then
                                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, (.Energieausweis_Daten.Nichtwohngebaeude.Zone(Anzahl_Zeilen).Nettogrundflaeche_Zone.Wert / .Energieausweis_Daten.Nichtwohngebaeude.Nettogrundflaeche.Wert * 100), 2, "", 3679, 4032, 3816 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 3816 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                    End If
                                    '-------------------------------------------------------
                                Next
                                '-------------------------------------------------------
                                For Anzahl_Zeilen = 1 To 5
                                    '-------------------------------------------------------
                                    Image_Linie(Grafik_Seite, 2189, 4030, 3816 + (Zeilenhoehe * Anzahl_Zeilen), 3816 + (Zeilenhoehe * Anzahl_Zeilen))
                                    '-------------------------------------------------------
                                Next
                                '-------------------------------------------------------
                                If .Zonen_Anzahl > 5 Then
                                    '-------------------------------------------------------
                                    Image_Auswahl_schreiben(Grafik_Seite, True, 2252, 2294, 4346, 4387)
                                    '-------------------------------------------------------
                                End If
                                '-------------------------------------------------------
                                'Erneuerbare Energien ----------------------------------
                                '-------------------------------------------------------
                                If .Energieausweis_Daten.Nutzung_zur_Erfuellung_von_EE_neue_Anlage.Wert = True Then
                                    '-------------------------------------------------------
                                    Select Case .Energieausweis_Daten.Keine_Pauschale_Erfuellungsoptionen_Anlagentyp.Wert
                                        Case True
                                            '-------------------------------------------------------
                                            Image_Auswahl_schreiben(Grafik_Seite, False, 351, 390, 3993, 4034)
                                            Image_Auswahl_schreiben(Grafik_Seite, True, 341, 379, 4612, 4655)
                                            '-------------------------------------------------------
                                            Image_Auswahl_schreiben(Grafik_Seite, False, 438, 476, 4136, 4173)
                                            Image_Auswahl_schreiben(Grafik_Seite, False, 438, 476, 4189, 4225)
                                            Image_Auswahl_schreiben(Grafik_Seite, False, 438, 476, 4245, 4280)
                                            Image_Auswahl_schreiben(Grafik_Seite, False, 438, 476, 4298, 4334)
                                            Image_Auswahl_schreiben(Grafik_Seite, False, 438, 476, 4348, 4384)
                                            Image_Auswahl_schreiben(Grafik_Seite, False, 438, 476, 4403, 4439)
                                            Image_Auswahl_schreiben(Grafik_Seite, False, 438, 476, 4452, 4488)
                                            Image_Auswahl_schreiben(Grafik_Seite, False, 438, 476, 4508, 4543)
                                            '-------------------------------------------------------
                                        Case Else
                                            '-------------------------------------------------------
                                            Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Nutzung_zur_Erfuellung_von_EE_neue_Anlage.Wert, 272, 314, 3847, 3888)
                                            '-------------------------------------------------------
                                            Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.EE_Angabe_Heizung.Wert, 1219, 1264, 3767, 3814)
                                            Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.EE_Angabe_Warmwasser.Wert, 1594, 1641, 3767, 3814)
                                            '-------------------------------------------------------
                                            Image_Auswahl_schreiben(Grafik_Seite, True, 351, 390, 3993, 4034)
                                            Image_Auswahl_schreiben(Grafik_Seite, False, 341, 379, 4612, 4655)
                                            '-------------------------------------------------------
                                            Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Hausuebergabestation.Wert, 438, 476, 4136, 4173)
                                            Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Waermepumpe.Wert, 438, 476, 4189, 4225)
                                            Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Stromdirektheizung.Wert, 438, 476, 4245, 4280)
                                            Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Solarthermische_Anlage.Wert, 438, 476, 4298, 4334)
                                            Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Heizungsanlage_Biomasse_Wasserstoff_Wasserstoffderivale.Wert, 438, 476, 4348, 4384)
                                            Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Waermepumpen_Hybridheizung.Wert, 438, 476, 4403, 4439)
                                            Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Solarthermie_Hybridheizung.Wert, 438, 476, 4452, 4488)
                                            Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Dezentral_elektrische_Warmwasserbereitung.Wert, 438, 476, 4508, 4543)
                                            '-------------------------------------------------------
                                    End Select
                                    '-------------------------------------------------------
                                End If
                                '-------------------------------------------------------
                                If .Energieausweis_Daten.Keine_Pauschale_Erfuellungsoptionen_Anlagentyp.Wert = False Then
                                    '-------------------------------------------------------
                                    Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Nutzung_bei_Bestandsanlagen.Wert, 269, 308, 5141, 5181)
                                    '-------------------------------------------------------
                                    'Tabelle Erneuerbare-Energien-65EE-Regel ---------------
                                    '-------------------------------------------------------
                                    .Erneuerbare_Energien_65EE_Regel_Anzahl = Anzahl_Erneuerbare_Energien_65EE_Regel()
                                    ''-------------------------------------------------------
                                    Zeilenzahl = 2
                                    Zeilenhoehe = 90
                                    '-------------------------------------------------------
                                    If .Erneuerbare_Energien_65EE_Regel_Anzahl < 2 Then
                                        Zeilenzahl = .Erneuerbare_Energien_65EE_Regel_Anzahl
                                    End If
                                    '-------------------------------------------------------
                                    For Anzahl_Zeilen = 1 To Zeilenzahl
                                        '-------------------------------------------------------
                                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .EE_65EE_Regel(Anzahl_Zeilen).Art_der_Nutzung_erneuerbaren_Energie.Wert, 281, 1297, 4842 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 4842 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                        '-------------------------------------------------------
                                        If .EE_65EE_Regel(Anzahl_Zeilen).Deckungsanteil.Wert <> 0 Then
                                            Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .EE_65EE_Regel(Anzahl_Zeilen).Deckungsanteil.Wert, 0, "", 1316, 1460, 4842 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 4842 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                        End If
                                        '-------------------------------------------------------
                                        If .EE_65EE_Regel(Anzahl_Zeilen).Anteil_der_Pflichterfuellung_Anlage.Wert <> 0 Then
                                            Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .EE_65EE_Regel(Anzahl_Zeilen).Anteil_der_Pflichterfuellung_Anlage.Wert, 0, "", 1577, 1735, 4842 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 4842 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                        End If
                                        '-------------------------------------------------------
                                        If .EE_65EE_Regel(Anzahl_Zeilen).Anteil_der_Pflichterfuellung_Gesamt.Wert <> 0 Then
                                            Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .EE_65EE_Regel(Anzahl_Zeilen).Anteil_der_Pflichterfuellung_Gesamt.Wert, 0, "", 1837, 1998, 4842 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 4842 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                        End If
                                        '-------------------------------------------------------
                                    Next
                                    '-------------------------------------------------------
                                    If Gesamtsumme_Erneuerbare_Energien_65EE_Regel() <> 0 Then
                                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, Gesamtsumme_Erneuerbare_Energien_65EE_Regel(), 0, "", 1837, 1998, 5027, 5098, Font_Schriftgroesse_40, Font_Druckfarbe)
                                    End If
                                    '-------------------------------------------------------
                                    'Tabelle Erneuerbare-Energien-65EE-keine-Regel ---------
                                    '-------------------------------------------------------
                                    .Erneuerbare_Energien_65EE_keine_Regel_Anzahl = Anzahl_Erneuerbare_Energien_65EE_keine_Regel()
                                    ''-------------------------------------------------------
                                    Zeilenzahl = 2
                                    Zeilenhoehe = 92
                                    '-------------------------------------------------------
                                    If .Erneuerbare_Energien_65EE_keine_Regel_Anzahl < 2 Then
                                        Zeilenzahl = .Erneuerbare_Energien_65EE_keine_Regel_Anzahl
                                    End If
                                    '-------------------------------------------------------
                                    For Anzahl_Zeilen = 1 To Zeilenzahl
                                        '-------------------------------------------------------
                                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .EE_65EE_keine_Regel(Anzahl_Zeilen).Art_der_Nutzung_erneuerbaren_Energie.Wert, 274, 1808, 5265 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 5265 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                        '-------------------------------------------------------
                                        If .EE_65EE_keine_Regel(Anzahl_Zeilen).Anteil_EE_Anlage.Wert <> 0 Then
                                            Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .EE_65EE_keine_Regel(Anzahl_Zeilen).Anteil_EE_Anlage.Wert, 0, "", 1837, 1998, 5265 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 5265 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                        End If
                                        '-------------------------------------------------------
                                    Next
                                    '-------------------------------------------------------
                                    If Gesamtsumme_Erneuerbare_Energien_65EE_keine_Regel() <> 0 Then
                                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, Gesamtsumme_Erneuerbare_Energien_65EE_keine_Regel(), 0, "", 1837, 1998, 5451, 5520, Font_Schriftgroesse_40, Font_Druckfarbe)
                                    End If
                                    '-------------------------------------------------------
                                    If .Erneuerbare_Energien_65EE_Regel_Anzahl > 2 Or .Erneuerbare_Energien_65EE_keine_Regel_Anzahl > 2 Or .Energieausweis_Daten.Nichtwohngebaeude.Bedarfswerte_NWG.Weitere_Eintraege_und_Erlaeuterungen_in_der_Anlage.Wert = True Then
                                        '-------------------------------------------------------
                                        Image_Auswahl_schreiben(Grafik_Seite, True, 270, 313, 5534, 5573)
                                        '-------------------------------------------------------
                                    End If
                                    '-------------------------------------------------------
                                End If
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
            Hintergrundgrafik = Image.FromFile(Variable_Parameter.Arbeitsverzeichnis & "/Image/GEG2024_NWB_Seite_3.png")
            Grafik_Seite.DrawImage(Hintergrundgrafik, 0, 0, 4200, 5940)
            '-------------------------------------------------------
            With Variable_XML_Import
                '-------------------------------------------------------
                If Variable_Steuerung.Berechnungsart = "Energieverbrauchsausweis" Or Variable_XML_Import.Zusaetzliche_Verbrauchsberechnung = True Then
                    '-------------------------------------------------------
                    Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, Datum_Energiegesetz, 2129, 2617, 422, 517, Font_Schriftgroesse_50, Font_Druckfarbe)
                    Image_Text_schreiben(Grafik_Seite, Ausrichtung.Mitte_kein_Abschneiden, Variable_Steuerung.Registriernummer, 3032, 3748, 718, 818, Font_Schriftgroesse_50, Font_Druckfarbe)
                    '-------------------------------------------------------
                    If .Berechnungsverfahren = "Verbrauchsberechnung" Or Variable_XML_Import.Zusaetzliche_Verbrauchsberechnung = True Then
                        '-------------------------------------------------------
                        Image_Pfeil_NWB(Grafik_Seite, Lage.Oben, .Energieausweis_Daten.Nichtwohngebaeude.Verbrauchswerte_NWG.Endenergieverbrauch_Waerme.Wert, 1, .Energieausweis_Daten.Nichtwohngebaeude.Verbrauchswerte_NWG.Endenergieverbrauch_Waerme_Vergleichswert.Wert, 2, "Endenergieverbrauch Wärme", "[Pflichtangabe in Immobilienanzeigen]", 616, 3697, 1510, 1847)
                        '-------------------------------------------------------
                        Image_Pfeil_NWB(Grafik_Seite, Lage.Oben, .Energieausweis_Daten.Nichtwohngebaeude.Verbrauchswerte_NWG.Endenergieverbrauch_Strom.Wert, 1, .Energieausweis_Daten.Nichtwohngebaeude.Verbrauchswerte_NWG.Endenergieverbrauch_Strom_Vergleichswert.Wert, 2, "Endenergieverbrauch Strom", "[Pflichtangabe in Immobilienanzeigen]", 616, 3697, 2511, 2846)
                        '-------------------------------------------------------
                        Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Nichtwohngebaeude.Verbrauchswerte_NWG.Warmwasser_enthalten.Wert, 295, 334, 1914, 1951)
                        Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Nichtwohngebaeude.Verbrauchswerte_NWG.Kuehlung_enthalten.Wert, 295, 334, 1998, 2036)
                        '-------------------------------------------------------
                        Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Nichtwohngebaeude.Verbrauchswerte_NWG.Strom_Daten.Stromverbrauch_enthaelt_Zusatzheizung.Wert, 299, 338, 3234, 3272)
                        Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Nichtwohngebaeude.Verbrauchswerte_NWG.Strom_Daten.Stromverbrauch_enthaelt_Warmwasser.Wert, 932, 972, 3234, 3272)
                        Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Nichtwohngebaeude.Verbrauchswerte_NWG.Strom_Daten.Stromverbrauch_enthaelt_Lueftung.Wert, 1546, 1586, 3234, 3272)
                        Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Nichtwohngebaeude.Verbrauchswerte_NWG.Strom_Daten.Stromverbrauch_enthaelt_Beleuchtung.Wert, 1990, 2029, 3234, 3272)
                        Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Nichtwohngebaeude.Verbrauchswerte_NWG.Strom_Daten.Stromverbrauch_enthaelt_Kuehlung.Wert, 2984, 3023, 3234, 3272)
                        Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Nichtwohngebaeude.Verbrauchswerte_NWG.Strom_Daten.Stromverbrauch_enthaelt_Sonstiges.Wert, 3475, 3514, 3234, 3272)
                        '-------------------------------------------------------
                        'Tabelle Verbrauchserfassung ---------------------------
                        '-------------------------------------------------------
                        Zusammenstellung_NWB_Verbrauchserfassung()
                        '-------------------------------------------------------
                        Zeilenzahl = 3
                        Zeilenhoehe = (4079 - 3780) / 3
                        '-------------------------------------------------------
                        If .Verbrauchserfassung_Anzahl < 3 Then
                            Zeilenzahl = .Verbrauchserfassung_Anzahl
                        End If
                        '-------------------------------------------------------
                        For Anzahl_Zeilen = 1 To Zeilenzahl
                            '-------------------------------------------------------
                            Image_Datum_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Verbrauchserfassung(Anzahl_Zeilen).Startdatum.Wert, 249, 573, 3780 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 3780 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                            '--------------------------------------------------
                            Image_Datum_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Verbrauchserfassung(Anzahl_Zeilen).Enddatum.Wert, 577, 902, 3780 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 3780 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                            '--------------------------------------------------
                            Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Verbrauchserfassung(Anzahl_Zeilen).Energietraeger_Verbrauch.Wert, 906, 1794, 3780 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 3780 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                            '--------------------------------------------------
                            If .Verbrauchserfassung(Anzahl_Zeilen).Primaerenergiefaktor.Wert <> 0 Then
                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Verbrauchserfassung(Anzahl_Zeilen).Primaerenergiefaktor.Wert, 1, "", 1798, 2034, 3780 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 3780 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                            End If
                            '--------------------------------------------------
                            If .Verbrauchserfassung(Anzahl_Zeilen).Energieverbrauch.Wert <> 0 Then
                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Verbrauchserfassung(Anzahl_Zeilen).Energieverbrauch.Wert, 0, "", 2038, 2394, 3780 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 3780 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                            End If
                            '--------------------------------------------------
                            If .Verbrauchserfassung(Anzahl_Zeilen).Energieverbrauchsanteil_Warmwasser_zentral.Wert <> 0 Then
                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Verbrauchserfassung(Anzahl_Zeilen).Energieverbrauchsanteil_Warmwasser_zentral.Wert, 0, "", 2398, 2754, 3780 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 3780 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                            End If
                            '--------------------------------------------------
                            If .Verbrauchserfassung(Anzahl_Zeilen).Energieverbrauchsanteil_thermisch_erzeugte_Kaelte.Wert <> 0 Then
                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Verbrauchserfassung(Anzahl_Zeilen).Energieverbrauchsanteil_thermisch_erzeugte_Kaelte.Wert, 0, "", 2758, 3113, 3780 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 3780 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                            End If
                            '--------------------------------------------------
                            If .Verbrauchserfassung(Anzahl_Zeilen).Energieverbrauchsanteil_Heizung.Wert <> 0 Then
                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Verbrauchserfassung(Anzahl_Zeilen).Energieverbrauchsanteil_Heizung.Wert, 0, "", 3118, 3474, 3780 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 3780 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                            End If
                            '--------------------------------------------------
                            If .Verbrauchserfassung(Anzahl_Zeilen).Klimafaktor.Wert <> 0 Then
                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Verbrauchserfassung(Anzahl_Zeilen).Klimafaktor.Wert, 2, "", 3478, 3673, 3780 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 3780 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                            End If
                            '-------------------------------------------------------
                            If .Verbrauchserfassung(Anzahl_Zeilen).Energieverbrauch_Strom.Wert <> 0 Then
                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Verbrauchserfassung(Anzahl_Zeilen).Energieverbrauch_Strom.Wert, 0, "", 3677, 4031, 3780 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 3780 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                            End If
                            '--------------------------------------------------
                        Next
                        '-------------------------------------------------------
                        For Anzahl_Zeilen = 1 To 3
                            '-------------------------------------------------------
                            Image_Linie(Grafik_Seite, 249, 4031, 3780 + (Zeilenhoehe * Anzahl_Zeilen), 3780 + (Zeilenhoehe * Anzahl_Zeilen))
                            '-------------------------------------------------------
                        Next
                        '-------------------------------------------------------
                        If .Verbrauchserfassung_Anzahl > 3 Then
                            '-------------------------------------------------------
                            Image_Auswahl_schreiben(Grafik_Seite, True, 300, 331, 4131, 4162)
                            '-------------------------------------------------------
                        End If
                        '-------------------------------------------------------
                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Nichtwohngebaeude.Verbrauchswerte_NWG.Primaerenergieverbrauch.Wert, 1, "", 3019, 3660, 4238, 4375, Font_Schriftgroesse_60, Font_Druckfarbe)
                        '-------------------------------------------------------
                        If Variable_XML_Import.Zusaetzliche_Verbrauchsberechnung = False Then
                            Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Treibhausgasemissionen.Wert, 2, "", 3019, 3660, 4379, 4503, Font_Schriftgroesse_60, Font_Druckfarbe)
                        Else
                            Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Treibhausgasemissionen_Zusaetzliche_Verbrauchsdaten.Wert, 2, "", 3019, 3660, 4379, 4503, Font_Schriftgroesse_60, Font_Druckfarbe)
                        End If
                        '-------------------------------------------------------
                        'Tabelle Gebäudenutzung --------------------------------
                        '-------------------------------------------------------
                        .Gebaeudenutzung_Anzahl = Anzahl_Gebaeudenutzung()
                        '-------------------------------------------------------
                        Zeilenzahl = 3
                        Zeilenhoehe = (5283 - 4955) / 3
                        '-------------------------------------------------------
                        If .Gebaeudenutzung_Anzahl < 3 Then
                            Zeilenzahl = .Gebaeudenutzung_Anzahl
                        End If
                        '-------------------------------------------------------
                        For Anzahl_Zeilen = 1 To Zeilenzahl
                            '-------------------------------------------------------
                            Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Nutzung_Gebaeudekategorie(Anzahl_Zeilen).Gebaeudekategorie.Wert, 249, 1262, 4955 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 4955 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                            '--------------------------------------------------
                            If .Nutzung_Gebaeudekategorie(Anzahl_Zeilen).Flaechenanteil_Nutzung.Wert <> 0 Then
                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Nutzung_Gebaeudekategorie(Anzahl_Zeilen).Flaechenanteil_Nutzung.Wert, 1, "", 1266, 1538, 4955 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 4955 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                            End If
                            '--------------------------------------------------
                            If .Nutzung_Gebaeudekategorie(Anzahl_Zeilen).Vergleichswert_Waerme.Wert <> 0 Then
                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Nutzung_Gebaeudekategorie(Anzahl_Zeilen).Vergleichswert_Waerme.Wert, 1, "", 1542, 1814, 4955 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 4955 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                            End If
                            '--------------------------------------------------
                            If .Nutzung_Gebaeudekategorie(Anzahl_Zeilen).Vergleichswert_Strom.Wert <> 0 Then
                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Nutzung_Gebaeudekategorie(Anzahl_Zeilen).Vergleichswert_Strom.Wert, 1, "", 1818, 2090, 4955 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 4955 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                            End If
                            '--------------------------------------------------
                        Next
                        '-------------------------------------------------------
                        For Anzahl_Zeilen = 1 To 3
                            '-------------------------------------------------------
                            Image_Linie(Grafik_Seite, 249, 2090, 4955 + (Zeilenhoehe * Anzahl_Zeilen), 4955 + (Zeilenhoehe * Anzahl_Zeilen))
                            '-------------------------------------------------------
                        Next
                        '-------------------------------------------------------
                        If .Gebaeudenutzung_Anzahl > 3 Then
                            '-------------------------------------------------------
                            Image_Auswahl_schreiben(Grafik_Seite, True, 300, 331, 5329, 5356)
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
            Hintergrundgrafik = Image.FromFile(Variable_Parameter.Arbeitsverzeichnis & "/Image/GEG2024_NWB_Seite_4.png")
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
                    Image_Auswahl_schreiben(Grafik_Seite, True, 2848, 2882, 1143, 1177)
                    '--------------------------------------------------
                    Image_Auswahl_schreiben(Grafik_Seite, False, 2440, 2474, 1143, 1177)
                    '--------------------------------------------------
                Else
                    '--------------------------------------------------
                    Image_Auswahl_schreiben(Grafik_Seite, False, 2848, 2882, 1143, 1177)
                    '--------------------------------------------------
                    Image_Auswahl_schreiben(Grafik_Seite, True, 2440, 2474, 1143, 1177)
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
                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Modernisierungsempfehlungen(Anzahl_Zeilen).Nummer.Wert, 0, "", 239, 404, 1926 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1926 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                        '-------------------------------------------------------
                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksMitte, .Modernisierungsempfehlungen(Anzahl_Zeilen).Bauteil_Anlagenteil.Wert, 408 + 20, 912, 1926 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1926 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                        '-------------------------------------------------------
                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksMitte, .Modernisierungsempfehlungen(Anzahl_Zeilen).Massnahmenbeschreibung.Wert, 916 + 20, 2572, 1926 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1926 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                        '-------------------------------------------------------
                        Select Case .Modernisierungsempfehlungen(Anzahl_Zeilen).Modernisierungskombination.Wert
                            Case "in Zusammenhang mit größerer Modernisierung"
                                Image_Auswahl_schreiben(Grafik_Seite, True, 2696, 2730, 1974 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 2008 + (Zeilenhoehe * (Anzahl_Zeilen - 1)))
                                Image_Auswahl_schreiben(Grafik_Seite, False, 2944, 2978, 1974 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 2008 + (Zeilenhoehe * (Anzahl_Zeilen - 1)))
                            Case "als Einzelmaßnahme"
                                Image_Auswahl_schreiben(Grafik_Seite, False, 2696, 2730, 1974 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 2008 + (Zeilenhoehe * (Anzahl_Zeilen - 1)))
                                Image_Auswahl_schreiben(Grafik_Seite, True, 2944, 2978, 1974 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 2008 + (Zeilenhoehe * (Anzahl_Zeilen - 1)))
                        End Select
                        '-------------------------------------------------------
                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksMitte, .Modernisierungsempfehlungen(Anzahl_Zeilen).Amortisation.Wert, 3072 + 20, 3400, 1926 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1926 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                        '-------------------------------------------------------
                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksMitte, .Modernisierungsempfehlungen(Anzahl_Zeilen).spezifische_Kosten.Wert, 3404 + 20, 4032, 1926 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1926 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                        '-------------------------------------------------------
                    Next
                    '--------------------------------------------------
                    If .Modernisierungsempfehlung_Anzahl > 5 Then
                        '-------------------------------------------------------
                        Image_Auswahl_schreiben(Grafik_Seite, True, 316, 347, 3433, 3463)
                        '-------------------------------------------------------
                    Else
                        '-------------------------------------------------------
                        Image_Auswahl_schreiben(Grafik_Seite, False, 317, 359, 3381, 3423)
                        '-------------------------------------------------------
                    End If
                    '--------------------------------------------------
                End If
                '--------------------------------------------------
                Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksMitte, .Gebaeudebezogene_Daten.Angaben_erhaeltlich.Wert, 1566 + 20, 4032, 3682, 3924, Font_Schriftgroesse_40, Font_Druckfarbe)
                '--------------------------------------------------
                Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksOben, .Gebaeudebezogene_Daten.Ergaenzdende_Erlaeuterungen.Wert, 244 + 20, 4032, 4195, 5599, Font_Schriftgroesse_40, Font_Druckfarbe)
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
            Hintergrundgrafik = Image.FromFile(Variable_Parameter.Arbeitsverzeichnis & "/Image/GEG2024_NWB_Seite_5.png")
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
#Region "Zusatzseite Energiebedarf"
            '-------------------------------------------------------
            Select Case Variable_Steuerung.Berechnungsart
                Case "Energiebedarfsausweis"
                    '-------------------------------------------------------
                    Select Case Variable_XML_Import.Berechnungsverfahren
                        Case "Bedarfsberechnung-18599"
                            '-------------------------------------------------------
                            Wert_Schleife_Offset = 2
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
                                        Hintergrundgrafik = Image.FromFile(Variable_Parameter.Arbeitsverzeichnis & "/Image/GEG2024_NWB_Ueberlauf_Bedarf.png")
                                        Grafik_Seite.DrawImage(Hintergrundgrafik, 0, 0, 4200, 5940)
                                        '-------------------------------------------------------
                                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, Seitennummer, 0, "", 3893, 4040, 688, 840, Font_Seitennummer, Font_Druckfarbe)
                                        '-------------------------------------------------------
                                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, Datum_Energiegesetz, 2128, 2616, 422, 516, Font_Schriftgroesse_50, Font_Druckfarbe)
                                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.Mitte_kein_Abschneiden, Variable_Steuerung.Registriernummer, 3049, 3764, 722, 820, Font_Schriftgroesse_50, Font_Druckfarbe)
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
                            Wert_Schleife_Offset = 5
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
                                        Hintergrundgrafik = Image.FromFile(Variable_Parameter.Arbeitsverzeichnis & "/Image/GEG2024_NWB_Ueberlauf_Gebaeudezonen.png")
                                        Grafik_Seite.DrawImage(Hintergrundgrafik, 0, 0, 4200, 5940)
                                        '-------------------------------------------------------
                                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, Seitennummer, 0, "", 3893, 4040, 688, 840, Font_Seitennummer, Font_Druckfarbe)
                                        '-------------------------------------------------------
                                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, Datum_Energiegesetz, 2129, 2617, 422, 517, Font_Schriftgroesse_50, Font_Druckfarbe)
                                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.Mitte_kein_Abschneiden, Variable_Steuerung.Registriernummer, 3032, 3748, 718, 818, Font_Schriftgroesse_50, Font_Druckfarbe)
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
                                Hintergrundgrafik = Image.FromFile(Variable_Parameter.Arbeitsverzeichnis & "/Image/GEG2024_NWB_Ueberlauf_Verbrauch.png")
                                Grafik_Seite.DrawImage(Hintergrundgrafik, 0, 0, 4200, 5940)
                                '-------------------------------------------------------
                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, Seitennummer, 0, "", 3893, 4040, 688, 840, Font_Seitennummer, Font_Druckfarbe)
                                '-------------------------------------------------------
                                Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, Datum_Energiegesetz, 2128, 2616, 422, 516, Font_Schriftgroesse_50, Font_Druckfarbe)
                                Image_Text_schreiben(Grafik_Seite, Ausrichtung.Mitte_kein_Abschneiden, Variable_Steuerung.Registriernummer, 3049, 3764, 722, 820, Font_Schriftgroesse_50, Font_Druckfarbe)
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
#Region "Gebäudenutzung"
            '-------------------------------------------------------
            Select Case Variable_Steuerung.Berechnungsart
                        '-------------------------------------------------------
                Case "Energieverbrauchsausweis"
                    '-------------------------------------------------------
                    Select Case Variable_XML_Import.Berechnungsverfahren
                        Case "Verbrauchsberechnung"
                            '-------------------------------------------------------
                            If Variable_XML_Import.Gebaeudenutzung_Anzahl > 3 Then
                                '-------------------------------------------------------
                                Seitennummer += 1
                                '-------------------------------------------------------
                                Grafik_Seite.Clear(Color.White)
                                '-------------------------------------------------------
                                Hintergrundgrafik = Image.FromFile(Variable_Parameter.Arbeitsverzeichnis & "/Image/GEG2024_NWB_Ueberlauf_Gebaeudenutzung.png")
                                Grafik_Seite.DrawImage(Hintergrundgrafik, 0, 0, 4200, 5940)
                                '-------------------------------------------------------
                                With Variable_XML_Import
                                    '-------------------------------------------------------
                                    Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, Seitennummer, 0, "", 3893, 4040, 688, 840, Font_Seitennummer, Font_Druckfarbe)
                                    '-------------------------------------------------------
                                    Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, Datum_Energiegesetz, 2128, 2616, 422, 516, Font_Schriftgroesse_50, Font_Druckfarbe)
                                    Image_Text_schreiben(Grafik_Seite, Ausrichtung.Mitte_kein_Abschneiden, Variable_Steuerung.Registriernummer, 3049, 3764, 722, 820, Font_Schriftgroesse_50, Font_Druckfarbe)
                                    '-------------------------------------------------------
                                    Zeilenzahl = 2
                                    Zeilenhoehe = (5300 - 1299) / 50
                                    '-------------------------------------------------------
                                    If .Gebaeudenutzung_Anzahl < 5 Then
                                        Zeilenzahl = .Gebaeudenutzung_Anzahl - 3
                                    End If
                                    '-------------------------------------------------------
                                    For Anzahl_Zeilen = 1 To Zeilenzahl
                                        '-------------------------------------------------------
                                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Nutzung_Gebaeudekategorie(Anzahl_Zeilen).Gebaeudekategorie.Wert, 249, 2881, 1299 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1299 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                        '--------------------------------------------------
                                        If .Nutzung_Gebaeudekategorie(Anzahl_Zeilen + 3).Flaechenanteil_Nutzung.Wert <> 0 Then
                                            Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Nutzung_Gebaeudekategorie(Anzahl_Zeilen + 3).Flaechenanteil_Nutzung.Wert, 1, "", 2884, 3232, 1299 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1299 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                        End If
                                        '--------------------------------------------------
                                        If .Nutzung_Gebaeudekategorie(Anzahl_Zeilen + 3).Vergleichswert_Waerme.Wert <> 0 Then
                                            Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Nutzung_Gebaeudekategorie(Anzahl_Zeilen + 3).Vergleichswert_Waerme.Wert, 1, "", 3226, 3621, 1299 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1299 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                        End If
                                        '--------------------------------------------------
                                        If .Nutzung_Gebaeudekategorie(Anzahl_Zeilen + 3).Vergleichswert_Strom.Wert <> 0 Then
                                            Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Nutzung_Gebaeudekategorie(Anzahl_Zeilen + 3).Vergleichswert_Strom.Wert, 1, "", 3625, 4033, 1299 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1299 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                        End If
                                        '-------------------------------------------------------
                                    Next
                                    '-------------------------------------------------------
                                    For Anzahl_Zeilen = 1 To 50
                                        '-------------------------------------------------------
                                        Image_Linie(Grafik_Seite, 249, 4033, 1299 + (Zeilenhoehe * Anzahl_Zeilen), 1299 + (Zeilenhoehe * Anzahl_Zeilen))
                                        '-------------------------------------------------------
                                    Next
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
                                    Listview.Items.Add("Gebäudenutzung", Seitennummer - 1)
                                    Listview.Update()
                                    '-------------------------------------------------------
                                End If
                                '-------------------------------------------------------
                                Variable_XML_Import.Bilddateien(Seitennummer) = Image_to_String(Bitmap_Seite)
                                Variable_XML_Import.Bildanzahl += 1
                                '--------------------------------------------------
                                Speicher_freigeben()
                                '-------------------------------------------------------
                            End If
                            '-------------------------------------------------------
                    End Select
                    '-------------------------------------------------------
            End Select
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
                        Hintergrundgrafik = Image.FromFile(Variable_Parameter.Arbeitsverzeichnis & "/Image/GEG2024_WB_Ueberlauf_Modernisierung.png")
                        Grafik_Seite.DrawImage(Hintergrundgrafik, 0, 0, 4200, 5940)
                        '-------------------------------------------------------
                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, Seitennummer, 0, "", 3893, 4040, 688, 840, Font_Seitennummer, Font_Druckfarbe)
                        '-------------------------------------------------------
                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, Datum_Energiegesetz, 2128, 2616, 422, 516, Font_Schriftgroesse_50, Font_Druckfarbe)
                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.Mitte_kein_Abschneiden, Variable_Steuerung.Registriernummer, 3053, 3768, 722, 820, Font_Schriftgroesse_50, Font_Druckfarbe)
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
#Region "Zusatzseite 65% mit Regel"
            '-------------------------------------------------------
            Select Case Variable_Steuerung.Berechnungsart
                Case "Energiebedarfsausweis"
                    '-------------------------------------------------------
                    Select Case Variable_XML_Import.Berechnungsverfahren
                        Case "Bedarfsberechnung-18599"
                            '-------------------------------------------------------
                            'Startseite
                            '-------------------------------------------------------
                            Wert_Schleife_Offset = 2
                            Wert_Schleife_Maximum = 32
                            Wert_Schleife_Zeilenzahl = 30
                            '-------------------------------------------------------
                            With Variable_XML_Import
                                '-------------------------------------------------------
                                If .Erneuerbare_Energien_65EE_Regel_Anzahl > Wert_Schleife_Offset Then
                                    '-------------------------------------------------------
                                    Seitennummer += 1
                                    '-------------------------------------------------------
                                    Grafik_Seite.Clear(Color.White)
                                    '-------------------------------------------------------
                                    Hintergrundgrafik = Image.FromFile(Variable_Parameter.Arbeitsverzeichnis & "/Image/GEG2024_NWB_Ueberlauf_Erneuerbare_Energien_1_1.png")
                                    Grafik_Seite.DrawImage(Hintergrundgrafik, 0, 0, 4200, 5940)
                                    '-------------------------------------------------------
                                    Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, Seitennummer, 0, "", 3893, 4040, 688, 840, Font_Seitennummer, Font_Druckfarbe)
                                    '-------------------------------------------------------
                                    Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, Datum_Energiegesetz, 2129, 2617, 422, 517, Font_Schriftgroesse_50, Font_Druckfarbe)
                                    Image_Text_schreiben(Grafik_Seite, Ausrichtung.Mitte_kein_Abschneiden, Variable_Steuerung.Registriernummer, 3032, 3748, 718, 818, Font_Schriftgroesse_50, Font_Druckfarbe)
                                    '-------------------------------------------------------
                                    If .Energieausweis_Daten.Keine_Pauschale_Erfuellungsoptionen_Anlagentyp.Wert = True Then
                                        Image_Auswahl_schreiben(Grafik_Seite, True, 291, 341, 1169, 1221)
                                    End If
                                    '-------------------------------------------------------
                                    Zeilenhoehe = (3859 - 1481) / Wert_Schleife_Zeilenzahl
                                    '-------------------------------------------------------
                                    If .Erneuerbare_Energien_65EE_Regel_Anzahl < (Wert_Schleife_Zeilenzahl + Wert_Schleife_Offset) Then
                                        '-------------------------------------------------------
                                        Wert_Schleife_Zeilenzahl = .Erneuerbare_Energien_65EE_Regel_Anzahl - Wert_Schleife_Offset
                                        '-------------------------------------------------------
                                    End If
                                    '-------------------------------------------------------
                                    For Anzahl_Zeilen = 1 To Wert_Schleife_Zeilenzahl
                                        '-------------------------------------------------------
                                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .EE_65EE_Regel(Anzahl_Zeilen + Wert_Schleife_Offset).Waermeerzeuger_Bauweise_18599.Wert, 295, 1627, 1481 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1481 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                        '-------------------------------------------------------
                                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .EE_65EE_Regel(Anzahl_Zeilen + Wert_Schleife_Offset).Art_der_Nutzung_erneuerbaren_Energie.Wert, 1632, 2964, 1481 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1481 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                        '-------------------------------------------------------
                                        If .EE_65EE_Regel(Anzahl_Zeilen + Wert_Schleife_Offset).Deckungsanteil.Wert <> 0 Then
                                            Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .EE_65EE_Regel(Anzahl_Zeilen + Wert_Schleife_Offset).Deckungsanteil.Wert, 0, "", 2969, 3286, 1481 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1481 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                        End If
                                        '-------------------------------------------------------
                                        If .EE_65EE_Regel(Anzahl_Zeilen + Wert_Schleife_Offset).Anteil_der_Pflichterfuellung_Anlage.Wert <> 0 Then
                                            Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .EE_65EE_Regel(Anzahl_Zeilen + Wert_Schleife_Offset).Anteil_der_Pflichterfuellung_Anlage.Wert, 0, "", 3291, 3641, 1481 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1481 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                        End If
                                        '-------------------------------------------------------
                                        If .EE_65EE_Regel(Anzahl_Zeilen + Wert_Schleife_Offset).Anteil_der_Pflichterfuellung_Gesamt.Wert <> 0 Then
                                            Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .EE_65EE_Regel(Anzahl_Zeilen + Wert_Schleife_Offset).Anteil_der_Pflichterfuellung_Gesamt.Wert, 0, "", 3645, 3977, 1481 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1481 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                        End If
                                        '-------------------------------------------------------
                                    Next
                                    '-------------------------------------------------------
                                    For Anzahl_Zeilen = 1 To Wert_Schleife_Zeilenzahl
                                        '-------------------------------------------------------
                                        Image_Linie(Grafik_Seite, 290, 3981, 1481 + (Zeilenhoehe * Anzahl_Zeilen), 1481 + (Zeilenhoehe * Anzahl_Zeilen))
                                        '-------------------------------------------------------
                                    Next
                                    '-------------------------------------------------------
                                    If .Erneuerbare_Energien_65EE_Regel_Anzahl > (Wert_Schleife_Zeilenzahl + Wert_Schleife_Offset) Then
                                        '-------------------------------------------------------
                                        Image_Auswahl_schreiben(Grafik_Seite, True, 344, 385, 3902, 3943)
                                        '-------------------------------------------------------
                                    Else
                                        '-------------------------------------------------------
                                        Image_Auswahl_schreiben(Grafik_Seite, False, 344, 385, 3902, 3943)
                                        '-------------------------------------------------------
                                    End If
                                    '--------------------------------------------------
                                    Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksOben, .Gebaeudebezogene_Daten.Ergaenzdende_Erlaeuterungen_EE.Wert, 245 + 20, 4032, 4263, 5513, Font_Schriftgroesse_40, Font_Druckfarbe)
                                    '--------------------------------------------------
                                    If Variable_Steuerung.Entwurf = True Then
                                        Image_Text_im_Winkel_schreiben(Grafik_Seite, -45, "ENTWURF", 0, 2100, 0, 2970, Font_Schriftgroesse_ENTWURF, Font_Druckfarbe_ENTWURF)
                                    End If
                                    '-------------------------------------------------------
                                    If Oberflaeche = True Then
                                        '-------------------------------------------------------
                                        Imagelist.Images.Add(Bitmap_skalieren(Bitmap_Seite, Skalierungsfaktor))
                                        Listview.Items.Add("65EE mit Regel", Seitennummer - 1)
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
                            'Folgeseite
                            '-------------------------------------------------------
                            Wert_Schleife_Offset = 32
                            Wert_Schleife_Maximum = 200
                            Wert_Schleife_Zeilenzahl = 42
                            '-------------------------------------------------------
                            For Wert_Schleife = 0 To (Wert_Schleife_Maximum) Step Wert_Schleife_Zeilenzahl
                                '-------------------------------------------------------
                                With Variable_XML_Import
                                    '-------------------------------------------------------
                                    If .Erneuerbare_Energien_65EE_Regel_Anzahl > (Wert_Schleife + Wert_Schleife_Offset) Then
                                        '-------------------------------------------------------
                                        Seitennummer += 1
                                        '-------------------------------------------------------
                                        Grafik_Seite.Clear(Color.White)
                                        '-------------------------------------------------------
                                        Hintergrundgrafik = Image.FromFile(Variable_Parameter.Arbeitsverzeichnis & "/Image/GEG2024_NWB_Ueberlauf_Erneuerbare_Energien_2_1.png")
                                        Grafik_Seite.DrawImage(Hintergrundgrafik, 0, 0, 4200, 5940)
                                        '-------------------------------------------------------
                                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, Seitennummer, 0, "", 3893, 4040, 688, 840, Font_Seitennummer, Font_Druckfarbe)
                                        '-------------------------------------------------------
                                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, Datum_Energiegesetz, 2129, 2617, 422, 517, Font_Schriftgroesse_50, Font_Druckfarbe)
                                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.Mitte_kein_Abschneiden, Variable_Steuerung.Registriernummer, 3032, 3748, 718, 818, Font_Schriftgroesse_50, Font_Druckfarbe)
                                        '-------------------------------------------------------
                                        If .Energieausweis_Daten.Keine_Pauschale_Erfuellungsoptionen_Anlagentyp.Wert = True Then
                                            Image_Auswahl_schreiben(Grafik_Seite, True, 291, 341, 1169, 1221)
                                        End If
                                        '-------------------------------------------------------
                                        Zeilenhoehe = (5422 - 1481) / Wert_Schleife_Zeilenzahl
                                        '-------------------------------------------------------
                                        If .Erneuerbare_Energien_65EE_Regel_Anzahl < (Wert_Schleife + Wert_Schleife_Zeilenzahl + Wert_Schleife_Offset) Then
                                            '-------------------------------------------------------
                                            Wert_Schleife_Zeilenzahl = .Erneuerbare_Energien_65EE_Regel_Anzahl - Wert_Schleife_Offset - Wert_Schleife
                                            '-------------------------------------------------------
                                        End If
                                        '-------------------------------------------------------
                                        For Anzahl_Zeilen = 1 To Wert_Schleife_Zeilenzahl
                                            '-------------------------------------------------------
                                            Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .EE_65EE_Regel(Anzahl_Zeilen + Wert_Schleife + Wert_Schleife_Offset).Waermeerzeuger_Bauweise_18599.Wert, 295, 1627, 1481 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1481 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                            '-------------------------------------------------------
                                            Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .EE_65EE_Regel(Anzahl_Zeilen + Wert_Schleife + Wert_Schleife_Offset).Art_der_Nutzung_erneuerbaren_Energie.Wert, 1632, 2964, 1481 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1481 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                            '-------------------------------------------------------
                                            If .EE_65EE_Regel(Anzahl_Zeilen + Wert_Schleife + Wert_Schleife_Offset).Deckungsanteil.Wert <> 0 Then
                                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .EE_65EE_Regel(Anzahl_Zeilen + Wert_Schleife + Wert_Schleife_Offset).Deckungsanteil.Wert, 0, "", 2969, 3286, 1481 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1481 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                            End If
                                            '-------------------------------------------------------
                                            If .EE_65EE_Regel(Anzahl_Zeilen + Wert_Schleife + Wert_Schleife_Offset).Anteil_der_Pflichterfuellung_Anlage.Wert <> 0 Then
                                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .EE_65EE_Regel(Anzahl_Zeilen + Wert_Schleife + Wert_Schleife_Offset).Anteil_der_Pflichterfuellung_Anlage.Wert, 0, "", 3291, 3641, 1481 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1481 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                            End If
                                            '-------------------------------------------------------
                                            If .EE_65EE_Regel(Anzahl_Zeilen + Wert_Schleife + Wert_Schleife_Offset).Anteil_der_Pflichterfuellung_Gesamt.Wert <> 0 Then
                                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .EE_65EE_Regel(Anzahl_Zeilen + Wert_Schleife + Wert_Schleife_Offset).Anteil_der_Pflichterfuellung_Gesamt.Wert, 0, "", 3645, 3977, 1481 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1481 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                            End If
                                            '-------------------------------------------------------
                                        Next
                                        '-------------------------------------------------------
                                        For Anzahl_Zeilen = 1 To Wert_Schleife_Zeilenzahl
                                            '-------------------------------------------------------
                                            Image_Linie(Grafik_Seite, 290, 3981, 1481 + (Zeilenhoehe * Anzahl_Zeilen), 1481 + (Zeilenhoehe * Anzahl_Zeilen))
                                            '-------------------------------------------------------
                                        Next
                                        '-------------------------------------------------------
                                        If .Erneuerbare_Energien_65EE_Regel_Anzahl > (Wert_Schleife + Wert_Schleife_Zeilenzahl + Wert_Schleife_Offset) Then
                                            '-------------------------------------------------------
                                            Image_Auswahl_schreiben(Grafik_Seite, True, 351, 393, 5451, 5492)
                                            '-------------------------------------------------------
                                        Else
                                            '-------------------------------------------------------
                                            Image_Auswahl_schreiben(Grafik_Seite, False, 351, 393, 5451, 5492)
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
                                            Listview.Items.Add("65EE mit Regel", Seitennummer - 1)
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
#Region "Zusatzseite 65% ohne Regel"
            '-------------------------------------------------------
            Select Case Variable_Steuerung.Berechnungsart
                Case "Energiebedarfsausweis"
                    '-------------------------------------------------------
                    Select Case Variable_XML_Import.Berechnungsverfahren
                        Case "Bedarfsberechnung-18599"
                            '-------------------------------------------------------
                            'Startseite
                            '-------------------------------------------------------
                            Wert_Schleife_Offset = 2
                            Wert_Schleife_Maximum = 32
                            Wert_Schleife_Zeilenzahl = 30
                            '-------------------------------------------------------
                            With Variable_XML_Import
                                '-------------------------------------------------------
                                If .Erneuerbare_Energien_65EE_keine_Regel_Anzahl > Wert_Schleife_Offset Then
                                    '-------------------------------------------------------
                                    Seitennummer += 1
                                    '-------------------------------------------------------
                                    Grafik_Seite.Clear(Color.White)
                                    '-------------------------------------------------------
                                    Hintergrundgrafik = Image.FromFile(Variable_Parameter.Arbeitsverzeichnis & "/Image/GEG2024_NWB_Ueberlauf_Erneuerbare_Energien_1_2.png")
                                    Grafik_Seite.DrawImage(Hintergrundgrafik, 0, 0, 4200, 5940)
                                    '-------------------------------------------------------
                                    Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, Seitennummer, 0, "", 3893, 4040, 688, 840, Font_Seitennummer, Font_Druckfarbe)
                                    '-------------------------------------------------------
                                    Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, Datum_Energiegesetz, 2129, 2617, 422, 517, Font_Schriftgroesse_50, Font_Druckfarbe)
                                    Image_Text_schreiben(Grafik_Seite, Ausrichtung.Mitte_kein_Abschneiden, Variable_Steuerung.Registriernummer, 3032, 3748, 718, 818, Font_Schriftgroesse_50, Font_Druckfarbe)
                                    '-------------------------------------------------------
                                    If .Energieausweis_Daten.Nutzung_bei_Bestandsanlagen.Wert = True Then
                                        Image_Auswahl_schreiben(Grafik_Seite, True, 291, 341, 1169, 1221)
                                    End If
                                    '-------------------------------------------------------
                                    Zeilenhoehe = (3859 - 1368) / Wert_Schleife_Zeilenzahl
                                    '-------------------------------------------------------
                                    If .Erneuerbare_Energien_65EE_keine_Regel_Anzahl < (Wert_Schleife_Zeilenzahl + Wert_Schleife_Offset) Then
                                        '-------------------------------------------------------
                                        Wert_Schleife_Zeilenzahl = .Erneuerbare_Energien_65EE_keine_Regel_Anzahl - Wert_Schleife_Offset
                                        '-------------------------------------------------------
                                    End If
                                    '-------------------------------------------------------
                                    For Anzahl_Zeilen = 1 To Wert_Schleife_Zeilenzahl
                                        '-------------------------------------------------------
                                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .EE_65EE_keine_Regel(Anzahl_Zeilen + Wert_Schleife_Offset).Waermeerzeuger_Bauweise_18599.Wert, 295, 1627, 1368 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1368 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                        '-------------------------------------------------------
                                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .EE_65EE_keine_Regel(Anzahl_Zeilen + Wert_Schleife_Offset).Art_der_Nutzung_erneuerbaren_Energie.Wert, 1632, 3640, 1368 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1368 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                        '-------------------------------------------------------
                                        If .EE_65EE_keine_Regel(Anzahl_Zeilen + Wert_Schleife_Offset).Anteil_EE_Anlage.Wert <> 0 Then
                                            Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .EE_65EE_keine_Regel(Anzahl_Zeilen + Wert_Schleife_Offset).Anteil_EE_Anlage.Wert, 0, "", 3645, 3977, 1368 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1368 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                        End If
                                        '-------------------------------------------------------
                                    Next
                                    '-------------------------------------------------------
                                    For Anzahl_Zeilen = 1 To Wert_Schleife_Zeilenzahl
                                        '-------------------------------------------------------
                                        Image_Linie(Grafik_Seite, 290, 3981, 1368 + (Zeilenhoehe * Anzahl_Zeilen), 1368 + (Zeilenhoehe * Anzahl_Zeilen))
                                        '-------------------------------------------------------
                                    Next
                                    '-------------------------------------------------------
                                    If .Erneuerbare_Energien_65EE_keine_Regel_Anzahl > (Wert_Schleife_Zeilenzahl + Wert_Schleife_Offset) Then
                                        '-------------------------------------------------------
                                        Image_Auswahl_schreiben(Grafik_Seite, True, 344, 385, 3902, 3943)
                                        '-------------------------------------------------------
                                    Else
                                        '-------------------------------------------------------
                                        Image_Auswahl_schreiben(Grafik_Seite, False, 344, 385, 3902, 3943)
                                        '-------------------------------------------------------
                                    End If
                                    '--------------------------------------------------
                                    Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksOben, .Gebaeudebezogene_Daten.Ergaenzdende_Erlaeuterungen_EE.Wert, 245 + 20, 4032, 4263, 5513, Font_Schriftgroesse_40, Font_Druckfarbe)
                                    '--------------------------------------------------
                                    If Variable_Steuerung.Entwurf = True Then
                                        Image_Text_im_Winkel_schreiben(Grafik_Seite, -45, "ENTWURF", 0, 2100, 0, 2970, Font_Schriftgroesse_ENTWURF, Font_Druckfarbe_ENTWURF)
                                    End If
                                    '-------------------------------------------------------
                                    If Oberflaeche = True Then
                                        '-------------------------------------------------------
                                        Imagelist.Images.Add(Bitmap_skalieren(Bitmap_Seite, Skalierungsfaktor))
                                        Listview.Items.Add("65EE ohne Regel", Seitennummer - 1)
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
                            'Folgeseite
                            '-------------------------------------------------------
                            Wert_Schleife_Offset = 32
                            Wert_Schleife_Maximum = 200
                            Wert_Schleife_Zeilenzahl = 42
                            '-------------------------------------------------------
                            For Wert_Schleife = 0 To (Wert_Schleife_Maximum) Step Wert_Schleife_Zeilenzahl
                                '-------------------------------------------------------
                                With Variable_XML_Import
                                    '-------------------------------------------------------
                                    If .Erneuerbare_Energien_65EE_keine_Regel_Anzahl > (Wert_Schleife + Wert_Schleife_Offset) Then
                                        '-------------------------------------------------------
                                        Seitennummer += 1
                                        '-------------------------------------------------------
                                        Grafik_Seite.Clear(Color.White)
                                        '-------------------------------------------------------
                                        Hintergrundgrafik = Image.FromFile(Variable_Parameter.Arbeitsverzeichnis & "/Image/GEG2024_NWB_Ueberlauf_Erneuerbare_Energien_2_2.png")
                                        Grafik_Seite.DrawImage(Hintergrundgrafik, 0, 0, 4200, 5940)
                                        '-------------------------------------------------------
                                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, Seitennummer, 0, "", 3893, 4040, 688, 840, Font_Seitennummer, Font_Druckfarbe)
                                        '-------------------------------------------------------
                                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, Datum_Energiegesetz, 2129, 2617, 422, 517, Font_Schriftgroesse_50, Font_Druckfarbe)
                                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.Mitte_kein_Abschneiden, Variable_Steuerung.Registriernummer, 3032, 3748, 718, 818, Font_Schriftgroesse_50, Font_Druckfarbe)
                                        '-------------------------------------------------------
                                        If .Energieausweis_Daten.Nutzung_bei_Bestandsanlagen.Wert = True Then
                                            Image_Auswahl_schreiben(Grafik_Seite, True, 291, 341, 1169, 1221)
                                        End If
                                        '-------------------------------------------------------
                                        Zeilenhoehe = (5422 - 1368) / Wert_Schleife_Zeilenzahl
                                        '-------------------------------------------------------
                                        If .Erneuerbare_Energien_65EE_keine_Regel_Anzahl < (Wert_Schleife + Wert_Schleife_Zeilenzahl + Wert_Schleife_Offset) Then
                                            '-------------------------------------------------------
                                            Wert_Schleife_Zeilenzahl = .Erneuerbare_Energien_65EE_keine_Regel_Anzahl - Wert_Schleife_Offset - Wert_Schleife
                                            '-------------------------------------------------------
                                        End If
                                        '-------------------------------------------------------
                                        For Anzahl_Zeilen = 1 To Wert_Schleife_Zeilenzahl
                                            '-------------------------------------------------------
                                            Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .EE_65EE_keine_Regel(Anzahl_Zeilen + Wert_Schleife + Wert_Schleife_Offset).Waermeerzeuger_Bauweise_18599.Wert, 295, 1627, 1368 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1368 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                            '-------------------------------------------------------
                                            Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .EE_65EE_keine_Regel(Anzahl_Zeilen + Wert_Schleife + Wert_Schleife_Offset).Art_der_Nutzung_erneuerbaren_Energie.Wert, 1632, 3640, 1368 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1368 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                            '-------------------------------------------------------
                                            If .EE_65EE_keine_Regel(Anzahl_Zeilen + Wert_Schleife + Wert_Schleife_Offset).Anteil_EE_Anlage.Wert <> 0 Then
                                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .EE_65EE_keine_Regel(Anzahl_Zeilen + Wert_Schleife + Wert_Schleife_Offset).Anteil_EE_Anlage.Wert, 0, "", 3645, 3977, 1368 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 1368 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                                            End If
                                            '-------------------------------------------------------
                                        Next
                                        '-------------------------------------------------------
                                        For Anzahl_Zeilen = 1 To Wert_Schleife_Zeilenzahl
                                            '-------------------------------------------------------
                                            Image_Linie(Grafik_Seite, 290, 3981, 1368 + (Zeilenhoehe * Anzahl_Zeilen), 1368 + (Zeilenhoehe * Anzahl_Zeilen))
                                            '-------------------------------------------------------
                                        Next
                                        '-------------------------------------------------------
                                        If .Erneuerbare_Energien_65EE_keine_Regel_Anzahl > (Wert_Schleife + Wert_Schleife_Zeilenzahl + Wert_Schleife_Offset) Then
                                            '-------------------------------------------------------
                                            Image_Auswahl_schreiben(Grafik_Seite, True, 351, 393, 5451, 5492)
                                            '-------------------------------------------------------
                                        Else
                                            '-------------------------------------------------------
                                            Image_Auswahl_schreiben(Grafik_Seite, False, 351, 393, 5451, 5492)
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
                                            Listview.Items.Add("65EE ohne Regel", Seitennummer - 1)
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
#Region "Zusatzseite EE ohne Tabelle"
            '-------------------------------------------------------
            Select Case Variable_Steuerung.Berechnungsart
                Case "Energiebedarfsausweis"
                    '-------------------------------------------------------
                    Select Case Variable_XML_Import.Berechnungsverfahren
                        Case "Bedarfsberechnung-18599"
                            '-------------------------------------------------------
                            With Variable_XML_Import
                                '-------------------------------------------------------
                                If .Erneuerbare_Energien_65EE_Regel_Anzahl <= 2 And .Erneuerbare_Energien_65EE_keine_Regel_Anzahl <= 2 And .Energieausweis_Daten.Nichtwohngebaeude.Bedarfswerte_NWG.Weitere_Eintraege_und_Erlaeuterungen_in_der_Anlage.Wert = True Then
                                    '-------------------------------------------------------
                                    Seitennummer += 1
                                    '-------------------------------------------------------
                                    Grafik_Seite.Clear(Color.White)
                                    '-------------------------------------------------------
                                    Hintergrundgrafik = Image.FromFile(Variable_Parameter.Arbeitsverzeichnis & "/Image/GEG2024_NWB_Ueberlauf_Erneuerbare_Energien_1_3.png")
                                    Grafik_Seite.DrawImage(Hintergrundgrafik, 0, 0, 4200, 5940)
                                    '-------------------------------------------------------
                                    Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, Seitennummer, 0, "", 3893, 4040, 688, 840, Font_Seitennummer, Font_Druckfarbe)
                                    '-------------------------------------------------------
                                    Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, Datum_Energiegesetz, 2129, 2617, 422, 517, Font_Schriftgroesse_50, Font_Druckfarbe)
                                    Image_Text_schreiben(Grafik_Seite, Ausrichtung.Mitte_kein_Abschneiden, Variable_Steuerung.Registriernummer, 3032, 3748, 718, 818, Font_Schriftgroesse_50, Font_Druckfarbe)
                                    '-------------------------------------------------------
                                    Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksOben, .Gebaeudebezogene_Daten.Ergaenzdende_Erlaeuterungen_EE.Wert, 245 + 20, 4034, 1109, 5513, Font_Schriftgroesse_50, Font_Druckfarbe)
                                    '--------------------------------------------------
                                    If Variable_Steuerung.Entwurf = True Then
                                        Image_Text_im_Winkel_schreiben(Grafik_Seite, -45, "ENTWURF", 0, 2100, 0, 2970, Font_Schriftgroesse_ENTWURF, Font_Druckfarbe_ENTWURF)
                                    End If
                                    '-------------------------------------------------------
                                    If Oberflaeche = True Then
                                        '-------------------------------------------------------
                                        Imagelist.Images.Add(Bitmap_skalieren(Bitmap_Seite, Skalierungsfaktor))
                                        Listview.Items.Add("65EE", Seitennummer - 1)
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
                    End Select
                    '-------------------------------------------------------
            End Select
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
                            Hintergrundgrafik = Image.FromFile(Variable_Parameter.Arbeitsverzeichnis & "/Image/GEG2024_NWB_Bedarf_Aushang.png")
                            Grafik_Seite.DrawImage(Hintergrundgrafik, 0, 0, 4200, 5940)
                            '-------------------------------------------------------
                            With Variable_XML_Import
                                '-------------------------------------------------------
                                Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, Datum_Energiegesetz, 2129, 2617, 422, 517, Font_Schriftgroesse_50, Font_Druckfarbe)
                                Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, DateAdd(DateInterval.Year, 10, Variable_Steuerung.Ausstellungsdatum), 621 + 20, 1339, 718, 818, Font_Schriftgroesse_50, Font_Druckfarbe)
                                Image_Text_schreiben(Grafik_Seite, Ausrichtung.Mitte_kein_Abschneiden, Variable_Steuerung.Registriernummer, 2770, 3489, 722, 820, Font_Schriftgroesse_50, Font_Druckfarbe)
                                '-------------------------------------------------------
                                Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksOben, .Energieausweis_Daten.Nichtwohngebaeude.Hauptnutzung_Gebaeudekategorie.Wert, 1483 + 20, 3262, 1065, 1210, Font_Schriftgroesse_40, Font_Druckfarbe)
                                Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksMitte, .Gebaeudebezogene_Daten.Gebaeudeadresse_Strasse_Nr.Wert & ", " & .Gebaeudebezogene_Daten.Gebaeudeadresse_Postleitzahl.Wert & " " & .Gebaeudebezogene_Daten.Gebaeudeadresse_Ort.Wert, 1483 + 20, 3262, 1214, 1358, Font_Schriftgroesse_40, Font_Druckfarbe)
                                Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Energieausweis_Daten.Gebaeudeteil.Wert, 1483 + 20, 3262, 1362, 1459, Font_Schriftgroesse_40, Font_Druckfarbe)
                                Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Energieausweis_Daten.Baujahr_Gebaeude.Wert, 1483 + 20, 3262, 1463, 1606, Font_Schriftgroesse_40, Font_Druckfarbe)
                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Energieausweis_Daten.Nichtwohngebaeude.Nettogrundflaeche.Wert, 2, "m²", 1483 + 20, 3262, 1610, 1706, Font_Schriftgroesse_40, Font_Druckfarbe)
                                Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksMitte, .Energieausweis_Daten.wesentliche_Energietraeger_Heizung.Wert, 1483 + 20, 4030, 1710, 1806, Font_Schriftgroesse_40, Font_Druckfarbe)
                                Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksMitte, .Energieausweis_Daten.wesentliche_Energietraeger_Warmwasser.Wert, 1483 + 20, 4030, 1810, 1906, Font_Schriftgroesse_40, Font_Druckfarbe)
                                '--------------------------------------------------
                                Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Lueftungsart_Fensterlueftung.Wert, 1528, 1558, 1945, 1975)
                                Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Lueftungsart_Schachtlueftung.Wert, 1528, 1558, 2016, 2046)
                                Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Lueftungsart_Anlage_m_WRG.Wert, 2695, 2724, 1945, 1975)
                                Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Lueftungsart_Anlage_o_WRG.Wert, 2695, 2724, 2016, 2046)
                                '-------------------------------------------------------            
                                Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Kuehlungsart_passive_Kuehlung.Wert, 1528, 1558, 2110, 2140)
                                Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Kuehlungsart_gelieferte_Kaelte.Wert, 1528, 1558, 2180, 2210)
                                Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Kuehlungsart_Strom.Wert, 2695, 2724, 2110, 2140)
                                Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Kuehlungsart_Waerme.Wert, 2695, 2724, 2180, 2210)
                                '-------------------------------------------------------
                                Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksMitte, "Art: " & .Energieausweis_Daten.Erneuerbare_Art.Wert, 1483 + 20, 2647, 2238, 2393, Font_Schriftgroesse_40, Font_Druckfarbe)
                                Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksMitte, "Verwendung: " & .Energieausweis_Daten.Erneuerbare_Verwendung.Wert, 2651 + 20, 4030, 2238, 2393, Font_Schriftgroesse_40, Font_Druckfarbe)
                                '-------------------------------------------------------
                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Treibhausgasemissionen.Wert, 2, "", 2635, 2937, 2672, 2775, Font_Schriftgroesse_50, Font_Druckfarbe)
                                '-------------------------------------------------------
                                Image_Pfeil_NWB(Grafik_Seite, Lage.Oben, .Energieausweis_Daten.Nichtwohngebaeude.Bedarfswerte_NWG.Primaerenergiebedarf_NGF.Wert, 1, .Energieausweis_Daten.Nichtwohngebaeude.Bedarfswerte_NWG.Primaerenergiebedarf_Hoechstwert_Bestand.Wert, 3, "Primärenergiebedarf dieses Gebäudes", "", 616, 3697, 3121, 3456)
                                '-------------------------------------------------------
                                Image_Diagram_NWB_Energieaufteilung(Grafik_Seite, 471, 2406, 4007, 4795)
                                '-------------------------------------------------------
                                Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, FormatDateTime(Variable_Steuerung.Ausstellungsdatum, DateFormat.ShortDate).ToString, 3545, 3983, 5427, 5522, Font_Schriftgroesse_50, Font_Druckfarbe)
                                '-------------------------------------------------------
                                Image_Grafik_Projekt_schreiben(Grafik_Seite, Variable_Steuerung.Bild_Projekt, 3266, 1065, 4030 - 3266, 1704 - 1065)
                                '-------------------------------------------------------
                                Image_Grafik_Aussteller_schreiben(Grafik_Seite, Variable_Steuerung.Bild_Aussteller, .Gebaeudebezogene_Daten.Aussteller_Bezeichnung.Wert, .Gebaeudebezogene_Daten.Ausstellervorname.Wert, .Gebaeudebezogene_Daten.Ausstellername.Wert, .Gebaeudebezogene_Daten.Aussteller_Strasse_Nr.Wert, .Gebaeudebezogene_Daten.Aussteller_PLZ.Wert, .Gebaeudebezogene_Daten.Aussteller_Ort.Wert, 271, 5132, 2500, 300)
                                '-------------------------------------------------------
                                Image_Grafik_Unterschrift_schreiben(Grafik_Seite, Variable_Steuerung.Bild_Unterschrift, 3000, 5132, 1000, 300)
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
                    Hintergrundgrafik = Image.FromFile(Variable_Parameter.Arbeitsverzeichnis & "/Image/GEG2024_NWB_Verbrauch_Aushang.png")
                    Grafik_Seite.DrawImage(Hintergrundgrafik, 0, 0, 4200, 5940)
                    '-------------------------------------------------------
                    With Variable_XML_Import
                        '-------------------------------------------------------
                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, Datum_Energiegesetz, 2129, 2617, 422, 517, Font_Schriftgroesse_50, Font_Druckfarbe)
                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, DateAdd(DateInterval.Year, 10, Variable_Steuerung.Ausstellungsdatum), 621 + 20, 1339, 718, 818, Font_Schriftgroesse_50, Font_Druckfarbe)
                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.Mitte_kein_Abschneiden, Variable_Steuerung.Registriernummer, 2770, 3489, 722, 820, Font_Schriftgroesse_50, Font_Druckfarbe)
                        '-------------------------------------------------------
                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksOben, .Energieausweis_Daten.Nichtwohngebaeude.Hauptnutzung_Gebaeudekategorie.Wert, 1483 + 20, 3262, 1053, 1198, Font_Schriftgroesse_40, Font_Druckfarbe)
                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksMitte, .Gebaeudebezogene_Daten.Gebaeudeadresse_Strasse_Nr.Wert & ", " & .Gebaeudebezogene_Daten.Gebaeudeadresse_Postleitzahl.Wert & " " & .Gebaeudebezogene_Daten.Gebaeudeadresse_Ort.Wert, 1483 + 20, 3262, 1202, 1346, Font_Schriftgroesse_40, Font_Druckfarbe)
                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Energieausweis_Daten.Gebaeudeteil.Wert, 1483 + 20, 3262, 1350, 1447, Font_Schriftgroesse_40, Font_Druckfarbe)
                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Energieausweis_Daten.Baujahr_Gebaeude.Wert, 1483 + 20, 3262, 1451, 1594, Font_Schriftgroesse_40, Font_Druckfarbe)
                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Energieausweis_Daten.Nichtwohngebaeude.Nettogrundflaeche.Wert, 2, "m²", 1483 + 20, 3262, 1598, 1694, Font_Schriftgroesse_40, Font_Druckfarbe)
                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksMitte, .Energieausweis_Daten.wesentliche_Energietraeger_Heizung.Wert, 1483 + 20, 4030, 1698, 1794, Font_Schriftgroesse_40, Font_Druckfarbe)
                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksMitte, .Energieausweis_Daten.wesentliche_Energietraeger_Warmwasser.Wert, 1483 + 20, 4030, 1798, 1894, Font_Schriftgroesse_40, Font_Druckfarbe)
                        '--------------------------------------------------
                        Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Lueftungsart_Fensterlueftung.Wert, 1528, 1558, 1933, 1964)
                        Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Lueftungsart_Schachtlueftung.Wert, 1528, 1558, 2004, 2035)
                        Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Lueftungsart_Anlage_m_WRG.Wert, 2695, 2724, 1933, 1964)
                        Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Lueftungsart_Anlage_o_WRG.Wert, 2695, 2724, 2004, 2035)
                        '-------------------------------------------------------            
                        Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Kuehlungsart_passive_Kuehlung.Wert, 1528, 1558, 2097, 2128)
                        Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Kuehlungsart_gelieferte_Kaelte.Wert, 1528, 1558, 2168, 2199)
                        Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Kuehlungsart_Strom.Wert, 2695, 2724, 2097, 2128)
                        Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Kuehlungsart_Waerme.Wert, 2695, 2724, 2168, 2199)
                        '-------------------------------------------------------
                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksMitte, "Art: " & .Energieausweis_Daten.Erneuerbare_Art.Wert, 1483 + 20, 2647, 2227, 2381, Font_Schriftgroesse_40, Font_Druckfarbe)
                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksMitte, "Verwendung: " & .Energieausweis_Daten.Erneuerbare_Verwendung.Wert, 2651 + 20, 4030, 2227, 2381, Font_Schriftgroesse_40, Font_Druckfarbe)
                        '-------------------------------------------------------
                        Image_Pfeil_NWB(Grafik_Seite, Lage.Oben, .Energieausweis_Daten.Nichtwohngebaeude.Verbrauchswerte_NWG.Endenergieverbrauch_Waerme.Wert, 1, .Energieausweis_Daten.Nichtwohngebaeude.Verbrauchswerte_NWG.Endenergieverbrauch_Waerme_Vergleichswert.Wert, 2, "Endenergieverbrauch Wärme", "[Pflichtangabe in Immobilienanzeigen]", 616, 3697, 2981, 3316)
                        '-------------------------------------------------------
                        Image_Pfeil_NWB(Grafik_Seite, Lage.Oben, .Energieausweis_Daten.Nichtwohngebaeude.Verbrauchswerte_NWG.Endenergieverbrauch_Strom.Wert, 1, .Energieausweis_Daten.Nichtwohngebaeude.Verbrauchswerte_NWG.Endenergieverbrauch_Strom_Vergleichswert.Wert, 2, "Endenergieverbrauch Strom", "[Pflichtangabe in Immobilienanzeigen]", 616, 3697, 3917, 4252)
                        '-------------------------------------------------------
                        Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Nichtwohngebaeude.Verbrauchswerte_NWG.Warmwasser_enthalten.Wert, 301, 340, 3389, 3422)
                        Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Nichtwohngebaeude.Verbrauchswerte_NWG.Kuehlung_enthalten.Wert, 301, 340, 3470, 3507)
                        '-------------------------------------------------------
                        Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Nichtwohngebaeude.Verbrauchswerte_NWG.Strom_Daten.Stromverbrauch_enthaelt_Zusatzheizung.Wert, 309, 344, 4617, 4653)
                        Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Nichtwohngebaeude.Verbrauchswerte_NWG.Strom_Daten.Stromverbrauch_enthaelt_Warmwasser.Wert, 939, 977, 4617, 4653)
                        Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Nichtwohngebaeude.Verbrauchswerte_NWG.Strom_Daten.Stromverbrauch_enthaelt_Lueftung.Wert, 1553, 1592, 4617, 4653)
                        Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Nichtwohngebaeude.Verbrauchswerte_NWG.Strom_Daten.Stromverbrauch_enthaelt_Beleuchtung.Wert, 1997, 2039, 4617, 4653)
                        Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Nichtwohngebaeude.Verbrauchswerte_NWG.Strom_Daten.Stromverbrauch_enthaelt_Kuehlung.Wert, 2991, 3029, 4617, 4653)
                        Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Nichtwohngebaeude.Verbrauchswerte_NWG.Strom_Daten.Stromverbrauch_enthaelt_Sonstiges.Wert, 3482, 3520, 4617, 4653)
                        '-------------------------------------------------------
                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Nichtwohngebaeude.Verbrauchswerte_NWG.Primaerenergieverbrauch.Wert, 1, "", 2962, 3668, 4749, 4888, Font_Schriftgroesse_60, Font_Druckfarbe)
                        '-------------------------------------------------------
                        If Variable_XML_Import.Zusaetzliche_Verbrauchsberechnung = False Then
                            Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Treibhausgasemissionen.Wert, 2, "", 2962, 3668, 4892, 5020, Font_Schriftgroesse_60, Font_Druckfarbe)
                        Else
                            Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Treibhausgasemissionen_Zusaetzliche_Verbrauchsdaten.Wert, 2, "", 2962, 3668, 4892, 5020, Font_Schriftgroesse_60, Font_Druckfarbe)
                        End If
                        '-------------------------------------------------------
                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, FormatDateTime(Variable_Steuerung.Ausstellungsdatum, DateFormat.ShortDate).ToString, 3549, 3983, 5488, 5582, Font_Schriftgroesse_50, Font_Druckfarbe)
                        '-------------------------------------------------------
                        Image_Grafik_Projekt_schreiben(Grafik_Seite, Variable_Steuerung.Bild_Projekt, 3262, 1053, 4030 - 3262, 1694 - 1053)
                        '-------------------------------------------------------
                        Image_Grafik_Aussteller_schreiben(Grafik_Seite, Variable_Steuerung.Bild_Aussteller, .Gebaeudebezogene_Daten.Aussteller_Bezeichnung.Wert, .Gebaeudebezogene_Daten.Ausstellervorname.Wert, .Gebaeudebezogene_Daten.Ausstellername.Wert, .Gebaeudebezogene_Daten.Aussteller_Strasse_Nr.Wert, .Gebaeudebezogene_Daten.Aussteller_PLZ.Wert, .Gebaeudebezogene_Daten.Aussteller_Ort.Wert, 271, 5193, 2500, 300)
                        '-------------------------------------------------------
                        Image_Grafik_Unterschrift_schreiben(Grafik_Seite, Variable_Steuerung.Bild_Unterschrift, 3000, 5193, 1000, 300)
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
