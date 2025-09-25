#Region "Imports"
Imports System.Drawing
Imports System.Windows.Forms
#End Region
'--------------------------------------------------
Public Class Klasse_Energieausweis_WB_2016
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
    ''' Hier wird der gesamte Ablauf für die Erstellung der Energieausweises abgearbeitet - Wohnbau - EnEV 2016
    ''' </summary>
    Sub Energieausweis_erzeugen(ByVal Oberflaeche As Boolean, ByVal PictureBox As PictureBox, ByVal Listview As ListView, ByVal Imagelist As ImageList)
        '-------------------------------------------------------
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
            Hintergrundgrafik = Image.FromFile(Variable_Parameter.Arbeitsverzeichnis & "/Image/ENEV2016_WB_Seite_1.png")
            Grafik_Seite.DrawImage(Hintergrundgrafik, 0, 0, 4200, 5940)
            '-------------------------------------------------------
            With Variable_XML_Import
                '-------------------------------------------------------
                Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, Datum_Energiegesetz, 2450, 2912, 448, 545, Font_Schriftgroesse_50, Font_Druckfarbe)
                Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, DateAdd(DateInterval.Year, 10, Variable_Steuerung.Ausstellungsdatum), 549 + 20, 1000, 781, 831, Font_Schriftgroesse_50, Font_Druckfarbe)
                Image_Text_schreiben(Grafik_Seite, Ausrichtung.Mitte_kein_Abschneiden, Variable_Steuerung.Registriernummer, 3263, 3726, 662, 759, Font_Bold_Schriftgroesse_40_Narrow, Font_Druckfarbe)
                '-------------------------------------------------------
                Select Case Variable_Steuerung.Berechnungsart
                    Case "Energiebedarfsausweis"
                        '-------------------------------------------------------
                        Select Case .Berechnungsverfahren
                            Case "Bedarfsberechnung-4108-4701"
                                '-------------------------------------------------------
                                Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Energieausweis_Daten.Wohngebaeude.Gebaeudetyp.Wert & ", " & .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_4108_4701.Wohngebaeude_Anbaugrad.Wert, 1307 + 20, 3233, 1126, 1247, Font_Schriftgroesse_50, Font_Druckfarbe)
                            '-------------------------------------------------------
                            Case "Bedarfsberechnung-18599"
                                '-------------------------------------------------------
                                Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Energieausweis_Daten.Wohngebaeude.Gebaeudetyp.Wert & ", " & .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_18599.Wohngebaeude_Anbaugrad.Wert, 1307 + 20, 3233, 1126, 1247, Font_Schriftgroesse_50, Font_Druckfarbe)
                            '-------------------------------------------------------
                            Case "Bedarfsberechnung-easy"
                                '-------------------------------------------------------
                                Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Energieausweis_Daten.Wohngebaeude.Gebaeudetyp.Wert & ", " & .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_easy.Wohngebaeude_Anbaugrad.Wert, 1307 + 20, 3233, 1126, 1247, Font_Schriftgroesse_50, Font_Druckfarbe)
                                '-------------------------------------------------------
                        End Select
                    '-------------------------------------------------------
                    Case "Energieverbrauchsausweis"
                        '-------------------------------------------------------
                        Select Case .Berechnungsverfahren
                            Case "Verbrauchsberechnung"
                                '-------------------------------------------------------
                                Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Energieausweis_Daten.Wohngebaeude.Gebaeudetyp.Wert, 1307 + 20, 3233, 1126, 1247, Font_Schriftgroesse_50, Font_Druckfarbe)
                                '-------------------------------------------------------
                        End Select
                        '-------------------------------------------------------
                End Select
                '-------------------------------------------------------
                Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Gebaeudebezogene_Daten.Gebaeudeadresse_Strasse_Nr.Wert & ", " & .Gebaeudebezogene_Daten.Gebaeudeadresse_Postleitzahl.Wert & " " & .Gebaeudebezogene_Daten.Gebaeudeadresse_Ort.Wert, 1307 + 20, 3233, 1255, 1376, Font_Schriftgroesse_50, Font_Druckfarbe)
                Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Energieausweis_Daten.Gebaeudeteil.Wert, 1307 + 20, 3233, 1386, 1506, Font_Schriftgroesse_50, Font_Druckfarbe)
                Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Energieausweis_Daten.Baujahr_Gebaeude.Wert, 1307 + 20, 3233, 1516, 1635, Font_Schriftgroesse_50, Font_Druckfarbe)
                Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Energieausweis_Daten.Baujahr_Waermeerzeuger.Wert, 1307 + 20, 3233, 1644, 1765, Font_Schriftgroesse_50, Font_Druckfarbe)
                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Energieausweis_Daten.Wohngebaeude.Anzahl_Wohneinheiten.Wert, 0, "", 1307 + 20, 3233, 1774, 1896, Font_Schriftgroesse_50, Font_Druckfarbe)
                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Energieausweis_Daten.Wohngebaeude.Gebaeudenutzflaeche.Wert, 2, "m²", 1307 + 20, 1766, 1905, 2026, Font_Schriftgroesse_50, Font_Druckfarbe)
                '-------------------------------------------------------
                Select Case Variable_Steuerung.Berechnungsart
                    Case "Energieverbrauchsausweis"
                        Select Case .Berechnungsverfahren
                            Case "Verbrauchsberechnung"
                                '-------------------------------------------------------
                                Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Wohngebaeude.Verbrauchswerte.Flaechenermittlung_AN_aus_Wohnflaeche.Wert, 1819, 1858, 1952, 1991)
                                '-------------------------------------------------------
                        End Select
                End Select
                '-------------------------------------------------------
                Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksMitte, .Energieausweis_Daten.wesentliche_Energietraeger_Heizung.Wert, 1307 + 20, 3233, 2034, 2242, Font_Schriftgroesse_50, Font_Druckfarbe)
                Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Energieausweis_Daten.Erneuerbare_Art.Wert, 1454 + 20, 2679, 2251 + 15, 2372, Font_Schriftgroesse_50, Font_Druckfarbe)
                Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Energieausweis_Daten.Erneuerbare_Verwendung.Wert, 3154 + 20, 4031, 2251 + 15, 2372, Font_Schriftgroesse_50, Font_Druckfarbe)
                ''--------------------------------------------------
                Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Lueftungsart_Fensterlueftung.Wert, 1350, 1389, 2429, 2467)
                Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Lueftungsart_Schachtlueftung.Wert, 1350, 1389, 2516, 2556)
                Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Lueftungsart_Anlage_m_WRG.Wert, 2034, 2072, 2429, 2467)
                Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Lueftungsart_Anlage_o_WRG.Wert, 2034, 2072, 2516, 2556)
                Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Anlage_zur_Kuehlung.Wert, 3574, 3612, 2429, 2467)
                '-------------------------------------------------------            
                Select Case .Energieausweis_Daten.Ausstellungsanlass.Wert
                    Case "Neubau"
                        Image_Auswahl_schreiben(Grafik_Seite, True, 1350, 1389, 2637, 2674)
                        Image_Auswahl_schreiben(Grafik_Seite, False, 2351, 2390, 2637, 2674)
                        Image_Auswahl_schreiben(Grafik_Seite, False, 1350, 1389, 2726, 2764)
                        Image_Auswahl_schreiben(Grafik_Seite, False, 3470, 3510, 2637, 2674)
                    Case "Modernisierung-Erweiterung"
                        Image_Auswahl_schreiben(Grafik_Seite, False, 1350, 1389, 2637, 2674)
                        Image_Auswahl_schreiben(Grafik_Seite, True, 2351, 2390, 2637, 2674)
                        Image_Auswahl_schreiben(Grafik_Seite, False, 1350, 1389, 2726, 2764)
                        Image_Auswahl_schreiben(Grafik_Seite, False, 3470, 3510, 2637, 2674)
                    Case "Vermietung-Verkauf"
                        Image_Auswahl_schreiben(Grafik_Seite, False, 1350, 1389, 2637, 2674)
                        Image_Auswahl_schreiben(Grafik_Seite, False, 2351, 2390, 2637, 2674)
                        Image_Auswahl_schreiben(Grafik_Seite, True, 1350, 1389, 2726, 2764)
                        Image_Auswahl_schreiben(Grafik_Seite, False, 3470, 3510, 2637, 2674)
                    Case "Sonstiges"
                        Image_Auswahl_schreiben(Grafik_Seite, False, 1350, 1389, 2637, 2674)
                        Image_Auswahl_schreiben(Grafik_Seite, False, 2351, 2390, 2637, 2674)
                        Image_Auswahl_schreiben(Grafik_Seite, False, 1350, 1389, 2726, 2764)
                        Image_Auswahl_schreiben(Grafik_Seite, True, 3470, 3510, 2637, 2674)
                End Select
                '-------------------------------------------------------
                Select Case Variable_Steuerung.Berechnungsart
                    Case "Energiebedarfsausweis"
                        Image_Auswahl_schreiben(Grafik_Seite, True, 269, 308, 3646, 3685)
                        Image_Auswahl_schreiben(Grafik_Seite, False, 269, 308, 3935, 3973)
                    Case "Energieverbrauchsausweis"
                        Image_Auswahl_schreiben(Grafik_Seite, False, 269, 308, 3646, 3685)
                        Image_Auswahl_schreiben(Grafik_Seite, True, 269, 308, 3935, 3973)
                End Select
                '-------------------------------------------------------
                Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Datenerhebung_Eigentuemer.Wert, 2096, 2136, 4181, 4220)
                Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Datenerhebung_Aussteller.Wert, 2969, 3006, 4181, 4220)
                '-------------------------------------------------------
                Image_Auswahl_schreiben(Grafik_Seite, .Gebaeudebezogene_Daten.Zusatzinfos_beigefuegt.Wert, 269, 308, 4299, 4338)
                '-------------------------------------------------------
                Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, FormatDateTime(Variable_Steuerung.Ausstellungsdatum, DateFormat.ShortDate).ToString, 2226, 2734, 5379, 5436, Font_Schriftgroesse_50, Font_Druckfarbe)
                '-------------------------------------------------------
                Image_Grafik_Projekt_schreiben(Grafik_Seite, Variable_Steuerung.Bild_Projekt, 3241, 1126, 4031 - 3241, 2242 - 1126)
                '-------------------------------------------------------
                Image_Grafik_Aussteller_schreiben(Grafik_Seite, Variable_Steuerung.Bild_Aussteller, .Gebaeudebezogene_Daten.Aussteller_Bezeichnung.Wert, .Gebaeudebezogene_Daten.Ausstellervorname.Wert, .Gebaeudebezogene_Daten.Ausstellername.Wert, .Gebaeudebezogene_Daten.Aussteller_Strasse_Nr.Wert, .Gebaeudebezogene_Daten.Aussteller_PLZ.Wert, .Gebaeudebezogene_Daten.Aussteller_Ort.Wert, 277, 5117, 1750, 300)
                '-------------------------------------------------------
                Image_Grafik_Unterschrift_schreiben(Grafik_Seite, Variable_Steuerung.Bild_Unterschrift, 3000, 5125, 1000, 300)
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
            Hintergrundgrafik = Image.FromFile(Variable_Parameter.Arbeitsverzeichnis & "/Image/ENEV2016_WB_Seite_2.png")
            Grafik_Seite.DrawImage(Hintergrundgrafik, 0, 0, 4200, 5940)
            '-------------------------------------------------------
            With Variable_XML_Import
                Select Case Variable_Steuerung.Berechnungsart
                    Case "Energiebedarfsausweis"
                        '-------------------------------------------------------
                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, Datum_Energiegesetz, 2453, 2913, 446, 545, Font_Schriftgroesse_50, Font_Druckfarbe)
                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.Mitte_kein_Abschneiden, Variable_Steuerung.Registriernummer, 3237, 3699, 646, 743, Font_Bold_Schriftgroesse_40_Narrow, Font_Druckfarbe)
                        '-------------------------------------------------------
                        Select Case .Berechnungsverfahren
                            Case "Bedarfsberechnung-4108-4701"
                                '-------------------------------------------------------
                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Gebaeudebezogene_Daten.Treibhausgasemissionen.Wert, 1, "", 3467, 3675, 1128, 1253, Font_Schriftgroesse_40, Font_Druckfarbe)
                                '-------------------------------------------------------
                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_4108_4701.Endenergiebedarf_Waerme_AN.Wert + .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_4108_4701.Endenergiebedarf_Hilfsenergie_AN.Wert, 1, "", 3102, 3633, 3025, 3237, Font_Schriftgroesse_70, Font_Druckfarbe)
                                '-------------------------------------------------------
                                Select Case .Energieausweis_Daten.Ausstellungsanlass.Wert
                                    Case "Neubau", "Modernisierung-Erweiterung"
                                        '-------------------------------------------------------
                                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_4108_4701.Primaerenergiebedarf.Wert, 1, "", 514, 725, 2511, 2615, Font_Schriftgroesse_40, Font_Druckfarbe)
                                        '-------------------------------------------------------
                                        Select Case Variable_Steuerung.Neubau
                                            Case 0 'Modernisierung
                                                '-------------------------------------------------------
                                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_4108_4701.Primaerenergiebedarf_Hoechstwert_Bestand.Wert, 1, "", 1587, 1783, 2511, 2615, Font_Schriftgroesse_40, Font_Druckfarbe)
                                                '-------------------------------------------------------
                                            Case 1 'Neubau
                                                '-------------------------------------------------------
                                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_4108_4701.Primaerenergiebedarf_Hoechstwert_Neubau.Wert, 1, "", 1587, 1783, 2511, 2615, Font_Schriftgroesse_40, Font_Druckfarbe)
                                                '-------------------------------------------------------
                                        End Select
                                        '-------------------------------------------------------
                                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_4108_4701.spezifischer_Transmissionswaermeverlust_Ist.Wert, 2, "", 514, 725, 2739, 2851, Font_Schriftgroesse_40, Font_Druckfarbe)
                                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_4108_4701.spezifischer_Transmissionswaermeverlust_Hoechstwert.Wert, 2, "", 1587, 1783, 2739, 2851, Font_Schriftgroesse_40, Font_Druckfarbe)
                                        '-------------------------------------------------------
                                        Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_4108_4701.Sommerlicher_Waermeschutz.Wert, 1466, 1497, 2915, 2948)
                                        '-------------------------------------------------------
                                    Case "Vermietung-Verkauf", "Sonstiges"
                                        '-------------------------------------------------------
                                End Select
                                '-------------------------------------------------------
                                Image_Auswahl_schreiben(Grafik_Seite, True, 2313, 2344, 2553, 2585)
                                Image_Auswahl_schreiben(Grafik_Seite, False, 2313, 2344, 2667, 2698)
                                Image_Auswahl_schreiben(Grafik_Seite, False, 2313, 2344, 2784, 2815)
                                '-------------------------------------------------------
                                Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_4108_4701.Vereinfachte_Datenaufnahme.Wert, 2313, 2344, 2915, 2948)
                                '-------------------------------------------------------
                                Image_Pfeil_WB(Grafik_Seite, Lage.Oben, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_4108_4701.Endenergiebedarf_Waerme_AN.Wert + .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_4108_4701.Endenergiebedarf_Hilfsenergie_AN.Wert, 1, "Endenergiebedarf dieses Gebäudes", 542, 3623, 1535, 1874)
                                '-------------------------------------------------------
                                Image_Pfeil_WB(Grafik_Seite, Lage.Unten, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_4108_4701.Primaerenergiebedarf.Wert, 1, "Primärenergiebedarf dieses Gebäudes", 542, 3623, 1535, 1874)
                                '-------------------------------------------------------
                                'Erneuerbare Energien ----------------------------------
                                '-------------------------------------------------------
                                Select Case .Energieausweis_Daten.Ausstellungsanlass.Wert
                                    Case "Neubau"
                                        '-------------------------------------------------------
                                        If .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_4108_4701.Art_der_Nutzung_erneuerbaren_Energie_1.Wert <> "" Then
                                            Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_4108_4701.Art_der_Nutzung_erneuerbaren_Energie_1.Wert, 406, 927, 3748, 3868, Font_Schriftgroesse_40, Font_Druckfarbe)
                                            Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_4108_4701.Deckungsanteil_1.Wert, 1, "", 1426, 1669, 3748, 3868, Font_Schriftgroesse_40, Font_Druckfarbe)
                                        End If
                                        If .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_4108_4701.Art_der_Nutzung_erneuerbaren_Energie_2.Wert <> "" Then
                                            Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_4108_4701.Art_der_Nutzung_erneuerbaren_Energie_2.Wert, 406, 927, 3878, 3993, Font_Schriftgroesse_40, Font_Druckfarbe)
                                            Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_4108_4701.Deckungsanteil_2.Wert, 1, "", 1426, 1669, 3878, 3993, Font_Schriftgroesse_40, Font_Druckfarbe)
                                        End If
                                        If .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_4108_4701.Art_der_Nutzung_erneuerbaren_Energie_3.Wert <> "" Then
                                            Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_4108_4701.Art_der_Nutzung_erneuerbaren_Energie_3.Wert, 406, 927, 4002, 4109, Font_Schriftgroesse_40, Font_Druckfarbe)
                                            Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_4108_4701.Deckungsanteil_3.Wert, 1, "", 1426, 1669, 4002, 4109, Font_Schriftgroesse_40, Font_Druckfarbe)
                                        End If
                                        '-------------------------------------------------------
                                        Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_4108_4701.verschaerft_nach_EEWaermeG_7_1_2_eingehalten.Wert, 273, 304, 4557, 4588)
                                        If .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_4108_4701.verschaerft_nach_EEWaermeG_8.Wert > 0 Then
                                            Image_Auswahl_schreiben(Grafik_Seite, True, 273, 304, 4796, 4827)
                                            Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_4108_4701.verschaerft_nach_EEWaermeG_8.Wert, 0, "", 1469, 1665, 4749, 4831, Font_Schriftgroesse_40, Font_Druckfarbe)
                                        Else
                                            Image_Auswahl_schreiben(Grafik_Seite, False, 273, 304, 4796, 4827)
                                        End If
                                        If .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_4108_4701.Primaerenergiebedarf_Hoechstwert_verschaerft.Wert > 0 Then
                                            Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_4108_4701.Primaerenergiebedarf_Hoechstwert_verschaerft.Wert, 1, "", 1213, 1483, 5051, 5135, Font_Schriftgroesse_40, Font_Druckfarbe)
                                        End If
                                        If .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_4108_4701.spezifischer_Transmissionswaermetransferkoeffizient_verschaerft.Wert > 0 Then
                                            Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_4108_4701.spezifischer_Transmissionswaermetransferkoeffizient_verschaerft.Wert, 2, "", 1213, 1483, 5286, 5369, Font_Schriftgroesse_40, Font_Druckfarbe)
                                        End If
                                        '-------------------------------------------------------
                                    Case "Modernisierung-Erweiterung", "Vermietung-Verkauf", "Sonstiges"
                                        '-------------------------------------------------------
                                End Select
                                '-------------------------------------------------------
                            Case "Bedarfsberechnung-18599"
                                '-------------------------------------------------------
                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Gebaeudebezogene_Daten.Treibhausgasemissionen.Wert, 1, "", 3467, 3675, 1128, 1253, Font_Schriftgroesse_40, Font_Druckfarbe)
                                '-------------------------------------------------------
                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_18599.Endenergiebedarf_Waerme_AN.Wert + .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_18599.Endenergiebedarf_Hilfsenergie_AN.Wert, 1, "", 3102, 3633, 3025, 3237, Font_Schriftgroesse_70, Font_Druckfarbe)
                                '-------------------------------------------------------
                                Select Case .Energieausweis_Daten.Ausstellungsanlass.Wert
                                    Case "Neubau", "Modernisierung-Erweiterung"
                                        '-------------------------------------------------------
                                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_18599.Primaerenergiebedarf_AN.Wert, 1, "", 514, 725, 2511, 2615, Font_Schriftgroesse_40, Font_Druckfarbe)
                                        '-------------------------------------------------------
                                        Select Case Variable_Steuerung.Neubau
                                            Case 0 'Modernisierung
                                                '-------------------------------------------------------
                                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_18599.Primaerenergiebedarf_Hoechstwert_Bestand.Wert, 1, "", 1587, 1783, 2511, 2615, Font_Schriftgroesse_40, Font_Druckfarbe)
                                                '-------------------------------------------------------
                                            Case 1 'Neubau
                                                '-------------------------------------------------------
                                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_18599.Primaerenergiebedarf_Hoechstwert_Neubau.Wert, 1, "", 1587, 1783, 2511, 2615, Font_Schriftgroesse_40, Font_Druckfarbe)
                                                '-------------------------------------------------------
                                        End Select
                                        '-------------------------------------------------------
                                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_18599.spezifischer_Transmissionswaermetransferkoeffizient_Ist.Wert, 2, "", 514, 725, 2739, 2851, Font_Schriftgroesse_40, Font_Druckfarbe)
                                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_18599.spezifischer_Transmissionswaermetransferkoeffizient_Hoechstwert.Wert, 2, "", 1587, 1783, 2739, 2851, Font_Schriftgroesse_40, Font_Druckfarbe)
                                        '-------------------------------------------------------
                                        Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_18599.Sommerlicher_Waermeschutz.Wert, 1466, 1497, 2915, 2948)
                                        '-------------------------------------------------------
                                    Case "Vermietung-Verkauf", "Sonstiges"
                                        '-------------------------------------------------------
                                End Select
                                '-------------------------------------------------------
                                Image_Auswahl_schreiben(Grafik_Seite, False, 2313, 2344, 2553, 2585)
                                Image_Auswahl_schreiben(Grafik_Seite, True, 2313, 2344, 2667, 2698)
                                Image_Auswahl_schreiben(Grafik_Seite, False, 2313, 2344, 2784, 2815)
                                '-------------------------------------------------------
                                Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_18599.Vereinfachte_Datenaufnahme.Wert, 2313, 2344, 2915, 2948)
                                '-------------------------------------------------------
                                Image_Pfeil_WB(Grafik_Seite, Lage.Oben, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_18599.Endenergiebedarf_Waerme_AN.Wert + .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_18599.Endenergiebedarf_Hilfsenergie_AN.Wert, 1, "Endenergiebedarf dieses Gebäudes", 542, 3623, 1535, 1874)
                                '-------------------------------------------------------
                                Image_Pfeil_WB(Grafik_Seite, Lage.Unten, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_18599.Primaerenergiebedarf_AN.Wert, 1, "Primärenergiebedarf dieses Gebäudes", 542, 3623, 1535, 1874)
                                '-------------------------------------------------------
                                'Erneuerbare Energien ----------------------------------
                                '-------------------------------------------------------
                                Select Case .Energieausweis_Daten.Ausstellungsanlass.Wert
                                    Case "Neubau"
                                        '-------------------------------------------------------
                                        If .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_18599.Art_der_Nutzung_erneuerbaren_Energie_1.Wert <> "" Then
                                            Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_18599.Art_der_Nutzung_erneuerbaren_Energie_1.Wert, 406, 927, 3748, 3868, Font_Schriftgroesse_40, Font_Druckfarbe)
                                            Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_18599.Deckungsanteil_1.Wert, 1, "", 1426, 1669, 3748, 3868, Font_Schriftgroesse_40, Font_Druckfarbe)
                                        End If
                                        If .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_18599.Art_der_Nutzung_erneuerbaren_Energie_2.Wert <> "" Then
                                            Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_18599.Art_der_Nutzung_erneuerbaren_Energie_2.Wert, 406, 927, 3878, 3993, Font_Schriftgroesse_40, Font_Druckfarbe)
                                            Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_18599.Deckungsanteil_2.Wert, 1, "", 1426, 1669, 3878, 3993, Font_Schriftgroesse_40, Font_Druckfarbe)
                                        End If
                                        If .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_18599.Art_der_Nutzung_erneuerbaren_Energie_3.Wert <> "" Then
                                            Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_18599.Art_der_Nutzung_erneuerbaren_Energie_3.Wert, 406, 927, 4002, 4109, Font_Schriftgroesse_40, Font_Druckfarbe)
                                            Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_18599.Deckungsanteil_3.Wert, 1, "", 1426, 1669, 4002, 4109, Font_Schriftgroesse_40, Font_Druckfarbe)
                                        End If
                                        '-------------------------------------------------------
                                        Image_Auswahl_schreiben(Grafik_Seite, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_18599.verschaerft_nach_EEWaermeG_7_1_2_eingehalten.Wert, 273, 304, 4557, 4588)
                                        If .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_18599.verschaerft_nach_EEWaermeG_8.Wert > 0 Then
                                            Image_Auswahl_schreiben(Grafik_Seite, True, 273, 304, 4796, 4827)
                                            Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_18599.verschaerft_nach_EEWaermeG_8.Wert, 0, "", 1469, 1665, 4749, 4831, Font_Schriftgroesse_40, Font_Druckfarbe)
                                        Else
                                            Image_Auswahl_schreiben(Grafik_Seite, False, 273, 304, 4796, 4827)
                                        End If
                                        If .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_18599.Primaerenergiebedarf_Hoechstwert_verschaerft.Wert > 0 Then
                                            Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_18599.Primaerenergiebedarf_Hoechstwert_verschaerft.Wert, 1, "", 1213, 1483, 5051, 5135, Font_Schriftgroesse_40, Font_Druckfarbe)
                                        End If
                                        If .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_18599.spezifischer_Transmissionswaermetransferkoeffizient_verschaerft.Wert > 0 Then
                                            Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_18599.spezifischer_Transmissionswaermetransferkoeffizient_verschaerft.Wert, 2, "", 1213, 1483, 5286, 5369, Font_Schriftgroesse_40, Font_Druckfarbe)
                                        End If
                                        '-------------------------------------------------------
                                    Case "Modernisierung-Erweiterung", "Vermietung-Verkauf", "Sonstiges"
                                        '-------------------------------------------------------
                                End Select
                                '-------------------------------------------------------
                            Case "Bedarfsberechnung-easy"
                                '-------------------------------------------------------
                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Treibhausgasemissionen.Wert, 1, "", 3467, 3675, 1128, 1253, Font_Schriftgroesse_40, Font_Druckfarbe)
                                '-------------------------------------------------------
                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_easy.Primaerenergiebedarf_Anforderungswert.Wert, 1, "", 514, 725, 2511, 2615, Font_Schriftgroesse_40, Font_Druckfarbe)
                                '-------------------------------------------------------
                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_easy.Primaerenergiebedarf_Ist_Wert.Wert, 1, "", 1587, 1783, 2511, 2615, Font_Schriftgroesse_40, Font_Druckfarbe)
                                '-------------------------------------------------------
                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_easy.Energetische_Qualitaet_Ist_Wert.Wert, 2, "", 514, 725, 2739, 2851, Font_Schriftgroesse_40, Font_Druckfarbe)
                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_easy.Energetische_Qualitaet_Anforderungs_Wert.Wert, 2, "", 1587, 1783, 2739, 2851, Font_Schriftgroesse_40, Font_Druckfarbe)
                                '-------------------------------------------------------
                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_easy.Endenergiebedarf.Wert, 1, "", 3102, 3633, 3025, 3237, Font_Schriftgroesse_70, Font_Druckfarbe)
                                '-------------------------------------------------------
                                Image_Auswahl_schreiben(Grafik_Seite, False, 1466, 1497, 2915, 2948)
                                '-------------------------------------------------------
                                Image_Auswahl_schreiben(Grafik_Seite, False, 2313, 2344, 2553, 2585)
                                Image_Auswahl_schreiben(Grafik_Seite, False, 2313, 2344, 2667, 2698)
                                Image_Auswahl_schreiben(Grafik_Seite, True, 2313, 2344, 2784, 2815)
                                '-------------------------------------------------------
                                Image_Auswahl_schreiben(Grafik_Seite, False, 2313, 2344, 2915, 2948)
                                '-------------------------------------------------------
                                Image_Pfeil_WB(Grafik_Seite, Lage.Oben, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_easy.Endenergiebedarf.Wert, 1, "Endenergiebedarf dieses Gebäudes", 542, 3623, 1535, 1874)
                                '-------------------------------------------------------
                                Image_Pfeil_WB(Grafik_Seite, Lage.Unten, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_easy.Primaerenergiebedarf_Ist_Wert.Wert, 1, "Primärenergiebedarf dieses Gebäudes", 542, 3623, 1535, 1874)
                                '-------------------------------------------------------
                                'Erneuerbare Energien ----------------------------------
                                '-------------------------------------------------------
                                If .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_easy.Art_der_Nutzung_erneuerbaren_Energie_1.Wert <> "" Then
                                    Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_easy.Art_der_Nutzung_erneuerbaren_Energie_1.Wert, 406, 927, 3748, 3868, Font_Schriftgroesse_40, Font_Druckfarbe)
                                    Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_easy.Deckungsanteil_1.Wert, 1, "", 1426, 1669, 3748, 3868, Font_Schriftgroesse_40, Font_Druckfarbe)
                                End If
                                If .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_easy.Art_der_Nutzung_erneuerbaren_Energie_2.Wert <> "" Then
                                    Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_easy.Art_der_Nutzung_erneuerbaren_Energie_2.Wert, 406, 927, 3878, 3993, Font_Schriftgroesse_40, Font_Druckfarbe)
                                    Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_easy.Deckungsanteil_2.Wert, 1, "", 1426, 1669, 3878, 3993, Font_Schriftgroesse_40, Font_Druckfarbe)
                                End If
                                If .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_easy.Art_der_Nutzung_erneuerbaren_Energie_3.Wert <> "" Then
                                    Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_easy.Art_der_Nutzung_erneuerbaren_Energie_3.Wert, 406, 927, 4002, 4109, Font_Schriftgroesse_40, Font_Druckfarbe)
                                    Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_easy.Deckungsanteil_3.Wert, 1, "", 1426, 1669, 4002, 4109, Font_Schriftgroesse_40, Font_Druckfarbe)
                                End If
                                '-------------------------------------------------------
                        End Select
                        '-------------------------------------------------------
                End Select
                '--------------------------------------------------
            End With
            '-------------------------------------------------------
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
            Hintergrundgrafik = Image.FromFile(Variable_Parameter.Arbeitsverzeichnis & "/Image/ENEV2016_WB_Seite_3.png")
            Grafik_Seite.DrawImage(Hintergrundgrafik, 0, 0, 4200, 5940)
            '-------------------------------------------------------
            With Variable_XML_Import
                '-------------------------------------------------------
                If Variable_Steuerung.Berechnungsart = "Energieverbrauchsausweis" Or Variable_XML_Import.Zusaetzliche_Verbrauchsberechnung = True Then
                    '-------------------------------------------------------
                    Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, Datum_Energiegesetz, 2451, 2913, 447, 544, Font_Schriftgroesse_50, Font_Druckfarbe)
                    Image_Text_schreiben(Grafik_Seite, Ausrichtung.Mitte_kein_Abschneiden, Variable_Steuerung.Registriernummer, 3237, 3699, 645, 743, Font_Bold_Schriftgroesse_40_Narrow, Font_Druckfarbe)
                    '-------------------------------------------------------
                    If .Berechnungsverfahren = "Verbrauchsberechnung" Or Variable_XML_Import.Zusaetzliche_Verbrauchsberechnung = True Then
                        '-------------------------------------------------------
                        Image_Pfeil_WB(Grafik_Seite, Lage.Oben, .Energieausweis_Daten.Wohngebaeude.Verbrauchswerte.Mittlerer_Endenergieverbrauch.Wert, 1, "Endenergieverbrauch dieses Gebäudes", 549, 3631, 1515, 1856)
                        '-------------------------------------------------------
                        Image_Pfeil_WB(Grafik_Seite, Lage.Unten, .Energieausweis_Daten.Wohngebaeude.Verbrauchswerte.Mittlerer_Primaerenergieverbrauch.Wert, 1, "Primärenergieverbrauch dieses Gebäudes", 549, 3631, 1515, 1856)
                        '-------------------------------------------------------
                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Energieausweis_Daten.Wohngebaeude.Verbrauchswerte.Mittlerer_Endenergieverbrauch.Wert, 1, "", 3325, 3645, 2270, 2473, Font_Schriftgroesse_60, Font_Druckfarbe)
                        '-------------------------------------------------------
                        Zusammenstellung_WB_Verbrauchserfassung()
                        '-------------------------------------------------------
                        Zeilenzahl = 6
                        Zeilenhoehe = 130
                        '-------------------------------------------------------
                        If .Verbrauchserfassung_Anzahl < 6 Then
                            Zeilenzahl = .Verbrauchserfassung_Anzahl
                        End If
                        '-------------------------------------------------------
                        For Anzahl_Zeilen = 1 To Zeilenzahl
                            '-------------------------------------------------------
                            Image_Datum_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Verbrauchserfassung(Anzahl_Zeilen).Startdatum.Wert, 242, 644, 2976 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 2976 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                            '--------------------------------------------------
                            Image_Datum_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Verbrauchserfassung(Anzahl_Zeilen).Enddatum.Wert, 653, 1069, 2976 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 2976 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                            '--------------------------------------------------
                            Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligLinks, .Verbrauchserfassung(Anzahl_Zeilen).Energietraeger_Verbrauch.Wert, 1078, 1649, 2976 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 2976 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                            '--------------------------------------------------
                            Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Verbrauchserfassung(Anzahl_Zeilen).Primaerenergiefaktor.Wert, 1, "", 1658, 2034, 2976 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 2976 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                            '--------------------------------------------------
                            If .Verbrauchserfassung(Anzahl_Zeilen).Energieverbrauch.Wert <> 0 Then
                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Verbrauchserfassung(Anzahl_Zeilen).Energieverbrauch.Wert, 0, "", 2043, 2673, 2976 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 2976 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                            End If
                            '--------------------------------------------------
                            If .Verbrauchserfassung(Anzahl_Zeilen).Energieverbrauchsanteil_Warmwasser_zentral.Wert <> 0 Then
                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Verbrauchserfassung(Anzahl_Zeilen).Energieverbrauchsanteil_Warmwasser_zentral.Wert, 0, "", 2682, 3145, 2976 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 2976 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                            End If
                            '--------------------------------------------------
                            If .Verbrauchserfassung(Anzahl_Zeilen).Energieverbrauch.Wert - .Verbrauchserfassung(Anzahl_Zeilen).Energieverbrauchsanteil_Warmwasser_zentral.Wert <> 0 Then
                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Verbrauchserfassung(Anzahl_Zeilen).Energieverbrauch.Wert - .Verbrauchserfassung(Anzahl_Zeilen).Energieverbrauchsanteil_Warmwasser_zentral.Wert, 0, "", 3154, 3637, 2976 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 2976 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                            End If
                            '--------------------------------------------------
                            If .Verbrauchserfassung(Anzahl_Zeilen).Klimafaktor.Wert <> 0 Then
                                Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Verbrauchserfassung(Anzahl_Zeilen).Klimafaktor.Wert, 2, "", 3646, 4034, 2976 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 2976 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                            End If
                            '-------------------------------------------------------
                        Next
                        '-------------------------------------------------------
                    End If
                    '-------------------------------------------------------
                End If
                '--------------------------------------------------
            End With
            '-------------------------------------------------------
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
            Hintergrundgrafik = Image.FromFile(Variable_Parameter.Arbeitsverzeichnis & "/Image/ENEV2016_WB_Seite_4.png")
            Grafik_Seite.DrawImage(Hintergrundgrafik, 0, 0, 4200, 5940)
            '-------------------------------------------------------
            With Variable_XML_Import
                '-------------------------------------------------------
                Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, Datum_Energiegesetz, 2451, 2912, 447, 545, Font_Schriftgroesse_50, Font_Druckfarbe)
                Image_Text_schreiben(Grafik_Seite, Ausrichtung.Mitte_kein_Abschneiden, Variable_Steuerung.Registriernummer, 3237, 3698, 645, 743, Font_Bold_Schriftgroesse_40_Narrow, Font_Druckfarbe)
                '-------------------------------------------------------
                .Modernisierungsempfehlung_Anzahl = Anzahl_Modernisierungsempfehlung()
                '--------------------------------------------------
                If .Energieausweis_Daten.Empfehlungen_moeglich.Wert = False Then
                    '--------------------------------------------------
                    Image_Auswahl_schreiben(Grafik_Seite, False, 2833, 2887, 1197, 1251)
                    Image_Auswahl_schreiben(Grafik_Seite, True, 3452, 3506, 1197, 1251)
                    '--------------------------------------------------
                Else
                    '--------------------------------------------------
                    Image_Auswahl_schreiben(Grafik_Seite, True, 2833, 2887, 1197, 1251)
                    Image_Auswahl_schreiben(Grafik_Seite, False, 3452, 3506, 1197, 1251)
                    '--------------------------------------------------
                    Zeilenzahl = 10
                    Zeilenhoehe = 157
                    '-------------------------------------------------------
                    If .Modernisierungsempfehlung_Anzahl < 10 Then
                        Zeilenzahl = .Modernisierungsempfehlung_Anzahl
                    End If
                    '-------------------------------------------------------
                    For Anzahl_Zeilen = 1 To Zeilenzahl
                        '-------------------------------------------------------
                        Image_Zahl_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, .Modernisierungsempfehlungen(Anzahl_Zeilen).Nummer.Wert, 0, "", 251, 435, 2037 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 2037 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                        '-------------------------------------------------------
                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksMitte, .Modernisierungsempfehlungen(Anzahl_Zeilen).Bauteil_Anlagenteil.Wert, 444 + 20, 1106, 2037 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 2037 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                        '-------------------------------------------------------
                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksMitte, .Modernisierungsempfehlungen(Anzahl_Zeilen).Massnahmenbeschreibung.Wert, 1115 + 20, 2383, 2037 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 2037 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                        '-------------------------------------------------------
                        Select Case .Modernisierungsempfehlungen(Anzahl_Zeilen).Modernisierungskombination.Wert
                            Case "in Zusammenhang mit größerer Modernisierung"
                                Image_Auswahl_schreiben(Grafik_Seite, True, 2625, 2678, 2085 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 2139 + (Zeilenhoehe * (Anzahl_Zeilen - 1)))
                                Image_Auswahl_schreiben(Grafik_Seite, False, 3045, 3099, 2085 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 2139 + (Zeilenhoehe * (Anzahl_Zeilen - 1)))
                            Case "als Einzelmaßnahme"
                                Image_Auswahl_schreiben(Grafik_Seite, False, 2625, 2678, 2085 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 2139 + (Zeilenhoehe * (Anzahl_Zeilen - 1)))
                                Image_Auswahl_schreiben(Grafik_Seite, True, 3045, 3099, 2085 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 2139 + (Zeilenhoehe * (Anzahl_Zeilen - 1)))
                        End Select
                        '-------------------------------------------------------
                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksMitte, .Modernisierungsempfehlungen(Anzahl_Zeilen).Amortisation.Wert, 3231 + 20, 3592, 2037 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 2037 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                        '-------------------------------------------------------
                        Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksMitte, .Modernisierungsempfehlungen(Anzahl_Zeilen).spezifische_Kosten.Wert, 3601 + 20, 4041, 2037 + (Zeilenhoehe * (Anzahl_Zeilen - 1)), 2037 + (Zeilenhoehe * Anzahl_Zeilen), Font_Schriftgroesse_40, Font_Druckfarbe)
                        '-------------------------------------------------------
                    Next
                    '-------------------------------------------------------
                    If .Modernisierungsempfehlung_Anzahl > 10 Then
                        '-------------------------------------------------------
                        Image_Auswahl_schreiben(Grafik_Seite, True, 319, 373, 3644, 3698)
                        '-------------------------------------------------------
                    Else
                        '-------------------------------------------------------
                        Image_Auswahl_schreiben(Grafik_Seite, False, 319, 373, 3644, 3698)
                        '-------------------------------------------------------
                    End If
                    '--------------------------------------------------
                End If
                '--------------------------------------------------
                Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksMitte, .Gebaeudebezogene_Daten.Angaben_erhaeltlich.Wert, 1845 + 20, 4041, 3991, 4291, Font_Schriftgroesse_40, Font_Druckfarbe)
                '--------------------------------------------------
                Image_Text_schreiben(Grafik_Seite, Ausrichtung.MehrzeiligLinksOben, .Gebaeudebezogene_Daten.Ergaenzdende_Erlaeuterungen.Wert, 251 + 20, 4041, 4597, 5584, Font_Schriftgroesse_40, Font_Druckfarbe)
                '--------------------------------------------------
            End With
            '-------------------------------------------------------
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
            Hintergrundgrafik = Image.FromFile(Variable_Parameter.Arbeitsverzeichnis & "/Image/ENEV2016_WB_Seite_5.png")
            Grafik_Seite.DrawImage(Hintergrundgrafik, 0, 0, 4200, 5940)
            '-------------------------------------------------------
            With Variable_XML_Import
                '-------------------------------------------------------
                Image_Text_schreiben(Grafik_Seite, Ausrichtung.EinzeiligMitte, Datum_Energiegesetz, 2451, 2912, 447, 545, Font_Schriftgroesse_50, Font_Druckfarbe)
                '-------------------------------------------------------
            End With
            '-------------------------------------------------------
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
                                Hintergrundgrafik = Image.FromFile(Variable_Parameter.Arbeitsverzeichnis & "/Image/ENEV2016_WB_Ueberlauf_Verbrauch.png")
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
                                For Anzahl_Zeilen = 1 To Wert_Schleife_Zeilenzahl
                                    '-------------------------------------------------------
                                    Image_Linie(Grafik_Seite, 245, 4032, 1366 + (Zeilenhoehe * Anzahl_Zeilen), 1366 + (Zeilenhoehe * Anzahl_Zeilen))
                                    '-------------------------------------------------------
                                Next
                                '-------------------------------------------------------
                                If .Verbrauchserfassung_Anzahl > (Wert_Schleife + Wert_Schleife_Zeilenzahl + Wert_Schleife_Offset) Then
                                    '-------------------------------------------------------
                                    Image_Auswahl_schreiben(Grafik_Seite, True, 300, 330, 5462, 5492)
                                    '-------------------------------------------------------
                                End If
                                '-------------------------------------------------------
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
                        Hintergrundgrafik = Image.FromFile(Variable_Parameter.Arbeitsverzeichnis & "/Image/ENEV2016_WB_Ueberlauf_Modernisierung.png")
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
                        '-------------------------------------------------------
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
