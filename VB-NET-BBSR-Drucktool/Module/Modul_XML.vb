#Region "Imports"
'Imports System.ServiceModel
'Imports System.Xml.Linq
'Imports BBSR_Energieausweis.Modul_Variablen
'Imports System.Collections.Generic
'Imports System.Linq
'Imports System.Diagnostics.Eventing
Imports System.IO
Imports System.Text
Imports System.Xml
Imports System.Xml.Schema
#End Region
'--------------------------------------------------
Module Modul_XML
    '--------------------------------------------------
#Region "XML Steuerungsdatei einlesen/schreiben"
    '--------------------------------------------------
    ''' <summary>
    ''' In dieser Anweisung werden die Inhalte der Steuerungsdatei ausgelesen.
    ''' </summary>
    Sub XML_Steuerungsdatei_einlesen(ByVal Datei As String)
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            'Standardfestlegungen
            Variable_Steuerung.Sandbox = False
            Variable_Steuerung.Entwurf = True
            Variable_Steuerung.Ausgabepfad = Variable_Parameter.Arbeitsverzeichnis
            Variable_Steuerung.Ausgabedatei = "Energieausweis.pdf"
            Variable_Steuerung.PDF_Registriernummer = True
            '--------------------------------------------------
            Dim XMLReader As Xml.XmlReader = New Xml.XmlTextReader(Datei)
            '--------------------------------------------------
            Dim Bezeichnung As String = ""
            Dim Wert As String = ""
            '--------------------------------------------------
            With XMLReader
                '--------------------------------------------------
                Do While .Read
                    '--------------------------------------------------
                    Select Case .NodeType
                        '--------------------------------------------------
                        Case Xml.XmlNodeType.Element
                            '--------------------------------------------------
                            If .Name <> "" Then
                                Bezeichnung = .Name
                            End If
                            '--------------------------------------------------
                        Case Xml.XmlNodeType.Text
                            '--------------------------------------------------
                            If .Value <> "" Then
                                '--------------------------------------------------
                                Wert = LTrim(RTrim(.Value))
                                '--------------------------------------------------
                                Select Case Bezeichnung
                                '--------------------------------------------------
                                    Case "Methode"
                                        Variable_Steuerung.Methode = Wert
                                    Case "Sandbox"
                                        Variable_Steuerung.Sandbox = ImportBool(Wert)
                                    Case "Testdaten"
                                        Variable_Steuerung.Testdaten = ImportBool(Wert)
                                    Case "PDF-erzeugen"
                                        Variable_Steuerung.PDF_erzeugen = ImportBool(Wert)
                                    Case "PDF-oeffnen"
                                        Variable_Steuerung.PDF_oeffnen = ImportBool(Wert)
                                    Case "PDF-Registriernummer"
                                        Variable_Steuerung.PDF_Registriernummer = ImportBool(Wert)
                                    Case "Anwendung-minimieren"
                                        Variable_Steuerung.Anwendung_minimieren = ImportBool(Wert)
                                    Case "Anwendung-beenden"
                                        Variable_Steuerung.Anwendung_beenden = ImportBool(Wert)
                                    '--------------------------------------------------
                                    Case "Registriernummer"
                                        Variable_Steuerung.Registriernummer = Wert
                                    '--------------------------------------------------
                                    Case "Gesetzesgrundlage"
                                        Variable_Steuerung.Gesetzesgrundlage = Wert
                                    Case "Ausstellungsdatum"
                                        Variable_Steuerung.Ausstellungsdatum = ImportDate(Wert)
                                    Case "Bundesland"
                                        Variable_Steuerung.Bundesland = Wert
                                    Case "Postleitzahl"
                                        Variable_Steuerung.Postleitzahl = Wert
                                    Case "Gebaeudeart"
                                        Variable_Steuerung.Gebaeudeart = Wert
                                    Case "Berechnungsart"
                                        Variable_Steuerung.Berechnungsart = Wert
                                    Case "Neubau"
                                        Variable_Steuerung.Neubau = ImportInt(Wert)
                                    Case "Importdatei"
                                        Variable_Steuerung.Importdatei = Wert
                                    Case "Ausgabepfad"
                                        Variable_Steuerung.Ausgabepfad = Wert
                                    Case "Ausgabedatei"
                                        Variable_Steuerung.Ausgabedatei = Wert
                                    Case "Ausgabequalitaet"
                                        Variable_Steuerung.Ausgabequalitaet = ImportInt(Wert)
                                    Case "Bild-Projekt"
                                        Variable_Steuerung.Bild_Projekt = Wert
                                    Case "Bild-Aussteller"
                                        Variable_Steuerung.Bild_Aussteller = Wert
                                    Case "Bild-Unterschrift"
                                        Variable_Steuerung.Bild_Unterschrift = Wert
                                End Select
                                '--------------------------------------------------
                            End If
                            '--------------------------------------------------
                    End Select
                    '--------------------------------------------------
                Loop
                '--------------------------------------------------
                .Close()
                '--------------------------------------------------
            End With
            '--------------------------------------------------
        Catch ex As Exception
            Fehlerfenster(ex)
        End Try
        '--------------------------------------------------
    End Sub
    '--------------------------------------------------
    ''' <summary>
    ''' In dieser Anweisung werden die Ergebnisse am Ende in die Steuerungsdatei geschrieben.
    ''' </summary>
    Sub XML_Steuerungsdatei_schreiben(ByVal Datei As String)
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            Dim WertX As Integer
            '--------------------------------------------------
            Dim Xw As New XmlTextWriter(Datei, Encoding.UTF8)
            '---------------------------------------------------------------------------------------------------
            Xw.WriteStartDocument()
            '---------------------------------------------------------------------------------------------------
            Xw.WriteStartElement("root")
            '---------------------------------------------------------------------------------------------------
            '###################################################################################################
            '---------------------------------------------------------------------------------------------------
            Xw.WriteStartElement("Daten")
            '---------------------------------------------------------------------------------------------------
            XML_Schreiben(Xw, "Methode", Variable_Steuerung.Methode)
            '---------------------------------------------------------------------------------------------------
            XML_Schreiben(Xw, "Gesetzesgrundlage", Variable_Steuerung.Gesetzesgrundlage)
            XML_Schreiben(Xw, "Ausstellungsdatum", Variable_Steuerung.Ausstellungsdatum)
            XML_Schreiben(Xw, "Bundesland", Variable_Steuerung.Bundesland)
            XML_Schreiben(Xw, "Postleitzahl", Variable_Steuerung.Postleitzahl)
            XML_Schreiben(Xw, "Gebaeudeart", Variable_Steuerung.Gebaeudeart)
            XML_Schreiben(Xw, "Berechnungsart", Variable_Steuerung.Berechnungsart)
            XML_Schreiben(Xw, "Neubau", Variable_Steuerung.Neubau)
            '---------------------------------------------------------------------------------------------------
            XML_Schreiben(Xw, "Importdatei", Variable_Steuerung.Importdatei)
            XML_Schreiben(Xw, "Ausgabepfad", Variable_Steuerung.Ausgabepfad)
            XML_Schreiben(Xw, "Ausgabedatei", Variable_Steuerung.Ausgabedatei)
            '---------------------------------------------------------------------------------------------------
            XML_Schreiben(Xw, "Bild-Projekt", Variable_Steuerung.Bild_Projekt)
            XML_Schreiben(Xw, "Bild-Aussteller", Variable_Steuerung.Bild_Aussteller)
            XML_Schreiben(Xw, "Bild-Unterschrift", Variable_Steuerung.Bild_Unterschrift)
            '---------------------------------------------------------------------------------------------------
            Xw.WriteEndElement() 'Daten
            '---------------------------------------------------------------------------------------------------
            '###################################################################################################
            '---------------------------------------------------------------------------------------------------
            Xw.WriteStartElement("Datenregistratur")
            '---------------------------------------------------------------------------------------------------
            If Variable_Steuerung.Datenregistratur_Fehler_Anzahl = 0 Then
                '---------------------------------------------------------------------------------------------------
                XML_Schreiben(Xw, "Registriernummer", Variable_Steuerung.Registriernummer)
                '---------------------------------------------------------------------------------------------------
                If Variable_Steuerung.Datenregistratur = True Then
                    '---------------------------------------------------------------------------------------------------
                    XML_Schreiben(Xw, "Datenregistratur-Ergebnis", True)
                    '---------------------------------------------------------------------------------------------------
                    XML_Schreiben(Xw, "Restkontingent", Variable_Steuerung.Restkontingent_Anzahl)
                    '---------------------------------------------------------------------------------------------------
                    If Variable_Steuerung.Datenregistratur_Pruefung_Datendatei = 1 Then
                        '---------------------------------------------------------------------------------------------------
                        XML_Schreiben(Xw, "Datenregistratur-Pruefung-Datendatei", Variable_Steuerung.Datenregistratur_Pruefung_Datendatei)
                        XML_Schreiben(Xw, "Datenregistratur-Pruefung-Datendatei-Bemerkungen", Variable_Steuerung.Datenregistratur_Pruefung_Datendatei_Bemerkungen)
                        '---------------------------------------------------------------------------------------------------
                    End If
                    '---------------------------------------------------------------------------------------------------
                End If
                '---------------------------------------------------------------------------------------------------
            Else
                '---------------------------------------------------------------------------------------------------
                XML_Schreiben(Xw, "Datenregistratur-Ergebnis", False)
                '---------------------------------------------------------------------------------------------------
                Xw.WriteStartElement("Datenregistratur-Fehlerliste")
                '---------------------------------------------------------------------------------------------------
                For WertX = 0 To (Variable_Steuerung.Datenregistratur_Fehler_Anzahl - 1)
                    '---------------------------------------------------------------------------------------------------
                    XML_Schreiben(Xw, "Datenregistratur-Fehler", Variable_Steuerung.Datenregistratur_Fehler_Text(WertX), "ID", Variable_Steuerung.Datenregistratur_Fehler_ID(WertX))
                    '---------------------------------------------------------------------------------------------------
                Next
                '---------------------------------------------------------------------------------------------------
                Xw.WriteEndElement() 'Datenregistratur-Fehlerliste
                '---------------------------------------------------------------------------------------------------
            End If
            '---------------------------------------------------------------------------------------------------
            Xw.WriteEndElement() 'Datenregistratur
            '---------------------------------------------------------------------------------------------------
            '###################################################################################################
            '---------------------------------------------------------------------------------------------------
            If Variable_Steuerung.KontrolldateiPruefen = True Then
                '---------------------------------------------------------------------------------------------------
                Xw.WriteStartElement("KontrolldateiPruefen")
                '---------------------------------------------------------------------------------------------------
                If Variable_Steuerung.KontrolldateiPruefen_Fehler_Anzahl = 0 Then
                    '---------------------------------------------------------------------------------------------------
                    XML_Schreiben(Xw, "KontrolldateiPruefen-Ergebnis", True)
                    '---------------------------------------------------------------------------------------------------
                Else
                    '---------------------------------------------------------------------------------------------------
                    XML_Schreiben(Xw, "KontrolldateiPruefen-Ergebnis", False)
                    '---------------------------------------------------------------------------------------------------
                    Xw.WriteStartElement("KontrolldateiPruefen-Fehlerliste")
                    '---------------------------------------------------------------------------------------------------
                    For WertX = 0 To (Variable_Steuerung.KontrolldateiPruefen_Fehler_Anzahl - 1)
                        '---------------------------------------------------------------------------------------------------
                        XML_Schreiben(Xw, "KontrolldateiPruefen-Fehler", Variable_Steuerung.KontrolldateiPruefen_Fehler_Langtext(WertX), "ID", Variable_Steuerung.KontrolldateiPruefen_Fehler_ID(WertX), "Kurztext", Variable_Steuerung.KontrolldateiPruefen_Fehler_Kurztext(WertX))
                        '---------------------------------------------------------------------------------------------------
                    Next
                    '---------------------------------------------------------------------------------------------------
                    Xw.WriteEndElement() 'KontrolldateiPruefen-Fehlerliste
                    '---------------------------------------------------------------------------------------------------
                End If
                '---------------------------------------------------------------------------------------------------
                Xw.WriteEndElement() 'KontrolldateiPruefen
                '---------------------------------------------------------------------------------------------------
            End If
            '---------------------------------------------------------------------------------------------------
            '###################################################################################################
            '---------------------------------------------------------------------------------------------------
            If Variable_Steuerung.Restkontingent = True Then
                '---------------------------------------------------------------------------------------------------
                Xw.WriteStartElement("Restkontingent")
                '---------------------------------------------------------------------------------------------------
                If Variable_Steuerung.Restkontingent_Fehlermeldung = "" Then
                    '---------------------------------------------------------------------------------------------------
                    XML_Schreiben(Xw, "Restkontingent-Ergebnis", True)
                    '---------------------------------------------------------------------------------------------------
                    XML_Schreiben(Xw, "Restkontingent", Variable_Steuerung.Restkontingent_Anzahl)
                    '---------------------------------------------------------------------------------------------------
                Else
                    '---------------------------------------------------------------------------------------------------
                    XML_Schreiben(Xw, "Restkontingent-Ergebnis", False)
                    '---------------------------------------------------------------------------------------------------
                    XML_Schreiben(Xw, "Restkontingent-Fehlermeldung", Variable_Steuerung.Restkontingent_Fehlermeldung)
                    '---------------------------------------------------------------------------------------------------
                End If
                '---------------------------------------------------------------------------------------------------
                Xw.WriteEndElement() 'Restkontingent
                '---------------------------------------------------------------------------------------------------
            End If
            '---------------------------------------------------------------------------------------------------
            '###################################################################################################
            '---------------------------------------------------------------------------------------------------
            If Variable_Steuerung.ZusatzdatenErfassung = True Then
                '---------------------------------------------------------------------------------------------------
                Xw.WriteStartElement("Zusatzdatenerfassung")
                '---------------------------------------------------------------------------------------------------
                If Variable_Steuerung.ZusatzdatenErfassung_Fehler_Anzahl = 0 Then
                    '---------------------------------------------------------------------------------------------------
                    XML_Schreiben(Xw, "Zusatzdatenerfassung-Ergebnis", True)
                    '---------------------------------------------------------------------------------------------------
                Else
                    '---------------------------------------------------------------------------------------------------
                    XML_Schreiben(Xw, "Zusatzdatenerfassung-Ergebnis", False)
                    '---------------------------------------------------------------------------------------------------
                    Xw.WriteStartElement("Zusatzdatenerfassung-Fehlerliste")
                    '---------------------------------------------------------------------------------------------------
                    For WertX = 0 To (Variable_Steuerung.ZusatzdatenErfassung_Fehler_Anzahl - 1)
                        '---------------------------------------------------------------------------------------------------
                        XML_Schreiben(Xw, "Zusatzdatenerfassung-Fehler", Variable_Steuerung.ZusatzdatenErfassung_Fehler_Text(WertX), "ID", Variable_Steuerung.ZusatzdatenErfassung_Fehler_ID(WertX))
                        '---------------------------------------------------------------------------------------------------
                    Next
                    '---------------------------------------------------------------------------------------------------
                    Xw.WriteEndElement() 'Zusatzdatenerfassung-Fehlerliste
                    '---------------------------------------------------------------------------------------------------
                End If
                '---------------------------------------------------------------------------------------------------
                Xw.WriteEndElement() 'Zusatzdatenerfassung
                '---------------------------------------------------------------------------------------------------
            End If
            '---------------------------------------------------------------------------------------------------
            '###################################################################################################
            '---------------------------------------------------------------------------------------------------
            If Variable_Steuerung.OffeneKontrolldateien = True Then
                '---------------------------------------------------------------------------------------------------
                Xw.WriteStartElement("Offenekontrolldateien")
                '---------------------------------------------------------------------------------------------------
                If Variable_Steuerung.OffeneKontrolldateien_Fehlermeldung = "" Then
                    '---------------------------------------------------------------------------------------------------
                    XML_Schreiben(Xw, "Offenekontrolldateien-Ergebnis", True)
                    '---------------------------------------------------------------------------------------------------
                    XML_Schreiben(Xw, "Offenekontrolldateien-Anzahl", Variable_Steuerung.OffeneKontrolldateien_Anzahl)
                    '---------------------------------------------------------------------------------------------------
                    Xw.WriteStartElement("Offenekontrolldateien-Ausweise")
                    '---------------------------------------------------------------------------------------------------
                    For WertX = 0 To (Variable_Steuerung.OffeneKontrolldateien_Anzahl - 1)
                        '--------------------------------------------------
                        XML_Schreiben(Xw, "Offenekontrolldateien-Ausweis", Variable_Steuerung.OffeneKontrolldateien_Ausweis_Registriernummer(WertX), "NummerErzeugtAm", Variable_Steuerung.OffeneKontrolldateien_Ausweis_NummerErzeugtAm(WertX), "Aussteller", Variable_Steuerung.OffeneKontrolldateien_Ausweis_Aussteller(WertX))
                        '--------------------------------------------------
                    Next
                    '---------------------------------------------------------------------------------------------------
                    Xw.WriteEndElement() 'Offenekontrolldateien-Ausweise
                    '---------------------------------------------------------------------------------------------------
                Else
                    '---------------------------------------------------------------------------------------------------
                    XML_Schreiben(Xw, "Offenekontrolldateien-Ergebnis", False)
                    '---------------------------------------------------------------------------------------------------
                    XML_Schreiben(Xw, "Offenekontrolldateien-Fehlermeldung", Variable_Steuerung.OffeneKontrolldateien_Fehlermeldung)
                    '---------------------------------------------------------------------------------------------------
                End If
                '---------------------------------------------------------------------------------------------------
                Xw.WriteEndElement() 'Offenekontrolldateien
                '---------------------------------------------------------------------------------------------------
            End If
            '---------------------------------------------------------------------------------------------------
            '###################################################################################################
            '---------------------------------------------------------------------------------------------------
            Xw.WriteStartElement("Ergebnisprotokoll")
            '---------------------------------------------------------------------------------------------------
            XML_Schreiben(Xw, "Ergebnisprotokoll-Base64", Variable_Steuerung.Ergebnisprotokoll_Base64)
            '---------------------------------------------------------------------------------------------------
            Xw.WriteEndElement() 'Ergebnisprotokoll
            '---------------------------------------------------------------------------------------------------
            '###################################################################################################
            '---------------------------------------------------------------------------------------------------
            Xw.WriteEndElement() 'root
            '---------------------------------------------------------------------------------------------------
            Xw.WriteEndDocument()
            '---------------------------------------------------------------------------------------------------
            Xw.Close()
            '--------------------------------------------------
        Catch ex As Exception
            Fehlerfenster(ex)
        End Try
        '--------------------------------------------------
    End Sub
    '--------------------------------------------------
#End Region
    '--------------------------------------------------
#Region "XML Kontrolldatei einlesen"
    ''' <summary>
    ''' In dieser Anweisung wird die Importdatei ausgelesen. Es werden alle Werte ausgelesen.
    ''' </summary>
    Sub XML_Kontrolldatei_einlesen(ByVal Datei As String)
        '--------------------------------------------------
        Dim Datensatz_Typ As String = ""
        Dim Datensatz_Bezeichnung As String = ""
        Dim Datensatz_Wert As String = ""
        '--------------------------------------------------
        Dim Datensatz_Ebene_1 As String = ""
        Dim Datensatz_Ebene_2 As String = ""
        Dim Datensatz_Ebene_3 As String = ""
        Dim Datensatz_Ebene_4 As String = ""
        Dim Datensatz_Ebene_5 As String = ""
        '--------------------------------------------------
        Dim Datensatz_Opak_Nr As Integer = 0
        Dim Datensatz_Transparent_Nr As Integer = 0
        Dim Datensatz_Dach_Nr As Integer = 0
        Dim Datensatz_Zone_Nr As Integer = 0
        Dim Datensatz_Heizungsanlage_Nr As Integer = 0
        Dim Datensatz_Trinkwarmwasseranlage_Nr As Integer = 0
        Dim Datensatz_Lueftung_Nr As Integer = 0
        Dim Datensatz_Klima_Nr As Integer = 0
        Dim Datensatz_Energietraeger_Nr As Integer = 0
        Dim Datensatz_erneuerbare_Energien_65EE_Regel_Nr As Integer = 0
        Dim Datensatz_erneuerbare_Energien_65EE_keine_Regel_Nr As Integer = 0
        Dim Datensatz_Nutzung_Gebaeudekategorie_Nr As Integer = 0
        Dim Datensatz_Energietraeger_Daten_Nr As Integer = 0
        Dim Datensatz_Modernisierungsempfehlungen_Nr As Integer = 0
        '--------------------------------------------------
        Dim Datensatz_Zeitraum_Daten_Nr As Integer = 0
        Dim Datensatz_Zeitraum_Strom_Daten_Nr As Integer = 0
        '--------------------------------------------------
        Try
            '-----------------------------------------------------
            Dim Reader As Xml.XmlReader = New Xml.XmlTextReader(Datei)
            While Reader.Read()
                If Reader.IsStartElement() Then
                    If Reader.Prefix = String.Empty Then
                        Variable_Steuerung.XML_Kontrolldatei_Prefix = ""
                    Else
                        Variable_Steuerung.XML_Kontrolldatei_Prefix = Reader.Prefix
                    End If
                End If
            End While
            Reader.Close()
            '-----------------------------------------------------
            Dim XMLReader As Xml.XmlReader = New Xml.XmlTextReader(Datei)
            '----------------------------------------------------- 
            With XMLReader
                '-----------------------------------------------------
                Do While .Read
                    '--------------------------------------------------
                    Datensatz_Bezeichnung = .LocalName
                    '--------------------------------------------------
                    Select Case .NodeType
                        '-----------------------------------------------------
                        Case Xml.XmlNodeType.Element
                            Select Case Datensatz_Bezeichnung
                                Case "Gebaeudebezogene-Daten", "Energieausweis-Daten"
                                    '--------------------------------------------------
                                    Datensatz_Ebene_1 = Datensatz_Bezeichnung
                                    '--------------------------------------------------
                                Case "Wohngebaeude", "Nichtwohngebaeude", "NWG-Diagramm-Daten"
                                    '--------------------------------------------------
                                    Datensatz_Ebene_2 = Datensatz_Bezeichnung
                                    '--------------------------------------------------
                                Case "Bedarfswerte-4108-4701", "Bedarfsdaten-4108-4701", "Bedarfswerte-easy", "Bedarfswerte-18599", "Verbrauchswerte", "Verbrauchsdaten", "Bedarfswerte-NWG", "Verbrauchswerte-NWG", "Zusaetzliche-Verbrauchsdaten"
                                    '--------------------------------------------------
                                    Datensatz_Ebene_3 = Datensatz_Bezeichnung
                                    '--------------------------------------------------
                                Case "Bauteil-Opak"
                                    '--------------------------------------------------
                                    Datensatz_Ebene_4 = Datensatz_Bezeichnung
                                    '--------------------------------------------------
                                    Datensatz_Opak_Nr += 1
                                    Datensatz_Transparent_Nr = 0
                                    Datensatz_Dach_Nr = 0
                                    Datensatz_Zone_Nr = 0
                                    Datensatz_Heizungsanlage_Nr = 0
                                    Datensatz_Trinkwarmwasseranlage_Nr = 0
                                    Datensatz_Lueftung_Nr = 0
                                    Datensatz_Klima_Nr = 0
                                    Datensatz_Energietraeger_Nr = 0
                                    Datensatz_erneuerbare_Energien_65EE_Regel_Nr = 0
                                    Datensatz_erneuerbare_Energien_65EE_keine_Regel_Nr = 0
                                    Datensatz_Energietraeger_Daten_Nr = 0
                                    Datensatz_Zeitraum_Strom_Daten_Nr = 0
                                    Datensatz_Nutzung_Gebaeudekategorie_Nr = 0
                                    Datensatz_Modernisierungsempfehlungen_Nr = 0
                                    '--------------------------------------------------
                                Case "Bauteil-Transparent"
                                    '--------------------------------------------------
                                    Datensatz_Ebene_4 = Datensatz_Bezeichnung
                                    '--------------------------------------------------
                                    Datensatz_Opak_Nr = 0
                                    Datensatz_Transparent_Nr += 1
                                    Datensatz_Dach_Nr = 0
                                    Datensatz_Zone_Nr = 0
                                    Datensatz_Heizungsanlage_Nr = 0
                                    Datensatz_Trinkwarmwasseranlage_Nr = 0
                                    Datensatz_Lueftung_Nr = 0
                                    Datensatz_Klima_Nr = 0
                                    Datensatz_Energietraeger_Nr = 0
                                    Datensatz_erneuerbare_Energien_65EE_Regel_Nr = 0
                                    Datensatz_erneuerbare_Energien_65EE_keine_Regel_Nr = 0
                                    Datensatz_Energietraeger_Daten_Nr = 0
                                    Datensatz_Zeitraum_Strom_Daten_Nr = 0
                                    Datensatz_Nutzung_Gebaeudekategorie_Nr = 0
                                    Datensatz_Modernisierungsempfehlungen_Nr = 0
                                    '--------------------------------------------------
                                Case "Bauteil-Dach"
                                    '--------------------------------------------------
                                    Datensatz_Ebene_4 = Datensatz_Bezeichnung
                                    '--------------------------------------------------
                                    Datensatz_Opak_Nr = 0
                                    Datensatz_Transparent_Nr = 0
                                    Datensatz_Dach_Nr += 1
                                    Datensatz_Zone_Nr = 0
                                    Datensatz_Heizungsanlage_Nr = 0
                                    Datensatz_Trinkwarmwasseranlage_Nr = 0
                                    Datensatz_Lueftung_Nr = 0
                                    Datensatz_Klima_Nr = 0
                                    Datensatz_Energietraeger_Nr = 0
                                    Datensatz_erneuerbare_Energien_65EE_Regel_Nr = 0
                                    Datensatz_erneuerbare_Energien_65EE_keine_Regel_Nr = 0
                                    Datensatz_Energietraeger_Daten_Nr = 0
                                    Datensatz_Zeitraum_Strom_Daten_Nr = 0
                                    Datensatz_Nutzung_Gebaeudekategorie_Nr = 0
                                    Datensatz_Modernisierungsempfehlungen_Nr = 0
                                    '--------------------------------------------------
                                Case "Zone"
                                    '--------------------------------------------------
                                    Datensatz_Ebene_4 = Datensatz_Bezeichnung
                                    '--------------------------------------------------
                                    Datensatz_Opak_Nr = 0
                                    Datensatz_Transparent_Nr = 0
                                    Datensatz_Dach_Nr = 0
                                    Datensatz_Zone_Nr += 1
                                    Datensatz_Heizungsanlage_Nr = 0
                                    Datensatz_Trinkwarmwasseranlage_Nr = 0
                                    Datensatz_Lueftung_Nr = 0
                                    Datensatz_Klima_Nr = 0
                                    Datensatz_Energietraeger_Nr = 0
                                    Datensatz_erneuerbare_Energien_65EE_Regel_Nr = 0
                                    Datensatz_erneuerbare_Energien_65EE_keine_Regel_Nr = 0
                                    Datensatz_Energietraeger_Daten_Nr = 0
                                    Datensatz_Zeitraum_Strom_Daten_Nr = 0
                                    Datensatz_Nutzung_Gebaeudekategorie_Nr = 0
                                    Datensatz_Modernisierungsempfehlungen_Nr = 0
                                    '--------------------------------------------------
                                Case "Heizungsanlage", "Heizsystem"
                                    '--------------------------------------------------
                                    Datensatz_Ebene_4 = Datensatz_Bezeichnung
                                    '--------------------------------------------------
                                    Datensatz_Opak_Nr = 0
                                    Datensatz_Transparent_Nr = 0
                                    Datensatz_Dach_Nr = 0
                                    Datensatz_Zone_Nr = 0
                                    Datensatz_Heizungsanlage_Nr += 1
                                    Datensatz_Trinkwarmwasseranlage_Nr = 0
                                    Datensatz_Lueftung_Nr = 0
                                    Datensatz_Klima_Nr = 0
                                    Datensatz_Energietraeger_Nr = 0
                                    Datensatz_erneuerbare_Energien_65EE_Regel_Nr = 0
                                    Datensatz_erneuerbare_Energien_65EE_keine_Regel_Nr = 0
                                    Datensatz_Energietraeger_Daten_Nr = 0
                                    Datensatz_Zeitraum_Strom_Daten_Nr = 0
                                    Datensatz_Nutzung_Gebaeudekategorie_Nr = 0
                                    Datensatz_Modernisierungsempfehlungen_Nr = 0
                                    '--------------------------------------------------
                                Case "Trinkwarmwasseranlage", "Warmwasserbereitungssystem"
                                    '--------------------------------------------------
                                    Datensatz_Ebene_4 = Datensatz_Bezeichnung
                                    '--------------------------------------------------
                                    Datensatz_Opak_Nr = 0
                                    Datensatz_Transparent_Nr = 0
                                    Datensatz_Dach_Nr = 0
                                    Datensatz_Zone_Nr = 0
                                    Datensatz_Heizungsanlage_Nr = 0
                                    Datensatz_Trinkwarmwasseranlage_Nr += 1
                                    Datensatz_Lueftung_Nr = 0
                                    Datensatz_Klima_Nr = 0
                                    Datensatz_Energietraeger_Nr = 0
                                    Datensatz_erneuerbare_Energien_65EE_Regel_Nr = 0
                                    Datensatz_erneuerbare_Energien_65EE_keine_Regel_Nr = 0
                                    Datensatz_Energietraeger_Daten_Nr = 0
                                    Datensatz_Zeitraum_Strom_Daten_Nr = 0
                                    Datensatz_Nutzung_Gebaeudekategorie_Nr = 0
                                    Datensatz_Modernisierungsempfehlungen_Nr = 0
                                    '--------------------------------------------------
                                Case "Kaelteanlage"
                                    '--------------------------------------------------
                                    Datensatz_Ebene_4 = Datensatz_Bezeichnung
                                    '--------------------------------------------------
                                    Datensatz_Opak_Nr = 0
                                    Datensatz_Transparent_Nr = 0
                                    Datensatz_Dach_Nr = 0
                                    Datensatz_Zone_Nr = 0
                                    Datensatz_Heizungsanlage_Nr = 0
                                    Datensatz_Trinkwarmwasseranlage_Nr = 0
                                    Datensatz_Lueftung_Nr = 0
                                    Datensatz_Klima_Nr += 1
                                    Datensatz_Energietraeger_Nr = 0
                                    Datensatz_erneuerbare_Energien_65EE_Regel_Nr = 0
                                    Datensatz_erneuerbare_Energien_65EE_keine_Regel_Nr = 0
                                    Datensatz_Energietraeger_Daten_Nr = 0
                                    Datensatz_Zeitraum_Strom_Daten_Nr = 0
                                    Datensatz_Nutzung_Gebaeudekategorie_Nr = 0
                                    Datensatz_Modernisierungsempfehlungen_Nr = 0
                                    '--------------------------------------------------
                                Case "RLT-System"
                                    '--------------------------------------------------
                                    Datensatz_Ebene_4 = Datensatz_Bezeichnung
                                    '--------------------------------------------------
                                    Datensatz_Opak_Nr = 0
                                    Datensatz_Transparent_Nr = 0
                                    Datensatz_Dach_Nr = 0
                                    Datensatz_Zone_Nr = 0
                                    Datensatz_Heizungsanlage_Nr = 0
                                    Datensatz_Trinkwarmwasseranlage_Nr = 0
                                    Datensatz_Lueftung_Nr += 1
                                    Datensatz_Klima_Nr = 0
                                    Datensatz_Energietraeger_Nr = 0
                                    Datensatz_erneuerbare_Energien_65EE_Regel_Nr = 0
                                    Datensatz_erneuerbare_Energien_65EE_keine_Regel_Nr = 0
                                    Datensatz_Energietraeger_Daten_Nr = 0
                                    Datensatz_Zeitraum_Strom_Daten_Nr = 0
                                    Datensatz_Nutzung_Gebaeudekategorie_Nr = 0
                                    Datensatz_Modernisierungsempfehlungen_Nr = 0
                                    '--------------------------------------------------
                                Case "Energietraeger-Liste"
                                    '--------------------------------------------------
                                    Datensatz_Ebene_4 = Datensatz_Bezeichnung
                                    '--------------------------------------------------
                                    Datensatz_Opak_Nr = 0
                                    Datensatz_Transparent_Nr = 0
                                    Datensatz_Dach_Nr = 0
                                    Datensatz_Zone_Nr = 0
                                    Datensatz_Heizungsanlage_Nr = 0
                                    Datensatz_Trinkwarmwasseranlage_Nr = 0
                                    Datensatz_Lueftung_Nr = 0
                                    Datensatz_Klima_Nr = 0
                                    Datensatz_Energietraeger_Nr += 1
                                    Datensatz_erneuerbare_Energien_65EE_Regel_Nr = 0
                                    Datensatz_erneuerbare_Energien_65EE_keine_Regel_Nr = 0
                                    Datensatz_Energietraeger_Daten_Nr = 0
                                    Datensatz_Zeitraum_Strom_Daten_Nr = 0
                                    Datensatz_Nutzung_Gebaeudekategorie_Nr = 0
                                    Datensatz_Modernisierungsempfehlungen_Nr = 0
                                    '--------------------------------------------------
                                Case "Angaben-erneuerbare-Energien-65EE-Regel"
                                    '--------------------------------------------------
                                    Datensatz_Ebene_4 = Datensatz_Bezeichnung
                                    '--------------------------------------------------
                                    Datensatz_Opak_Nr = 0
                                    Datensatz_Transparent_Nr = 0
                                    Datensatz_Dach_Nr = 0
                                    Datensatz_Zone_Nr = 0
                                    Datensatz_Heizungsanlage_Nr = 0
                                    Datensatz_Trinkwarmwasseranlage_Nr = 0
                                    Datensatz_Lueftung_Nr = 0
                                    Datensatz_Klima_Nr = 0
                                    Datensatz_Energietraeger_Nr = 0
                                    Datensatz_erneuerbare_Energien_65EE_Regel_Nr += 1
                                    Datensatz_erneuerbare_Energien_65EE_keine_Regel_Nr = 0
                                    Datensatz_Energietraeger_Daten_Nr = 0
                                    Datensatz_Zeitraum_Strom_Daten_Nr = 0
                                    Datensatz_Nutzung_Gebaeudekategorie_Nr = 0
                                    Datensatz_Modernisierungsempfehlungen_Nr = 0
                                    '--------------------------------------------------
                                Case "Angaben-erneuerbare-Energien-keine-65EE-Regel"
                                    '--------------------------------------------------
                                    Datensatz_Ebene_4 = Datensatz_Bezeichnung
                                    '--------------------------------------------------
                                    Datensatz_Opak_Nr = 0
                                    Datensatz_Transparent_Nr = 0
                                    Datensatz_Dach_Nr = 0
                                    Datensatz_Zone_Nr = 0
                                    Datensatz_Heizungsanlage_Nr = 0
                                    Datensatz_Trinkwarmwasseranlage_Nr = 0
                                    Datensatz_Lueftung_Nr = 0
                                    Datensatz_Klima_Nr = 0
                                    Datensatz_Energietraeger_Nr = 0
                                    Datensatz_erneuerbare_Energien_65EE_Regel_Nr = 0
                                    Datensatz_erneuerbare_Energien_65EE_keine_Regel_Nr += 1
                                    Datensatz_Energietraeger_Daten_Nr = 0
                                    Datensatz_Zeitraum_Strom_Daten_Nr = 0
                                    Datensatz_Nutzung_Gebaeudekategorie_Nr = 0
                                    Datensatz_Modernisierungsempfehlungen_Nr = 0
                                    '--------------------------------------------------
                                Case "Energietraeger"
                                    '--------------------------------------------------
                                    Datensatz_Ebene_4 = Datensatz_Bezeichnung
                                    '--------------------------------------------------
                                    Datensatz_Opak_Nr = 0
                                    Datensatz_Transparent_Nr = 0
                                    Datensatz_Dach_Nr = 0
                                    Datensatz_Zone_Nr = 0
                                    Datensatz_Heizungsanlage_Nr = 0
                                    Datensatz_Trinkwarmwasseranlage_Nr = 0
                                    Datensatz_Lueftung_Nr = 0
                                    Datensatz_Klima_Nr = 0
                                    Datensatz_Energietraeger_Nr = 0
                                    Datensatz_erneuerbare_Energien_65EE_Regel_Nr = 0
                                    Datensatz_erneuerbare_Energien_65EE_keine_Regel_Nr = 0
                                    Datensatz_Energietraeger_Daten_Nr += 1
                                    Datensatz_Zeitraum_Strom_Daten_Nr = 0
                                    Datensatz_Nutzung_Gebaeudekategorie_Nr = 0
                                    Datensatz_Modernisierungsempfehlungen_Nr = 0
                                    '--------------------------------------------------
                                    Datensatz_Zeitraum_Daten_Nr = 0
                                    '--------------------------------------------------
                                Case "Zeitraum-Strom"
                                    '--------------------------------------------------
                                    Datensatz_Ebene_4 = Datensatz_Bezeichnung
                                    '--------------------------------------------------
                                    Datensatz_Opak_Nr = 0
                                    Datensatz_Transparent_Nr = 0
                                    Datensatz_Dach_Nr = 0
                                    Datensatz_Zone_Nr = 0
                                    Datensatz_Heizungsanlage_Nr = 0
                                    Datensatz_Trinkwarmwasseranlage_Nr = 0
                                    Datensatz_Lueftung_Nr = 0
                                    Datensatz_Klima_Nr = 0
                                    Datensatz_Energietraeger_Nr = 0
                                    Datensatz_erneuerbare_Energien_65EE_Regel_Nr = 0
                                    Datensatz_erneuerbare_Energien_65EE_keine_Regel_Nr = 0
                                    Datensatz_Energietraeger_Daten_Nr = 0
                                    Datensatz_Zeitraum_Strom_Daten_Nr += 1
                                    Datensatz_Nutzung_Gebaeudekategorie_Nr = 0
                                    Datensatz_Modernisierungsempfehlungen_Nr = 0
                                    '--------------------------------------------------
                                Case "Nutzung-Gebaeudekategorie"
                                    '--------------------------------------------------
                                    Datensatz_Ebene_4 = Datensatz_Bezeichnung
                                    '--------------------------------------------------
                                    Datensatz_Opak_Nr = 0
                                    Datensatz_Transparent_Nr = 0
                                    Datensatz_Dach_Nr = 0
                                    Datensatz_Zone_Nr = 0
                                    Datensatz_Heizungsanlage_Nr = 0
                                    Datensatz_Trinkwarmwasseranlage_Nr = 0
                                    Datensatz_Lueftung_Nr = 0
                                    Datensatz_Klima_Nr = 0
                                    Datensatz_Energietraeger_Nr = 0
                                    Datensatz_erneuerbare_Energien_65EE_Regel_Nr = 0
                                    Datensatz_erneuerbare_Energien_65EE_keine_Regel_Nr = 0
                                    Datensatz_Energietraeger_Daten_Nr = 0
                                    Datensatz_Zeitraum_Strom_Daten_Nr = 0
                                    Datensatz_Nutzung_Gebaeudekategorie_Nr += 1
                                    Datensatz_Modernisierungsempfehlungen_Nr = 0
                                    '--------------------------------------------------
                                Case "Modernisierungsempfehlungen"
                                    '--------------------------------------------------
                                    Datensatz_Ebene_4 = Datensatz_Bezeichnung
                                    '--------------------------------------------------
                                    Datensatz_Opak_Nr = 0
                                    Datensatz_Transparent_Nr = 0
                                    Datensatz_Dach_Nr = 0
                                    Datensatz_Zone_Nr = 0
                                    Datensatz_Heizungsanlage_Nr = 0
                                    Datensatz_Trinkwarmwasseranlage_Nr = 0
                                    Datensatz_Lueftung_Nr = 0
                                    Datensatz_Klima_Nr = 0
                                    Datensatz_Energietraeger_Nr = 0
                                    Datensatz_erneuerbare_Energien_65EE_Regel_Nr = 0
                                    Datensatz_erneuerbare_Energien_65EE_keine_Regel_Nr = 0
                                    Datensatz_Energietraeger_Daten_Nr = 0
                                    Datensatz_Zeitraum_Strom_Daten_Nr = 0
                                    Datensatz_Nutzung_Gebaeudekategorie_Nr = 0
                                    Datensatz_Modernisierungsempfehlungen_Nr += 1
                                    '--------------------------------------------------
                                Case "Leerstandszuschlag-Heizung", "Leerstandszuschlag-Warmwasser", "Leerstandszuschlag-thermisch-erzeugte-Kaelte", "Leerstandszuschlag-Strom", "Warmwasserzuschlag", "Kuehlzuschlag", "Strom-Daten"
                                    '--------------------------------------------------
                                    Datensatz_Ebene_4 = Datensatz_Bezeichnung
                                    '--------------------------------------------------
                                    Datensatz_Opak_Nr = 0
                                    Datensatz_Transparent_Nr = 0
                                    Datensatz_Dach_Nr = 0
                                    Datensatz_Zone_Nr = 0
                                    Datensatz_Heizungsanlage_Nr = 0
                                    Datensatz_Trinkwarmwasseranlage_Nr = 0
                                    Datensatz_Lueftung_Nr = 0
                                    Datensatz_Klima_Nr = 0
                                    Datensatz_Energietraeger_Nr = 0
                                    Datensatz_erneuerbare_Energien_65EE_Regel_Nr = 0
                                    Datensatz_erneuerbare_Energien_65EE_keine_Regel_Nr = 0
                                    Datensatz_Energietraeger_Daten_Nr = 0
                                    Datensatz_Zeitraum_Strom_Daten_Nr = 0
                                    Datensatz_Nutzung_Gebaeudekategorie_Nr = 0
                                    Datensatz_Modernisierungsempfehlungen_Nr = 0
                                    '--------------------------------------------------
                                Case "Zeitraum", "Verbrauchsperiode"
                                    '--------------------------------------------------
                                    Datensatz_Ebene_5 = Datensatz_Bezeichnung
                                    '--------------------------------------------------
                                    Datensatz_Zeitraum_Daten_Nr += 1
                                    '--------------------------------------------------
                                Case Else
                                    If Datensatz_Bezeichnung <> "" Then
                                        Datensatz_Typ = Datensatz_Bezeichnung
                                    End If
                            End Select
                            '-----------------------------------------------------
                        Case Xml.XmlNodeType.EndElement
                            Select Case Datensatz_Bezeichnung
                                Case "Gebaeudebezogene-Daten", "Energieausweis-Daten"
                                    Datensatz_Ebene_1 = ""
                                Case "Wohngebaeude", "Nichtwohngebaeude", "NWG-Diagramm-Daten"
                                    Datensatz_Ebene_2 = ""
                                Case "Bedarfswerte-4108-4701", "Bedarfsdaten-4108-4701", "Bedarfswerte-easy", "Bedarfswerte-18599", "Verbrauchswerte", "Verbrauchsdaten", "Bedarfswerte-NWG", "Verbrauchswerte-NWG", "Zusaetzliche-Verbrauchsdaten"
                                    Datensatz_Ebene_3 = ""
                                Case "Bauteil-Opak", "Bauteil-Transparent", "Bauteil-Dach", "Zone", "Heizungsanlage", "Heizsystem", "Trinkwarmwasseranlage", "Warmwasserbereitungssystem", "Kaelteanlage", "RLT-System", "Energietraeger-Liste", "Angaben-erneuerbare-Energien-65EE-Regel", "Angaben-erneuerbare-Energien-keine-65EE-Regel", "Energietraeger", "Zeitraum-Strom", "Nutzung-Gebaeudekategorie", "Modernisierungsempfehlungen", "Leerstandszuschlag-Heizung", "Leerstandszuschlag-Warmwasser", "Leerstandszuschlag-thermisch-erzeugte-Kaelte", "Leerstandszuschlag-Strom", "Warmwasserzuschlag", "Kuehlzuschlag", "Strom-Daten"
                                    Datensatz_Ebene_4 = ""
                                Case "Zeitraum", "Verbrauchsperiode"
                                    Datensatz_Ebene_5 = ""
                            End Select
                            '-----------------------------------------------------
                        Case Xml.XmlNodeType.Text
                            If .Value <> "" Then
                                Datensatz_Wert = .Value
                                Datensatz_Lesen(Datensatz_Ebene_1, Datensatz_Ebene_2, Datensatz_Ebene_3, Datensatz_Ebene_4, Datensatz_Ebene_5, Datensatz_Opak_Nr, Datensatz_Transparent_Nr, Datensatz_Dach_Nr, Datensatz_Zone_Nr, Datensatz_Heizungsanlage_Nr, Datensatz_Trinkwarmwasseranlage_Nr, Datensatz_Lueftung_Nr, Datensatz_Klima_Nr, Datensatz_Energietraeger_Nr, Datensatz_erneuerbare_Energien_65EE_Regel_Nr, Datensatz_erneuerbare_Energien_65EE_keine_Regel_Nr, Datensatz_Energietraeger_Daten_Nr, Datensatz_Zeitraum_Daten_Nr, Datensatz_Zeitraum_Strom_Daten_Nr, Datensatz_Nutzung_Gebaeudekategorie_Nr, Datensatz_Modernisierungsempfehlungen_Nr, Datensatz_Typ, Datensatz_Wert)
                            End If
                            'Case Xml.XmlNodeType.Comment

                    End Select
                    '-----------------------------------------------------
                Loop
                '-----------------------------------------------------
                .Close()
                '-----------------------------------------------------
            End With
            '-----------------------------------------------------
            Dim test As VarXML = Variable_XML_Import
            '-----------------------------------------------------
        Catch ex As Exception
            Fehlerfenster(ex)
            Return
        End Try

    End Sub
    '--------------------------------------------------
    ''' <summary>
    ''' In dieser Anweisung wird die Importdatei ausgelesen. Werte werden zugewiesen.
    ''' </summary>
    Sub Datensatz_Lesen(ByVal Ebene_1 As String, ByVal Ebene_2 As String, ByVal Ebene_3 As String, ByVal Ebene_4 As String, ByVal Ebene_5 As String, ByVal Opak_Nr As Integer, ByVal Transparent_Nr As Integer, ByVal Dach_Nr As Integer, ByVal Zone_Nr As Integer, ByVal Heizungsanlage_Nr As Integer, ByVal Trinkwarmwasseranlage_Nr As Integer, ByVal Lueftung_Nr As Integer, ByVal Klima_Nr As Integer, ByVal Energietraeger_Nr As Integer, ByVal EE_65EE_Regel_Nr As Integer, ByVal EE_65EE_keine_Regel_Nr As Integer, ByVal Energietraeger_Daten_Nr As Integer, ByVal Zeitraum_Daten_Nr As Integer, ByVal Zeitraum_Strom_Daten_Nr As Integer, ByVal Nutzung_Gebaeudekategorie_Nr As Integer, ByVal Modernisierungsempfehlungen_Nr As Integer, ByVal Bezeichnung As String, ByVal Wert As String)
        '--------------------------------------------------
        'Ebene_1
        'Gebaeudebezogene-Daten
        'Energieausweis-Daten
        '--------------------------------------------------
        'Ebene_2
        'Wohngebaeude
        'Nichtwohngebaeude
        '--------------------------------------------------
        'Ebene_3
        'Bedarfswerte-4108-4701
        'Bedarfswerte-easy
        'Bedarfswerte-18599
        'Verbrauchswerte
        'Bedarfswerte-NWG
        'Verbrauchswerte-NWG
        'Zusaetzliche-Verbrauchsdaten
        '--------------------------------------------------
        'Ebene_4
        'Bauteil-Opak
        'Bauteil-Transparent
        'Bauteil-Dach
        'Zone
        'Heizungsanlage
        'Trinkwarmwasseranlage
        'Klima
        'Lüftung
        'Energietraeger-Liste
        'Angaben-erneuerbare-Energien-65EE-Regel
        'Angaben-erneuerbare-Energien-65EE-keine-Regel
        'Energietraeger
        'Zeitraum-Strom
        'Nutzung-Gebaeudekategorie
        'Modernisierungsempfehlungen
        'Leerstandszuschlag-Heizung
        'Leerstandszuschlag-Warmwasser
        'Leerstandszuschlag-thermisch-erzeugte-Kaelte
        'Leerstandszuschlag-Strom
        'Warmwasserzuschlag
        'Kuehlzuschlag
        'Strom-Daten
        '--------------------------------------------------
        'Ebene_5
        'Zeitraum
        '--------------------------------------------------
        Wert = LTrim(RTrim(Wert))
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            With Variable_XML_Import
                '--------------------------------------------------
                Select Case Ebene_1
                    Case "Gebaeudebezogene-Daten"
                        '--------------------------------------------------
                        With .Gebaeudebezogene_Daten
                            '-----------------------------------------------------
                            Select Case Bezeichnung
                                Case "Projektbezeichnung-Aussteller"
                                    .Projektbezeichnung_Aussteller.Wert = Wert
                                Case "Gebaeudeadresse-Strasse-Nr"
                                    .Gebaeudeadresse_Strasse_Nr.Wert = Wert
                                Case "Gebaeudeadresse-Postleitzahl"
                                    .Gebaeudeadresse_Postleitzahl.Wert = Wert
                                Case "Gebaeudeadresse-Ort"
                                    .Gebaeudeadresse_Ort.Wert = Wert
                                Case "Ausstellervorname"
                                    .Ausstellervorname.Wert = Wert
                                Case "Ausstellername"
                                    .Ausstellername.Wert = Wert
                                Case "Aussteller-Bezeichnung"
                                    .Aussteller_Bezeichnung.Wert = Wert
                                Case "Aussteller-Strasse-Nr"
                                    .Aussteller_Strasse_Nr.Wert = Wert
                                Case "Aussteller-PLZ"
                                    .Aussteller_PLZ.Wert = Wert
                                Case "Aussteller-Ort"
                                    .Aussteller_Ort.Wert = Wert
                                Case "Zusatzinfos-beigefuegt"
                                    .Zusatzinfos_beigefuegt.Wert = ImportBool(Wert)
                                Case "Angaben-erhaeltlich"
                                    .Angaben_erhaeltlich.Wert = Wert
                                Case "Ergaenzdende-Erlaeuterungen"
                                    .Ergaenzdende_Erlaeuterungen.Wert = Wert
                                Case "Ergaenzdende-Erlaeuterungen-EE"
                                    .Ergaenzdende_Erlaeuterungen_EE.Wert = Wert
                                Case "CO2-Emissionen"
                                    .Treibhausgasemissionen.Wert = ImportDec(Wert)
                            End Select
                            '-----------------------------------------------------
                        End With
                        '--------------------------------------------------
                    Case "Energieausweis-Daten"
                        '--------------------------------------------------
                        With .Energieausweis_Daten
                            '-----------------------------------------------------
                            Select Case Bezeichnung
                                Case "Registriernummer"
                                    .Registriernummer.Wert = Wert
                                Case "Ausstellungsdatum"
                                    .Ausstellungsdatum.Wert = Wert
                                Case "Bundesland"
                                    .Bundesland.Wert = Wert
                                Case "Postleitzahl"
                                    .Postleitzahl.Wert = Wert
                                Case "Gebaeudeteil"
                                    .Gebaeudeteil.Wert = Wert
                                Case "Baujahr-Gebaeude"
                                    .Baujahr_Gebaeude.Wert = Wert
                                Case "Altersklasse-Gebaeude"
                                    .Altersklasse_Gebaeude.Wert = Wert
                                Case "Baujahr-Waermeerzeuger"
                                    .Baujahr_Waermeerzeuger.Wert = Wert
                                Case "Altersklasse-Waermeerzeuger"
                                    .Altersklasse_Waermeerzeuger.Wert = Wert
                                Case "wesentliche-Energietraeger-Heizung", "wesentliche-Energietraeger"
                                    .wesentliche_Energietraeger_Heizung.Wert = Wert
                                Case "wesentliche-Energietraeger-Warmwasser"
                                    .wesentliche_Energietraeger_Warmwasser.Wert = Wert
                                Case "Erneuerbare-Art"
                                    .Erneuerbare_Art.Wert = Wert
                                Case "Erneuerbare-Verwendung"
                                    .Erneuerbare_Verwendung.Wert = Wert
                                Case "Lueftungsart-Fensterlueftung"
                                    .Lueftungsart_Fensterlueftung.Wert = ImportBool(Wert)
                                Case "Lueftungsart-Schachtlueftung"
                                    .Lueftungsart_Schachtlueftung.Wert = ImportBool(Wert)
                                Case "Lueftungsart-Anlage-o-WRG"
                                    .Lueftungsart_Anlage_o_WRG.Wert = ImportBool(Wert)
                                Case "Lueftungsart-Anlage-m-WRG"
                                    .Lueftungsart_Anlage_m_WRG.Wert = ImportBool(Wert)
                                Case "Anlage-zur-Kuehlung"
                                    .Anlage_zur_Kuehlung.Wert = ImportBool(Wert)
                                Case "Kuehlungsart-passive-Kuehlung"
                                    .Kuehlungsart_passive_Kuehlung.Wert = ImportBool(Wert)
                                Case "Kuehlungsart-Strom"
                                    .Kuehlungsart_Strom.Wert = ImportBool(Wert)
                                Case "Kuehlungsart-Waerme"
                                    .Kuehlungsart_Waerme.Wert = ImportBool(Wert)
                                Case "Kuehlungsart-gelieferte-Kaelte"
                                    .Kuehlungsart_gelieferte_Kaelte.Wert = ImportBool(Wert)
                                Case "Keine-inspektionspflichtige-Anlage"
                                    .Keine_inspektionspflichtige_Anlage.Wert = ImportBool(Wert)
                                Case "Anzahl-Klimanlagen"
                                    .Anzahl_Klimanlagen.Wert = ImportInt(Wert)
                                Case "Anlage-groesser-12kW-ohneGA"
                                    .Anlage_groesser_12kW_ohneGA.Wert = ImportBool(Wert)
                                Case "Anlage-groesser-12kW-mitGA"
                                    .Anlage_groesser_12kW_mitGA.Wert = ImportBool(Wert)
                                Case "Anlage-groesser-70kW"
                                    .Anlage_groesser_70kW.Wert = ImportBool(Wert)
                                Case "Faelligkeitsdatum-Inspektion"
                                    .Faelligkeitsdatum_Inspektion.Wert = CDate(ImportDate(Wert))
                                Case "Nutzung-zur-Erfuellung-von-EE-neue-Anlage"
                                    .Nutzung_zur_Erfuellung_von_EE_neue_Anlage.Wert = ImportBool(Wert)
                                Case "EE-Angabe-Warmwasser"
                                    .EE_Angabe_Warmwasser.Wert = ImportBool(Wert)
                                Case "EE-Angabe-Heizung"
                                    .EE_Angabe_Heizung.Wert = ImportBool(Wert)
                                Case "Keine-Pauschale-Erfuellungsoptionen-Anlagentyp"
                                    .Keine_Pauschale_Erfuellungsoptionen_Anlagentyp.Wert = ImportBool(Wert)
                                Case "Hausuebergabestation"
                                    .Hausuebergabestation.Wert = ImportBool(Wert)
                                Case "Waermepumpe"
                                    .Waermepumpe.Wert = ImportBool(Wert)
                                Case "Stromdirektheizung"
                                    .Stromdirektheizung.Wert = ImportBool(Wert)
                                Case "Solarthermische-Anlage"
                                    .Solarthermische_Anlage.Wert = ImportBool(Wert)
                                Case "Heizungsanlage-Biomasse-Wasserstoff-Wasserstoffderivale"
                                    .Heizungsanlage_Biomasse_Wasserstoff_Wasserstoffderivale.Wert = ImportBool(Wert)
                                Case "Waermepumpen-Hybridheizung"
                                    .Waermepumpen_Hybridheizung.Wert = ImportBool(Wert)
                                Case "Solarthermie-Hybridheizung"
                                    .Solarthermie_Hybridheizung.Wert = ImportBool(Wert)
                                Case "Dezentral-elektrische-Warmwasserbereitung"
                                    .Dezentral_elektrische_Warmwasserbereitung.Wert = ImportBool(Wert)
                                Case "Nutzung-bei-Bestandsanlagen"
                                    .Nutzung_bei_Bestandsanlagen.Wert = ImportBool(Wert)
                                Case "Treibhausgasemissionen"
                                    .Treibhausgasemissionen.Wert = ImportDec(Wert)
                                Case "Ausstellungsanlass"
                                    .Ausstellungsanlass.Wert = Wert
                                Case "Datenerhebung-Aussteller"
                                    .Datenerhebung_Aussteller.Wert = ImportBool(Wert)
                                Case "Datenerhebung-Eigentuemer"
                                    .Datenerhebung_Eigentuemer.Wert = ImportBool(Wert)
                                Case "Empfehlungen-moeglich"
                                    .Empfehlungen_moeglich.Wert = ImportBool(Wert)
                                Case "Keine-Modernisierung-Erweiterung-Vorhaben"
                                    .Keine_Modernisierung_Erweiterung_Vorhaben.Wert = ImportBool(Wert)
                                Case "Softwarehersteller-Programm-Version"
                                    .Softwarehersteller_Programm_Version.Wert = Wert
                            End Select
                            '-----------------------------------------------------
                        End With
                        '--------------------------------------------------
                End Select
                '--------------------------------------------------
                Select Case Ebene_2
                    Case "Wohngebaeude"
                        '--------------------------------------------------
                        With .Energieausweis_Daten.Wohngebaeude
                            '-----------------------------------------------------
                            Select Case Bezeichnung
                                Case "Gebaeudetyp"
                                    .Gebaeudetyp.Wert = Wert
                                Case "Anzahl-Wohneinheiten"
                                    .Anzahl_Wohneinheiten.Wert = ImportInt(Wert)
                                Case "Gebaeudenutzflaeche"
                                    .Gebaeudenutzflaeche.Wert = ImportInt(Wert)
                            End Select
                            '-----------------------------------------------------
                        End With
                        '--------------------------------------------------
                    Case "Nichtwohngebaeude"
                        '--------------------------------------------------
                        With .Energieausweis_Daten.Nichtwohngebaeude
                            '-----------------------------------------------------
                            Select Case Bezeichnung
                                Case "Hauptnutzung-Gebaeudekategorie"
                                    .Hauptnutzung_Gebaeudekategorie.Wert = ImportHauptnutzung(Wert)
                                Case "Hauptnutzung-Gebaeudekategorie-Sonstiges-Beschreibung"
                                    .Hauptnutzung_Gebaeudekategorie_Sonstiges_Beschreibung.Wert = Wert
                                Case "Nettogrundflaeche"
                                    .Nettogrundflaeche.Wert = ImportInt(Wert)
                            End Select
                            '-----------------------------------------------------
                        End With
                        '--------------------------------------------------
                    Case "NWG-Diagramm-Daten"
                        '--------------------------------------------------
                        With .Gebaeudebezogene_Daten.NWG_Aushang_Daten
                            '-----------------------------------------------------
                            Select Case Bezeichnung
                                Case "Nutzenergiebedarf-Heizung-Diagramm"
                                    .Nutzenergiebedarf_Heizung_Diagramm.Wert = ImportDec(Wert)
                                Case "Nutzenergiebedarf-Trinkwarmwasser-Diagramm"
                                    .Nutzenergiebedarf_Trinkwarmwasser_Diagramm.Wert = ImportDec(Wert)
                                Case "Nutzenergiebedarf-Beleuchtung-Diagramm"
                                    .Nutzenergiebedarf_Beleuchtung_Diagramm.Wert = ImportDec(Wert)
                                Case "Nutzenergiebedarf-Lueftung-Diagramm"
                                    .Nutzenergiebedarf_Lueftung_Diagramm.Wert = ImportDec(Wert)
                                Case "Nutzenergiebedarf-Kuehlung-Befeuchtung-Diagramm"
                                    .Nutzenergiebedarf_Kuehlung_Befeuchtung_Diagramm.Wert = ImportDec(Wert)
                                Case "Endenergiebedarf-Heizung-Diagramm"
                                    .Endenergiebedarf_Heizung_Diagramm.Wert = ImportDec(Wert)
                                Case "Endenergiebedarf-Trinkwarmwasser-Diagramm"
                                    .Endenergiebedarf_Trinkwarmwasser_Diagramm.Wert = ImportDec(Wert)
                                Case "Endenergiebedarf-Beleuchtung-Diagramm"
                                    .Endenergiebedarf_Beleuchtung_Diagramm.Wert = ImportDec(Wert)
                                Case "Endenergiebedarf-Lueftung-Diagramm"
                                    .Endenergiebedarf_Lueftung_Diagramm.Wert = ImportDec(Wert)
                                Case "Endenergiebedarf-Kuehlung-Befeuchtung-Diagramm"
                                    .Endenergiebedarf_Kuehlung_Befeuchtung_Diagramm.Wert = ImportDec(Wert)
                                Case "Primaerenergiebedarf-Heizung-Diagramm"
                                    .Primaerenergiebedarf_Heizung_Diagramm.Wert = ImportDec(Wert)
                                Case "Primaerenergiebedarf-Trinkwarmwasser-Diagramm"
                                    .Primaerenergiebedarf_Trinkwarmwasser_Diagramm.Wert = ImportDec(Wert)
                                Case "Primaerenergiebedarf-Beleuchtung-Diagramm"
                                    .Primaerenergiebedarf_Beleuchtung_Diagramm.Wert = ImportDec(Wert)
                                Case "Primaerenergiebedarf-Lueftung-Diagramm"
                                    .Primaerenergiebedarf_Lueftung_Diagramm.Wert = ImportDec(Wert)
                                Case "Primaerenergiebedarf-Kuehlung-Befeuchtung-Diagramm"
                                    .Primaerenergiebedarf_Kuehlung_Befeuchtung_Diagramm.Wert = ImportDec(Wert)
                            End Select
                            '-----------------------------------------------------
                        End With
                        '--------------------------------------------------
                End Select
                '--------------------------------------------------
                Select Case Ebene_3
                    Case "Bedarfswerte-18599"
                        '--------------------------------------------------
                        .Berechnungsverfahren = "Bedarfsberechnung-18599"
                        '--------------------------------------------------
                        With .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_18599
                            '-----------------------------------------------------
                            Select Case Bezeichnung
                                Case "Wohngebaeude-Anbaugrad"
                                    .Wohngebaeude_Anbaugrad.Wert = Wert
                                Case "Bruttovolumen"
                                    .Bruttovolumen.Wert = ImportInt(Wert)
                                Case "durchschnittliche-Geschosshoehe"
                                    .durchschnittliche_Geschosshoehe.Wert = ImportDec(Wert)
                                Case "Waermebrueckenzuschlag"
                                    .Waermebrueckenzuschlag.Wert = ImportDec(Wert)
                                Case "Transmissionswaermesenken"
                                    .Transmissionswaermesenken.Wert = ImportInt(Wert)
                                Case "Luftdichtheit"
                                    .Luftdichtheit.Wert = Wert
                                Case "Lueftungswaermesenken"
                                    .Lueftungswaermesenken.Wert = ImportInt(Wert)
                                Case "Waermequellen-durch-solare-Einstrahlung"
                                    .Waermequellen_durch_solare_Einstrahlung.Wert = ImportInt(Wert)
                                Case "Interne-Waermequellen"
                                    .Interne_Waermequellen.Wert = ImportInt(Wert)
                                Case "Pufferspeicher-Nenninhalt"
                                    .Pufferspeicher_Nenninhalt.Wert = ImportInt(Wert)
                                Case "Auslegungstemperatur"
                                    .Auslegungstemperatur.Wert = Wert
                                Case "Heizsystem-innerhalb-Huelle"
                                    .Heizsystem_innerhalb_Huelle.Wert = ImportBool(Wert)
                                Case "Trinkwarmwasserspeicher-Nenninhalt"
                                    .Trinkwarmwasserspeicher_Nenninhalt.Wert = ImportInt(Wert)
                                Case "Trinkwarmwasserverteilung-Zirkulation"
                                    .Trinkwarmwasserverteilung_Zirkulation.Wert = ImportBool(Wert)
                                Case "Vereinfachte-Datenaufnahme"
                                    .Vereinfachte_Datenaufnahme.Wert = ImportBool(Wert)
                                Case "spezifischer-Transmissionswaermetransferkoeffizient-Ist"
                                    .spezifischer_Transmissionswaermetransferkoeffizient_Ist.Wert = ImportDec(Wert)
                                Case "spezifischer-Transmissionswaermetransferkoeffizient-Hoechstwert"
                                    .spezifischer_Transmissionswaermetransferkoeffizient_Hoechstwert.Wert = ImportDec(Wert)
                                Case "angerechneter-lokaler-erneuerbarer-Strom"
                                    .angerechneter_lokaler_erneuerbarer_Strom.Wert = ImportInt(Wert)
                                Case "Innovationsklausel"
                                    .Innovationsklausel.Wert = ImportBool(Wert)
                                Case "Quartiersregelung"
                                    .Quartiersregelung.Wert = ImportBool(Wert)
                                    '---------------------------------------------------------
                                Case "Primaerenergiebedarf-Hoechstwert-Bestand"
                                    .Primaerenergiebedarf_Hoechstwert_Bestand.Wert = ImportDec(Wert)
                                Case "Endenergiebedarf-Hoechstwert-Bestand"
                                    .Endenergiebedarf_Hoechstwert_Bestand.Wert = ImportDec(Wert)
                                Case "Treibhausgasemissionen-Hoechstwert-Bestand"
                                    .Treibhausgasemissionen_Hoechstwert_Bestand.Wert = ImportDec(Wert)
                                    '---------------------------------------------------------
                                Case "Primaerenergiebedarf-Hoechstwert-Neubau"
                                    .Primaerenergiebedarf_Hoechstwert_Neubau.Wert = ImportDec(Wert)
                                Case "Endenergiebedarf-Hoechstwert-Neubau"
                                    .Endenergiebedarf_Hoechstwert_Neubau.Wert = ImportDec(Wert)
                                Case "Treibhausgasemissionen-Hoechstwert-Neubau"
                                    .Treibhausgasemissionen_Hoechstwert_Neubau.Wert = ImportDec(Wert)
                                    '---------------------------------------------------------
                                Case "Endenergiebedarf-Waerme-AN"
                                    .Endenergiebedarf_Waerme_AN.Wert = ImportDec(Wert)
                                Case "Endenergiebedarf-Hilfsenergie-AN"
                                    .Endenergiebedarf_Hilfsenergie_AN.Wert = ImportDec(Wert)
                                Case "Endenergiebedarf-Gesamt"
                                    .Endenergiebedarf_Gesamt.Wert = ImportDec(Wert)
                                Case "Primaerenergiebedarf-AN"
                                    .Primaerenergiebedarf_AN.Wert = ImportDec(Wert)
                                Case "Energieeffizienzklasse"
                                    .Energieeffizienzklasse.Wert = Wert
                            '---------------------------------------------------------
                            'EnEV 2016/GEG 2020/2023
                                Case "Art-der-Nutzung-erneuerbaren-Energie-1", "EEWaermeG-Nutzung-1"
                                    .Art_der_Nutzung_erneuerbaren_Energie_1.Wert = Wert
                                Case "Deckungsanteil-1"
                                    .Deckungsanteil_1.Wert = ImportInt(Wert)
                                Case "Anteil-der-Pflichterfuellung-1"
                                    .Anteil_der_Pflichterfuellung_1.Wert = ImportInt(Wert)
                                Case "Art-der-Nutzung-erneuerbaren-Energie-2", "EEWaermeG-Nutzung-2"
                                    .Art_der_Nutzung_erneuerbaren_Energie_2.Wert = Wert
                                Case "Deckungsanteil-2"
                                    .Deckungsanteil_2.Wert = ImportInt(Wert)
                                Case "Anteil-der-Pflichterfuellung-2"
                                    .Anteil_der_Pflichterfuellung_2.Wert = ImportInt(Wert)
                            '---------------------------------------------------------
                            'EnEV 2016
                                Case "EEWaermeG-Nutzung-3"
                                    .Art_der_Nutzung_erneuerbaren_Energie_3.Wert = Wert
                                Case "Deckungsanteil-3"
                                    .Deckungsanteil_3.Wert = ImportInt(Wert)
                                Case "verschaerft-nach-EEWaermeG-7-1-2-eingehalten"
                                    .verschaerft_nach_EEWaermeG_7_1_2_eingehalten.Wert = ImportBool(Wert)
                                Case "verschaerft-nach-EEWaermeG-8"
                                    .verschaerft_nach_EEWaermeG_8.Wert = ImportInt(Wert)
                                Case "Primaerenergiebedarf-Hoechstwert-verschaerft"
                                    .Primaerenergiebedarf_Hoechstwert_verschaerft.Wert = ImportDec(Wert)
                                Case "spezifischer-Transmissionswaermetransferkoeffizient-verschaerft"
                                    .spezifischer_Transmissionswaermetransferkoeffizient_verschaerft.Wert = ImportDec(Wert)
                                    '---------------------------------------------------------
                            'GEG 2020/2023
                                Case "verschaerft-nach-GEG-45-eingehalten"
                                    .verschaerft_nach_GEG_45_eingehalten.Wert = ImportBool(Wert)
                                Case "nicht-verschaerft-nach-GEG-34"
                                    .nicht_verschaerft_nach_GEG_34.Wert = ImportBool(Wert)
                                Case "verschaerft-nach-GEG-34"
                                    .verschaerft_nach_GEG_34.Wert = ImportInt(Wert)
                                Case "Anforderung-nach-GEG-16-unterschritten"
                                    .Anforderung_nach_GEG_16_unterschritten.Wert = ImportInt(Wert)
                                Case "spezifischer-Transmissionswaermetransferkoeffizient-verschaerft"
                                    .spezifischer_Transmissionswaermetransferkoeffizient_verschaerft.Wert = ImportDec(Wert)
                            '---------------------------------------------------------
                            'GEG 2024
                                Case "Anteil-an-Waermeenergiebedarf-Berechnung"
                                    .Anteil_an_Waermeenergiebedarf_Berechnung.Wert = ImportBool(Wert)
                                Case "Weitere-Eintraege-und-Erlaeuterungen-in-der-Anlage"
                                    .Weitere_Eintraege_und_Erlaeuterungen_in_der_Anlage.Wert = ImportBool(Wert)
                            '---------------------------------------------------------
                            'Sommerlicher Wärmeschutz
                                Case "Sommerlicher-Waermeschutz"
                                    .Sommerlicher_Waermeschutz.Wert = Wert
                            '---------------------------------------------------------
                            'Zusätzliche Verbrauchsdaten
                                Case "Treibhausgasemissionen-Zusaetzliche-Verbrauchsdaten"
                                    Variable_XML_Import.Energieausweis_Daten.Treibhausgasemissionen_Zusaetzliche_Verbrauchsdaten.Wert = ImportDec(Wert)
                                    '---------------------------------------------------------
                            End Select
                        End With
                        '--------------------------------------------------
                    Case "Bedarfswerte-NWG"
                        '--------------------------------------------------
                        .Berechnungsverfahren = "Bedarfsberechnung-18599"
                        '--------------------------------------------------
                        With .Energieausweis_Daten.Nichtwohngebaeude.Bedarfswerte_NWG
                            '-----------------------------------------------------
                            Select Case Bezeichnung
                                Case "Bruttovolumen"
                                    .Bruttovolumen.Wert = ImportInt(Wert)
                                Case "Waermebrueckenzuschlag"
                                    .Waermebrueckenzuschlag.Wert = ImportDec(Wert)
                                Case "mittlere-Waermedurchgangskoeffizienten"
                                    .mittlere_Waermedurchgangskoeffizienten.Wert = ImportBool(Wert)
                                Case "Transmissionswaermesenken"
                                    .Transmissionswaermesenken.Wert = ImportInt(Wert)
                                Case "Luftdichtheit"
                                    .Luftdichtheit.Wert = Wert
                                Case "Pufferspeicher-Nenninhalt"
                                    .Pufferspeicher_Nenninhalt.Wert = ImportInt(Wert)
                                Case "Auslegungstemperatur"
                                    .Auslegungstemperatur.Wert = Wert
                                Case "Heizsystem-innerhalb-Huelle"
                                    .Heizsystem_innerhalb_Huelle.Wert = ImportBool(Wert)
                                Case "Trinkwarmwasserspeicher-Nenninhalt"
                                    .Trinkwarmwasserspeicher_Nenninhalt.Wert = ImportInt(Wert)
                                Case "Trinkwarmwasserverteilung-Zirkulation"
                                    .Trinkwarmwasserverteilung_Zirkulation.Wert = ImportBool(Wert)
                                Case "Deckungsanteil-RLT-Kuehlung"
                                    .Deckungsanteil_RLT_Kuehlung.Wert = ImportInt(Wert)
                                Case "Deckungsanteil-Direkte-Raumkuehlung"
                                    .Deckungsanteil_Direkte_Raumkuehlung.Wert = ImportInt(Wert)
                                Case "Automatisierungsgrad"
                                    .Automatisierungsgrad.Wert = Wert
                                Case "Automatisierungsgrad-Technisches-Gebaeudemanagement"
                                    .Automatisierungsgrad_Technisches_Gebaeudemanagement.Wert = Wert
                                Case "angerechneter-lokaler-erneuerbarer-Strom"
                                    .angerechneter_lokaler_erneuerbarer_Strom.Wert = ImportInt(Wert)
                                Case "Innovationsklausel"
                                    .Innovationsklausel.Wert = ImportBool(Wert)
                                Case "Quartiersregelung"
                                    .Quartiersregelung.Wert = ImportBool(Wert)
                                    '---------------------------------------------------------
                                Case "Primaerenergiebedarf-Hoechstwert-Bestand"
                                    .Primaerenergiebedarf_Hoechstwert_Bestand.Wert = ImportDec(Wert)
                                Case "Endenergiebedarf-Hoechstwert-Bestand"
                                    .Endenergiebedarf_Hoechstwert_Bestand.Wert = ImportDec(Wert)
                                Case "Treibhausgasemissionen-Hoechstwert-Bestand"
                                    .Treibhausgasemissionen_Hoechstwert_Bestand.Wert = ImportDec(Wert)
                                    '---------------------------------------------------------
                                Case "Primaerenergiebedarf-Hoechstwert-Neubau"
                                    .Primaerenergiebedarf_Hoechstwert_Neubau.Wert = ImportDec(Wert)
                                Case "Endenergiebedarf-Hoechstwert-Neubau"
                                    .Endenergiebedarf_Hoechstwert_Neubau.Wert = ImportDec(Wert)
                                Case "Treibhausgasemissionen-Hoechstwert-Neubau"
                                    .Treibhausgasemissionen_Hoechstwert_Neubau.Wert = ImportDec(Wert)
                                    '---------------------------------------------------------
                                Case "Endenergiebedarf-Waerme-NGF"
                                    .Endenergiebedarf_Waerme_NGF.Wert = ImportDec(Wert)
                                Case "Endenergiebedarf-Strom-NGF"
                                    .Endenergiebedarf_Strom_NGF.Wert = ImportDec(Wert)
                                Case "Endenergiebedarf-Gesamt-NGF"
                                    .Endenergiebedarf_Gesamt_NGF.Wert = ImportDec(Wert)
                                Case "Primaerenergiebedarf-NGF"
                                    .Primaerenergiebedarf_NGF.Wert = ImportDec(Wert)
                                Case "Ein-Zonen-Modell"
                                    .Ein_Zonen_Modell.Wert = ImportBool(Wert)
                                Case "Vereinfachte-Datenaufnahme"
                                    .Vereinfachte_Datenaufnahme.Wert = ImportBool(Wert)
                                Case "Vereinfachungen-18599-1-D"
                                    .Vereinfachungen_18599_1_D.Wert = ImportBool(Wert)
                            '---------------------------------------------------------
                             'EnEV 2016/GEG 2020/2023
                                Case "Art-der-Nutzung-erneuerbaren-Energie-1", "EEWaermeG-Nutzung-1"
                                    .Art_der_Nutzung_erneuerbaren_Energie_1.Wert = Wert
                                Case "Deckungsanteil-1"
                                    .Deckungsanteil_1.Wert = ImportInt(Wert)
                                Case "Anteil-der-Pflichterfuellung-1"
                                    .Anteil_der_Pflichterfuellung_1.Wert = ImportInt(Wert)
                                Case "Art-der-Nutzung-erneuerbaren-Energie-2", "EEWaermeG-Nutzung-2"
                                    .Art_der_Nutzung_erneuerbaren_Energie_2.Wert = Wert
                                Case "Deckungsanteil-2"
                                    .Deckungsanteil_2.Wert = ImportInt(Wert)
                                Case "Anteil-der-Pflichterfuellung-2"
                                    .Anteil_der_Pflichterfuellung_2.Wert = ImportInt(Wert)
                            '---------------------------------------------------------
                            'EnEV 2016
                                Case "EEWaermeG-Nutzung-3"
                                    .Art_der_Nutzung_erneuerbaren_Energie_3.Wert = Wert
                                Case "Deckungsanteil-3"
                                    .Deckungsanteil_3.Wert = ImportInt(Wert)
                                Case "verschaerft-nach-EEWaermeG-7-1-2-eingehalten"
                                    .verschaerft_nach_EEWaermeG_7_1_2_eingehalten.Wert = ImportBool(Wert)
                                Case "verschaerft-nach-EEWaermeG-8"
                                    .verschaerft_nach_EEWaermeG_8.Wert = ImportInt(Wert)
                                Case "Primaerenergiebedarf-Hoechstwert-verschaerft"
                                    .Primaerenergiebedarf_Hoechstwert_verschaerft.Wert = ImportDec(Wert)
                                Case "spezifischer-Transmissionswaermetransferkoeffizient-verschaerft"
                                    .spezifischer_Transmissionswaermetransferkoeffizient_verschaerft.Wert = ImportDec(Wert)
                            '---------------------------------------------------------
                            'GEG 2020/2023
                                Case "verschaerft-nach-GEG-45-eingehalten"
                                    .verschaerft_nach_GEG_45_eingehalten.Wert = ImportBool(Wert)
                                Case "nicht-verschaerft-nach-GEG-34"
                                    .nicht_verschaerft_nach_GEG_34.Wert = ImportBool(Wert)
                                Case "verschaerft-nach-GEG-34"
                                    .verschaerft_nach_GEG_34.Wert = ImportInt(Wert)
                                Case "Anforderung-nach-GEG-19-unterschritten"
                                    .Anforderung_nach_GEG_19_unterschritten.Wert = ImportInt(Wert)
                                Case "Anforderung-nach-GEG-52-Renovierung-eingehalten"
                                    .Anforderung_nach_GEG_52_Renovierung_eingehalten.Wert = ImportBool(Wert)
                            '---------------------------------------------------------
                            'GEG 2024
                                Case "Anteil-an-Waermeenergiebedarf-Berechnung"
                                    .Anteil_an_Waermeenergiebedarf_Berechnung.Wert = ImportBool(Wert)
                                Case "Weitere-Eintraege-und-Erlaeuterungen-in-der-Anlage"
                                    .Weitere_Eintraege_und_Erlaeuterungen_in_der_Anlage.Wert = ImportBool(Wert)
                            '---------------------------------------------------------
                            'Sommerlicher Wärmeschutz
                                Case "Sommerlicher-Waermeschutz"
                                    .Sommerlicher_Waermeschutz.Wert = ImportBool(Wert)
                            '---------------------------------------------------------
                            'Zusätzliche Verbrauchsdaten
                                Case "Treibhausgasemissionen-Zusaetzliche-Verbrauchsdaten"
                                    Variable_XML_Import.Energieausweis_Daten.Treibhausgasemissionen_Zusaetzliche_Verbrauchsdaten.Wert = ImportDec(Wert)
                                    '---------------------------------------------------------
                            End Select
                            '-----------------------------------------------------
                        End With
                        '--------------------------------------------------
                    Case "Bedarfswerte-4108-4701", "Bedarfsdaten-4108-4701"
                        '--------------------------------------------------
                        .Berechnungsverfahren = "Bedarfsberechnung-4108-4701"
                        '--------------------------------------------------
                        Select Case Variable_Steuerung.Gesetzesgrundlage
                            Case "GEG-2024"
                                'Das Berechnungsverfahren nach DIN V 4108/4701 ist für die gewählte Gesetzesgrundlage nicht erlaubt.
                        End Select
                        '--------------------------------------------------
                        With .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_4108_4701
                            '-----------------------------------------------------
                            Select Case Bezeichnung
                                Case "Wohngebaeude-Anbaugrad"
                                    .Wohngebaeude_Anbaugrad.Wert = Wert
                                Case "Bruttovolumen", "Gebaeudevolumen"
                                    .Bruttovolumen.Wert = ImportInt(Wert)
                                Case "durchschnittliche-Geschosshoehe"
                                    .durchschnittliche_Geschosshoehe.Wert = ImportDec(Wert)
                                Case "Waermebrueckenzuschlag"
                                    .Waermebrueckenzuschlag.Wert = ImportDec(Wert)
                                Case "Transmissionswaermeverlust"
                                    .Transmissionswaermeverlust.Wert = ImportInt(Wert)
                                Case "Luftdichtheit"
                                    .Luftdichtheit.Wert = Wert
                                Case "Lueftungswaermeverlust"
                                    .Lueftungswaermeverlust.Wert = ImportInt(Wert)
                                Case "Solare-Waermegewinne", "solare-Gewinne"
                                    .Solare_Waermegewinne.Wert = ImportInt(Wert)
                                Case "Interne-Waermegewinne", "innere-Gewinne"
                                    .Interne_Waermegewinne.Wert = ImportInt(Wert)
                                Case "Pufferspeicher-Nenninhalt", "Pufferspeicher-Volumen"
                                    .Pufferspeicher_Nenninhalt.Wert = ImportInt(Wert)
                                Case "Heizkreisauslegungstemperatur", "Heizverteilung-Temperatur"
                                    .Heizkreisauslegungstemperatur.Wert = Wert
                                Case "Heizungsanlage-innerhalb-Huelle", "Heizanlage-innerhalb-Huelle"
                                    .Heizungsanlage_innerhalb_Huelle.Wert = ImportBool(Wert)
                                Case "Trinkwarmwasserspeicher-Nenninhalt", "Warmwasserspeicher-Volumen"
                                    .Trinkwarmwasserspeicher_Nenninhalt.Wert = ImportInt(Wert)
                                Case "Trinkwarmwasserverteilung-Zirkulation", "Warmwasser-Zirkulation"
                                    .Trinkwarmwasserverteilung_Zirkulation.Wert = ImportBool(Wert)
                                Case "Vereinfachte-Datenaufnahme"
                                    .Vereinfachte_Datenaufnahme.Wert = ImportBool(Wert)
                                Case "spezifischer-Transmissionswaermeverlust-Ist"
                                    .spezifischer_Transmissionswaermeverlust_Ist.Wert = ImportDec(Wert)
                                Case "spezifischer-Transmissionswaermeverlust-Hoechstwert", "spezifischer-Transmissionsverlust-Anforderung"
                                    .spezifischer_Transmissionswaermeverlust_Hoechstwert.Wert = ImportDec(Wert)
                                Case "angerechneter-lokaler-erneuerbarer-Strom"
                                    .angerechneter_lokaler_erneuerbarer_Strom.Wert = ImportInt(Wert)
                                Case "Innovationsklausel"
                                    .Innovationsklausel.Wert = ImportBool(Wert)
                                Case "Quartiersregelung"
                                    .Quartiersregelung.Wert = ImportBool(Wert)
                                    '---------------------------------------------------------
                                Case "Primaerenergiebedarf-Hoechstwert-Neubau"
                                    .Primaerenergiebedarf_Hoechstwert_Neubau.Wert = ImportDec(Wert)
                                Case "Endenergiebedarf-Hoechstwert-Neubau"
                                    .Endenergiebedarf_Hoechstwert_Neubau.Wert = ImportDec(Wert)
                                Case "Treibhausgasemissionen-Hoechstwert-Neubau"
                                    .Treibhausgasemissionen_Hoechstwert_Neubau.Wert = ImportDec(Wert)
                                    '---------------------------------------------------------
                                Case "Primaerenergiebedarf-Hoechstwert-Bestand"
                                    .Primaerenergiebedarf_Hoechstwert_Bestand.Wert = ImportDec(Wert)
                                Case "Endenergiebedarf-Hoechstwert-Bestand"
                                    .Endenergiebedarf_Hoechstwert_Bestand.Wert = ImportDec(Wert)
                                Case "Treibhausgasemissionen-Hoechstwert-Bestand"
                                    .Treibhausgasemissionen_Hoechstwert_Bestand.Wert = ImportDec(Wert)
                                    '---------------------------------------------------------
                                Case "Endenergiebedarf-Waerme-AN", "Endenergiekennwert-Waerme-AN"
                                    .Endenergiebedarf_Waerme_AN.Wert = ImportDec(Wert)
                                Case "Endenergiebedarf-Hilfsenergie-AN", "Endenergiekennwert-Hilfsenergie-AN"
                                    .Endenergiebedarf_Hilfsenergie_AN.Wert = ImportDec(Wert)
                                Case "Endenergiebedarf-Gesamt"
                                    .Endenergiebedarf_Gesamt.Wert = ImportDec(Wert)
                                Case "Primaerenergiebedarf", "Primaerenergiekennwert-Bedarf"
                                    .Primaerenergiebedarf.Wert = ImportDec(Wert)
                                Case "Energieeffizienzklasse"
                                    .Energieeffizienzklasse.Wert = ImportDec(Wert)
                            '---------------------------------------------------------
                            'EnEV 2016 / GEG 2020/2023
                                Case "Art-der-Nutzung-erneuerbaren-Energie-1", "EEWaermeG-Nutzung-1"
                                    .Art_der_Nutzung_erneuerbaren_Energie_1.Wert = Wert
                                Case "Deckungsanteil-1"
                                    .Deckungsanteil_1.Wert = ImportInt(Wert)
                                Case "Anteil-der-Pflichterfuellung-1"
                                    .Anteil_der_Pflichterfuellung_1.Wert = ImportInt(Wert)
                                Case "Art-der-Nutzung-erneuerbaren-Energie-2", "EEWaermeG-Nutzung-2"
                                    .Art_der_Nutzung_erneuerbaren_Energie_2.Wert = Wert
                                Case "Deckungsanteil-2"
                                    .Deckungsanteil_2.Wert = ImportInt(Wert)
                                Case "Anteil-der-Pflichterfuellung-2"
                                    .Anteil_der_Pflichterfuellung_2.Wert = ImportInt(Wert)
                            '---------------------------------------------------------
                            'EnEV 2016
                                Case "EEWaermeG-Nutzung-3"
                                    .Art_der_Nutzung_erneuerbaren_Energie_3.Wert = Wert
                                Case "Deckungsanteil-3"
                                    .Deckungsanteil_3.Wert = ImportInt(Wert)
                                Case "verschaerft-nach-EEWaermeG-7-1-2-eingehalten"
                                    .verschaerft_nach_EEWaermeG_7_1_2_eingehalten.Wert = ImportBool(Wert)
                                Case "verschaerft-nach-EEWaermeG-8"
                                    .verschaerft_nach_EEWaermeG_8.Wert = ImportInt(Wert)
                                Case "Primaerenergiebedarf-Hoechstwert-verschaerft"
                                    .Primaerenergiebedarf_Hoechstwert_verschaerft.Wert = ImportDec(Wert)
                                Case "spezifischer-Transmissionswaermetransferkoeffizient-verschaerft"
                                    .spezifischer_Transmissionswaermetransferkoeffizient_verschaerft.Wert = ImportDec(Wert)
                            '---------------------------------------------------------
                            'GEG 2020/2023
                                Case "verschaerft-nach-GEG-45-eingehalten"
                                    .verschaerft_nach_GEG_45_eingehalten.Wert = ImportBool(Wert)
                                Case "nicht-verschaerft-nach-GEG-34"
                                    .nicht_verschaerft_nach_GEG_34.Wert = ImportBool(Wert)
                                Case "verschaerft-nach-GEG-34"
                                    .verschaerft_nach_GEG_34.Wert = ImportInt(Wert)
                                Case "Anforderung-nach-GEG-16-unterschritten"
                                    .Anforderung_nach_GEG_16_unterschritten.Wert = ImportInt(Wert)
                                Case "spezifischer-Transmissionswaermeverlust-verschaerft"
                                    .spezifischer_Transmissionswaermetransferkoeffizient_verschaerft.Wert = ImportDec(Wert)
                            '---------------------------------------------------------
                            'GEG 2024
                            'in dem GEG 2024 gibt es keine EEWärme mehr für 4108-6
                            '---------------------------------------------------------
                            'Sommerlicher Wärmeschutz
                                Case "Sommerlicher-Waermeschutz"
                                    .Sommerlicher_Waermeschutz.Wert = ImportBool(Wert)
                            '---------------------------------------------------------
                            'Zusätzliche Verbrauchsdaten
                                Case "Treibhausgasemissionen-Zusaetzliche-Verbrauchsdaten"
                                    Variable_XML_Import.Energieausweis_Daten.Treibhausgasemissionen_Zusaetzliche_Verbrauchsdaten.Wert = ImportDec(Wert)
                                    '---------------------------------------------------------
                            End Select
                            '-----------------------------------------------------
                        End With
                        '--------------------------------------------------
                    Case "Verbrauchswerte", "Verbrauchsdaten"
                        '--------------------------------------------------
                        .Berechnungsverfahren = "Verbrauchsberechnung"
                        '--------------------------------------------------
                        With .Energieausweis_Daten.Wohngebaeude.Verbrauchswerte
                            '-----------------------------------------------------
                            Select Case Bezeichnung
                                Case "Flaechenermittlung-AN-aus-Wohnflaeche"
                                    .Flaechenermittlung_AN_aus_Wohnflaeche.Wert = ImportBool(Wert)
                                Case "Wohnflaeche"
                                    .Wohnflaeche.Wert = ImportInt(Wert)
                                Case "Keller-beheizt"
                                    .Keller_beheizt.Wert = ImportBool(Wert)
                                Case "Mittlerer-Endenergieverbrauch", "Endenergiekennwert-Verbrauch-AN"
                                    .Mittlerer_Endenergieverbrauch.Wert = ImportDec(Wert)
                                Case "Mittlerer-Primaerenergieverbrauch", "Primaerenergiekennwert-Verbrauch-AN"
                                    .Mittlerer_Primaerenergieverbrauch.Wert = ImportDec(Wert)
                                Case "Energieeffizienzklasse"
                                    .Energieeffizienzklasse.Wert = Wert
                            End Select
                            '-----------------------------------------------------
                        End With
                        '--------------------------------------------------
                    Case "Verbrauchswerte-NWG"
                        '--------------------------------------------------
                        .Berechnungsverfahren = "Verbrauchsberechnung"
                        '--------------------------------------------------
                        With .Energieausweis_Daten.Nichtwohngebaeude.Verbrauchswerte_NWG
                            '-----------------------------------------------------
                            Select Case Bezeichnung
                                Case "Warmwasser-enthalten"
                                    .Warmwasser_enthalten.Wert = ImportBool(Wert)
                                Case "Kuehlung-enthalten"
                                    .Kuehlung_enthalten.Wert = ImportBool(Wert)
                                Case "Endenergieverbrauch-Waerme"
                                    .Endenergieverbrauch_Waerme.Wert = ImportDec(Wert)
                                Case "Endenergieverbrauch-Strom"
                                    .Endenergieverbrauch_Strom.Wert = ImportDec(Wert)
                                Case "Endenergieverbrauch-Waerme-Vergleichswert"
                                    .Endenergieverbrauch_Waerme_Vergleichswert.Wert = ImportDec(Wert)
                                Case "Endenergieverbrauch-Strom-Vergleichswert"
                                    .Endenergieverbrauch_Strom_Vergleichswert.Wert = ImportDec(Wert)
                                Case "Primaerenergieverbrauch"
                                    .Primaerenergieverbrauch.Wert = ImportDec(Wert)
                                '-----------------------------------------------------
                                Case "Nutzung-1"
                                    Variable_XML_Import.Nutzung_Gebaeudekategorie(1).Gebaeudekategorie.Wert = Wert
                                Case "Flaechenanteil-Nutzung-1"
                                    Variable_XML_Import.Nutzung_Gebaeudekategorie(1).Flaechenanteil_Nutzung.Wert = ImportInt(Wert)
                                Case "Vergleichswert-Heizung-TWW-1"
                                    Variable_XML_Import.Nutzung_Gebaeudekategorie(1).Vergleichswert_Waerme.Wert = ImportDec(Wert)
                                Case "Vergleichswert-Strom-1"
                                    Variable_XML_Import.Nutzung_Gebaeudekategorie(1).Vergleichswert_Strom.Wert = ImportDec(Wert)
                                Case "Nutzung-2"
                                    Variable_XML_Import.Nutzung_Gebaeudekategorie(2).Gebaeudekategorie.Wert = Wert
                                Case "Flaechenanteil-Nutzung-2"
                                    Variable_XML_Import.Nutzung_Gebaeudekategorie(2).Flaechenanteil_Nutzung.Wert = ImportInt(Wert)
                                Case "Vergleichswert-Heizung-TWW-2"
                                    Variable_XML_Import.Nutzung_Gebaeudekategorie(2).Vergleichswert_Waerme.Wert = ImportDec(Wert)
                                Case "Vergleichswert-Strom-2"
                                    Variable_XML_Import.Nutzung_Gebaeudekategorie(2).Vergleichswert_Strom.Wert = ImportDec(Wert)
                                Case "Nutzung-3"
                                    Variable_XML_Import.Nutzung_Gebaeudekategorie(3).Gebaeudekategorie.Wert = Wert
                                Case "Flaechenanteil-Nutzung-3"
                                    Variable_XML_Import.Nutzung_Gebaeudekategorie(3).Flaechenanteil_Nutzung.Wert = ImportInt(Wert)
                                Case "Vergleichswert-Heizung-TWW-3"
                                    Variable_XML_Import.Nutzung_Gebaeudekategorie(3).Vergleichswert_Waerme.Wert = ImportDec(Wert)
                                Case "Vergleichswert-Strom-3"
                                    Variable_XML_Import.Nutzung_Gebaeudekategorie(3).Vergleichswert_Strom.Wert = ImportDec(Wert)
                            End Select
                            '-----------------------------------------------------
                        End With
                        '--------------------------------------------------
                    Case "Zusaetzliche-Verbrauchsdaten"
                        '--------------------------------------------------
                        Variable_XML_Import.Zusaetzliche_Verbrauchsberechnung = True
                        '--------------------------------------------------
                        With .Energieausweis_Daten.Wohngebaeude.Verbrauchswerte
                            '-----------------------------------------------------
                            Select Case Bezeichnung
                                Case "Flaechenermittlung-AN-aus-Wohnflaeche"
                                    .Flaechenermittlung_AN_aus_Wohnflaeche.Wert = ImportBool(Wert)
                                Case "Mittlerer-Endenergieverbrauch", "Endenergiekennwert-Verbrauch-AN"
                                    .Mittlerer_Endenergieverbrauch.Wert = ImportDec(Wert)
                                Case "Mittlerer-Primaerenergieverbrauch", "Primaerenergiekennwert-Verbrauch-AN"
                                    .Mittlerer_Primaerenergieverbrauch.Wert = ImportDec(Wert)
                                Case "Energieeffizienzklasse"
                                    .Energieeffizienzklasse.Wert = Wert
                            End Select
                            '-----------------------------------------------------
                        End With
                        '--------------------------------------------------
                        With .Energieausweis_Daten.Nichtwohngebaeude.Verbrauchswerte_NWG
                            '-----------------------------------------------------
                            Select Case Bezeichnung
                                Case "Warmwasser-enthalten"
                                    .Warmwasser_enthalten.Wert = ImportBool(Wert)
                                Case "Kuehlung-enthalten"
                                    .Kuehlung_enthalten.Wert = ImportBool(Wert)
                                Case "Endenergieverbrauch-Waerme"
                                    .Endenergieverbrauch_Waerme.Wert = ImportDec(Wert)
                                Case "Endenergieverbrauch-Strom"
                                    .Endenergieverbrauch_Strom.Wert = ImportDec(Wert)
                                Case "Endenergieverbrauch-Waerme-Vergleichswert"
                                    .Endenergieverbrauch_Waerme_Vergleichswert.Wert = ImportDec(Wert)
                                Case "Endenergieverbrauch-Strom-Vergleichswert"
                                    .Endenergieverbrauch_Strom_Vergleichswert.Wert = ImportDec(Wert)
                                Case "Primaerenergieverbrauch"
                                    .Primaerenergieverbrauch.Wert = ImportDec(Wert)
                                '-----------------------------------------------------
                                Case "Nutzung-1"
                                    Variable_XML_Import.Nutzung_Gebaeudekategorie(1).Gebaeudekategorie.Wert = ImportHauptnutzung(Wert)
                                Case "Flaechenanteil-Nutzung-1"
                                    Variable_XML_Import.Nutzung_Gebaeudekategorie(1).Flaechenanteil_Nutzung.Wert = ImportInt(Wert)
                                Case "Vergleichswert-Heizung-TWW-1"
                                    Variable_XML_Import.Nutzung_Gebaeudekategorie(1).Vergleichswert_Waerme.Wert = ImportDec(Wert)
                                Case "Vergleichswert-Strom-1"
                                    Variable_XML_Import.Nutzung_Gebaeudekategorie(1).Vergleichswert_Strom.Wert = ImportDec(Wert)
                                Case "Nutzung-2"
                                    Variable_XML_Import.Nutzung_Gebaeudekategorie(2).Gebaeudekategorie.Wert = ImportHauptnutzung(Wert)
                                Case "Flaechenanteil-Nutzung-2"
                                    Variable_XML_Import.Nutzung_Gebaeudekategorie(2).Flaechenanteil_Nutzung.Wert = ImportInt(Wert)
                                Case "Vergleichswert-Heizung-TWW-2"
                                    Variable_XML_Import.Nutzung_Gebaeudekategorie(2).Vergleichswert_Waerme.Wert = ImportDec(Wert)
                                Case "Vergleichswert-Strom-2"
                                    Variable_XML_Import.Nutzung_Gebaeudekategorie(2).Vergleichswert_Strom.Wert = ImportDec(Wert)
                                Case "Nutzung-3"
                                    Variable_XML_Import.Nutzung_Gebaeudekategorie(3).Gebaeudekategorie.Wert = ImportHauptnutzung(Wert)
                                Case "Flaechenanteil-Nutzung-3"
                                    Variable_XML_Import.Nutzung_Gebaeudekategorie(3).Flaechenanteil_Nutzung.Wert = ImportInt(Wert)
                                Case "Vergleichswert-Heizung-TWW-3"
                                    Variable_XML_Import.Nutzung_Gebaeudekategorie(3).Vergleichswert_Waerme.Wert = ImportDec(Wert)
                                Case "Vergleichswert-Strom-3"
                                    Variable_XML_Import.Nutzung_Gebaeudekategorie(3).Vergleichswert_Strom.Wert = ImportDec(Wert)
                            End Select
                            '-----------------------------------------------------
                        End With
                       '--------------------------------------------------
                    Case "Bedarfswerte-easy"
                        '--------------------------------------------------
                        .Berechnungsverfahren = "Bedarfsberechnung-easy"
                        '--------------------------------------------------
                        With .Energieausweis_Daten.Wohngebaeude.Bedarfswerte_easy
                            '-----------------------------------------------------
                            Select Case Bezeichnung
                                Case "Wohngebaeude-Anbaugrad"
                                    .Wohngebaeude_Anbaugrad.Wert = Wert
                                Case "Anzahl-Geschosse"
                                    .Anzahl_Geschosse.Wert = ImportInt(Wert)
                                Case "Geschoss-Bruttogeschossflaechenumfang"
                                    .Geschoss_Bruttogeschossflaechenumfang.Wert = ImportInt(Wert)
                                Case "Geschoss-Bruttogeschossflaeche"
                                    .Geschoss_Bruttogeschossflaeche.Wert = ImportInt(Wert)
                                Case "Dach-Bruttogeschossflaechenumfang"
                                    .Dach_Bruttogeschossflaechenumfang.Wert = ImportInt(Wert)
                                Case "Dach-Bruttogeschossflaeche"
                                    .Dach_Bruttogeschossflaeche.Wert = ImportInt(Wert)
                                Case "Aufsummierte-Bruttogeschossflaeche"
                                    .Aufsummierte_Bruttogeschossflaeche.Wert = ImportInt(Wert)
                                Case "Mittlere-Geschosshoehe"
                                    .Mittlere_Geschosshoehe.Wert = ImportDec(Wert)
                                Case "Kompaktheit"
                                    .Kompaktheit.Wert = ImportBool(Wert)
                                Case "Deckungsgleichheit"
                                    .Deckungsgleichheit.Wert = ImportBool(Wert)
                                Case "Fensterflaechenanteil-Nordost-Nord-Nordwest"
                                    .Fensterflaechenanteil_Nordost_Nord_Nordwest.Wert = ImportInt(Wert)
                                Case "Fensterflaechenanteil-Gesamt"
                                    .Fensterflaechenanteil_Gesamt.Wert = ImportInt(Wert)
                                Case "Dach-transparente-Bauteile-Fensterflaechenanteil"
                                    .Dach_transparente_Bauteile_Fensterflaechenanteil.Wert = ImportDec(Wert)
                                Case "Spezielle-Fenstertueren-Flaechenanteil"
                                    .Spezielle_Fenstertueren_Flaechenanteil.Wert = ImportDec(Wert)
                                Case "Außentueren-Flaechenanteil"
                                    .Außentueren_Flaechenanteil.Wert = ImportDec(Wert)
                                Case "Keine-Anlage-zur-Kuehlung"
                                    .Keine_Anlage_zur_Kuehlung.Wert = ImportBool(Wert)
                                Case "Anforderung-Waermebruecken-erfuellt"
                                    .Anforderung_Waermebruecken_erfuellt.Wert = ImportBool(Wert)
                                Case "Gebaeudedichtheit"
                                    .Gebaeudedichtheit.Wert = ImportBool(Wert)
                                Case "Heiz-Warmwassersystem"
                                    .Heiz_Warmwassersystem.Wert = Wert
                                Case "Lueftungsanlagenanforderungen"
                                    .Lueftungsanlagenanforderungen.Wert = ImportBool(Wert)
                                Case "Waermeschutz-Variante"
                                    .Waermeschutz_Variante.Wert = Wert
                                Case "Endenergiebedarf"
                                    .Endenergiebedarf.Wert = ImportDec(Wert)
                                Case "Energieeffizienzklasse"
                                    .Energieeffizienzklasse.Wert = Wert
                                Case "Primaerenergiebedarf-Ist-Wert"
                                    .Primaerenergiebedarf_Ist_Wert.Wert = ImportDec(Wert)
                                Case "Primaerenergiebedarf-Anforderungswert"
                                    .Primaerenergiebedarf_Anforderungswert.Wert = ImportDec(Wert)
                                Case "Energetische-Qualitaet-Ist-Wert"
                                    .Energetische_Qualitaet_Ist_Wert.Wert = ImportDec(Wert)
                                Case "Energetische-Qualitaet-Anforderungs-Wert"
                                    .Energetische_Qualitaet_Anforderungs_Wert.Wert = ImportDec(Wert)
                                    '---------------------------------------------------------
                            'EnEV 2016 / GEG 2020/2023
                                Case "Art-der-Nutzung-erneuerbaren-Energie-1", "EEWaermeG-Nutzung-1"
                                    .Art_der_Nutzung_erneuerbaren_Energie_1.Wert = Wert
                                Case "Deckungsanteil-1"
                                    .Deckungsanteil_1.Wert = ImportInt(Wert)
                                Case "Anteil-der-Pflichterfuellung-1"
                                    .Anteil_der_Pflichterfuellung_1.Wert = ImportInt(Wert)
                                Case "Art-der-Nutzung-erneuerbaren-Energie-2", "EEWaermeG-Nutzung-2"
                                    .Art_der_Nutzung_erneuerbaren_Energie_2.Wert = Wert
                                Case "Deckungsanteil-2"
                                    .Deckungsanteil_2.Wert = ImportInt(Wert)
                                Case "Anteil-der-Pflichterfuellung-2"
                                    .Anteil_der_Pflichterfuellung_2.Wert = ImportInt(Wert)
                                    '---------------------------------------------------------
                                    'EnEV 2016
                                Case "EEWaermeG-Nutzung-3"
                                    .Art_der_Nutzung_erneuerbaren_Energie_3.Wert = Wert
                                Case "Deckungsanteil-3"
                                    .Deckungsanteil_3.Wert = ImportInt(Wert)
                                    '---------------------------------------------------------
                                    'GEG 2024
                            End Select
                            '-----------------------------------------------------
                        End With
                        '--------------------------------------------------
                End Select
                '--------------------------------------------------
                Select Case Ebene_4
                    Case "Bauteil-Opak"
                        '--------------------------------------------------
                        If Opak_Nr > 0 Then
                            '-----------------------------------------------------
                            With .BauteilOpak(Opak_Nr)
                                '-----------------------------------------------------
                                Select Case Bezeichnung
                                    Case "Flaechenbezeichnung"
                                        .Flaechenbezeichnung.Wert = Wert
                                    Case "Flaeche"
                                        .Flaeche.Wert = ImportInt(Wert)
                                    Case "U-Wert"
                                        .U_Wert.Wert = ImportDec(Wert)
                                    Case "Ausrichtung"
                                        .Ausrichtung.Wert = Wert
                                    Case "grenztAn"
                                        .grenztAn.Wert = Wert
                                    Case "Glasdach-Lichtband-Lichtkuppel"
                                        .Glasdach_Lichtband_Lichtkuppel.Wert = ImportBool(Wert)
                                    Case "Vorhangfassade"
                                        .Vorhangfassade.Wert = ImportBool(Wert)
                                End Select
                                '-----------------------------------------------------
                            End With
                            '-----------------------------------------------------
                        End If
                        '--------------------------------------------------
                    Case "Bauteil-Transparent"
                        '--------------------------------------------------
                        If Transparent_Nr > 0 Then
                            '-----------------------------------------------------
                            With .BauteilTransparent(Transparent_Nr)
                                '-----------------------------------------------------
                                Select Case Bezeichnung
                                    Case "Flaechenbezeichnung"
                                        .Flaechenbezeichnung.Wert = Wert
                                    Case "Flaeche"
                                        .Flaeche.Wert = ImportInt(Wert)
                                    Case "U-Wert"
                                        .U_Wert.Wert = ImportDec(Wert)
                                    Case "g-Wert"
                                        .g_Wert.Wert = ImportDec(Wert)
                                    Case "Ausrichtung"
                                        .Ausrichtung.Wert = Wert
                                    Case "Glasdach-Lichtband-Lichtkuppel"
                                        .Glasdach_Lichtband_Lichtkuppel.Wert = ImportBool(Wert)
                                    Case "Vorhangfassade"
                                        .Vorhangfassade.Wert = ImportBool(Wert)
                                End Select
                                '-----------------------------------------------------
                            End With
                            '-----------------------------------------------------
                        End If
                        '--------------------------------------------------
                    Case "Bauteil-Dach"
                        '--------------------------------------------------
                        If Dach_Nr > 0 Then
                            '-----------------------------------------------------
                            With .BauteilDach(Dach_Nr)
                                '-----------------------------------------------------
                                Select Case Bezeichnung
                                    Case "Flaechenbezeichnung"
                                        .Flaechenbezeichnung.Wert = Wert
                                    Case "Flaeche"
                                        .Flaeche.Wert = ImportInt(Wert)
                                    Case "U-Wert"
                                        .U_Wert.Wert = ImportDec(Wert)
                                End Select
                                '-----------------------------------------------------
                            End With
                            '-----------------------------------------------------
                        End If
                        '--------------------------------------------------
                    Case "Zone"
                        '--------------------------------------------------
                        If Zone_Nr > 0 Then
                            '-----------------------------------------------------
                            With .Energieausweis_Daten.Nichtwohngebaeude.Zone(Zone_Nr)
                                '-----------------------------------------------------
                                Select Case Bezeichnung
                                    Case "Zonenbezeichnung"
                                        .Zonenbezeichnung.Wert = Wert
                                    Case "Nutzung"
                                        .Nutzung.Wert = ImportHauptnutzung(Wert)
                                    Case "Anwenderspezifische-Nutzung-Bezeichnung"
                                        .Anwenderspezifische_Nutzung_Bezeichnung.Wert = Wert
                                    Case "Zonenbesonderheiten"
                                        .Zonenbesonderheiten.Wert = Wert
                                    Case "Nettogrundflaeche-Zone"
                                        .Nettogrundflaeche_Zone.Wert = ImportInt(Wert)
                                    Case "mittlere-lichte-Raumhoehe"
                                        .mittlere_lichte_Raumhoehe.Wert = ImportDec(Wert)
                                    Case "Sonnenschutz-System"
                                        .Sonnenschutz_System.Wert = Wert
                                    Case "Beleuchtungs-System"
                                        .Beleuchtungs_System.Wert = Wert
                                    Case "Beleuchtungs-Verteilung"
                                        .Beleuchtungs_Verteilung.Wert = Wert
                                    Case "raesenzkontrolle-Kunstlicht"
                                        .Praesenzkontrolle_Kunstlicht.Wert = ImportBool(Wert)
                                    Case "Tageslichtabhaengige-Kontrollsysteme"
                                        .Tageslichtabhaengige_Kontrollsysteme.Wert = ImportBool(Wert)
                                    Case "Endenergiebedarf-Heizung"
                                        .Endenergiebedarf_Heizung.Wert = ImportInt(Wert)
                                    Case "Endenergiebedarf-Kuehlung"
                                        .Endenergiebedarf_Kuehlung.Wert = ImportInt(Wert)
                                    Case "Endenergiebedarf-Befeuchtung"
                                        .Endenergiebedarf_Befeuchtung.Wert = ImportInt(Wert)
                                    Case "Endenergiebedarf-Trinkwarmwasser"
                                        .Endenergiebedarf_Trinkwarmwasser.Wert = ImportInt(Wert)
                                    Case "Endenergiebedarf-Beleuchtung"
                                        .Endenergiebedarf_Beleuchtung.Wert = ImportInt(Wert)
                                    Case "Endenergiebedarf-Lufttransport"
                                        .Endenergiebedarf_Lufttransport.Wert = ImportInt(Wert)
                                    Case "Endenergiebedarf-Hilfsenergie"
                                        .Endenergiebedarf_Hilfsenergie.Wert = ImportInt(Wert)
                                End Select
                                '-----------------------------------------------------
                            End With
                            '-----------------------------------------------------
                        End If
                        '--------------------------------------------------
                    Case "Heizungsanlage", "Heizsystem"
                        '--------------------------------------------------
                        If Heizungsanlage_Nr > 0 Then
                            '-----------------------------------------------------
                            With .Heizungsanlage(Heizungsanlage_Nr)
                                '-----------------------------------------------------
                                Select Case Bezeichnung
                                    Case "Waermeerzeuger-Bauweise-4701"
                                        .Waermeerzeuger_Bauweise_4701.Wert = Wert
                                    Case "Waermeerzeuger-Bauweise-18599"
                                        .Waermeerzeuger_Bauweise_18599.Wert = Wert
                                    Case "Nennleistung"
                                        .Nennleistung.Wert = ImportInt(Wert)
                                    Case "Waermeerzeuger-Baujahr"
                                        .Waermeerzeuger_Baujahr.Wert = ImportInt(Wert)
                                    Case "Anzahl-baugleiche"
                                        .Anzahl_baugleiche.Wert = ImportInt(Wert)
                                    Case "Energietraeger"
                                        .Energietraeger.Wert = Wert
                                    Case "Primaerenergiefaktor"
                                        .Primaerenergiefaktor.Wert = ImportDec(Wert)
                                    Case "Emissionsfaktor"
                                        .Emissionsfaktor.Wert = ImportInt(Wert)
                                End Select
                                '-----------------------------------------------------
                            End With
                            '-----------------------------------------------------
                        End If
                        '--------------------------------------------------
                    Case "Trinkwarmwasseranlage", "Warmwasserbereitungssystem"
                        '--------------------------------------------------
                        If Trinkwarmwasseranlage_Nr > 0 Then
                            '-----------------------------------------------------
                            With .Trinkwarmwasseranlage(Trinkwarmwasseranlage_Nr)
                                '-----------------------------------------------------
                                Select Case Bezeichnung
                                    Case "Trinkwarmwassererzeuger-Bauweise-4701"
                                        .Trinkwarmwassererzeuger_Bauweise_4701.Wert = Wert
                                    Case "Trinkwarmwassererzeuger-Bauweise-18599"
                                        .Trinkwarmwassererzeuger_Bauweise_18599.Wert = Wert
                                    Case "Trinkwarmwassererzeuger-Baujahr"
                                        .Trinkwarmwassererzeuger_Baujahr.Wert = ImportInt(Wert)
                                    Case "Anzahl-baugleiche"
                                        .Anzahl_baugleiche.Wert = ImportInt(Wert)
                                End Select
                                '-----------------------------------------------------
                            End With
                            '-----------------------------------------------------
                        End If
                        '--------------------------------------------------
                    Case "Kaelteanlage"
                        '--------------------------------------------------
                        If Klima_Nr > 0 Then
                            '-----------------------------------------------------
                            With .Kaelteanlage(Klima_Nr)
                                '-----------------------------------------------------
                                Select Case Bezeichnung
                                    Case "Kaelteerzeuger-Bauweise"
                                        .Kaelteerzeuger_Bauweise.Wert = Wert
                                    Case "Kaelteerzeuger-Regelung"
                                        .Kaelteerzeuger_Regelung.Wert = Wert
                                    Case "Rueckkuehlung-Bauweise"
                                        .Rueckkuehlung_Bauweise.Wert = Wert
                                    Case "Kaelteverteilung-Primaerkreis-Temperatur"
                                        .Kaelteverteilung_Primaerkreis_Temperatur.Wert = Wert
                                    Case "Nennkaelteleistung"
                                        .Nennkaelteleistung.Wert = ImportInt(Wert)
                                    Case "Kaelteerzeuger-Baujahr"
                                        .Kaelteerzeuger_Baujahr.Wert = ImportInt(Wert)
                                    Case "Anzahl-baugleiche"
                                        .Anzahl_baugleiche.Wert = ImportInt(Wert)
                                    Case "Energietraeger"
                                        .Energietraeger.Wert = Wert
                                    Case "Primaerenergiefaktor"
                                        .Primaerenergiefaktor.Wert = ImportDec(Wert)
                                End Select
                                '-----------------------------------------------------
                            End With
                            '-----------------------------------------------------
                        End If
                        '--------------------------------------------------
                    Case "RLT-System"
                        '--------------------------------------------------
                        If Lueftung_Nr > 0 Then
                            '-----------------------------------------------------
                            With .RLT_System(Lueftung_Nr)
                                '-----------------------------------------------------
                                Select Case Bezeichnung
                                    Case "Funktion-Zuluft"
                                        .Funktion_Zuluft.Wert = ImportBool(Wert)
                                    Case "Funktion-Abluft"
                                        .Funktion_Abluft.Wert = ImportBool(Wert)
                                    Case "WRG-Rueckwaermzahl"
                                        .WRG_Rueckwaermzahl.Wert = ImportInt(Wert)
                                    Case "Funktion-Heizregister"
                                        .Funktion_Heizregister.Wert = ImportBool(Wert)
                                    Case "Funktion-Kuehlregister"
                                        .Funktion_Kuehlregister.Wert = ImportBool(Wert)
                                    Case "Funktion-Dampfbefeuchter"
                                        .Funktion_Dampfbefeuchter.Wert = ImportBool(Wert)
                                    Case "Funktion-Wasserbefeuchter"
                                        .Funktion_Wasserbefeuchter.Wert = ImportBool(Wert)
                                    Case "Energietraeger-Befeuchtung"
                                        .Energietraeger_Befeuchtung.Wert = Wert
                                    Case "Anzahl-baugleiche"
                                        .Anzahl_baugleiche.Wert = ImportInt(Wert)
                                End Select
                                '-----------------------------------------------------
                            End With
                            '-----------------------------------------------------
                        End If
                        '--------------------------------------------------
                    Case "Energietraeger-Liste"
                        '--------------------------------------------------
                        If Energietraeger_Nr > 0 Then
                            '-----------------------------------------------------
                            With .Energietraeger(Energietraeger_Nr)
                                '-----------------------------------------------------
                                Select Case Bezeichnung
                                    Case "Energietraegerbezeichnung"
                                        .Energietraegerbezeichnung.Wert = Wert
                                    Case "Primaerenergiefaktor"
                                        .Primaerenergiefaktor.Wert = ImportDec(Wert)
                                    Case "Endenergiebedarf-Heizung-spezifisch"
                                        .Endenergiebedarf_Heizung_spezifisch.Wert = ImportDec(Wert)
                                    Case "Endenergiebedarf-Kuehlung-Befeuchtung-spezifisch"
                                        .Endenergiebedarf_Kuehlung_Befeuchtung_spezifisch.Wert = ImportDec(Wert)
                                    Case "Endenergiebedarf-Trinkwarmwasser-spezifisch"
                                        .Endenergiebedarf_Trinkwarmwasser_spezifisch.Wert = ImportDec(Wert)
                                    Case "Endenergiebedarf-Beleuchtung-spezifisch"
                                        .Endenergiebedarf_Beleuchtung_spezifisch.Wert = ImportDec(Wert)
                                    Case "Endenergiebedarf-Lueftung-spezifisch"
                                        .Endenergiebedarf_Lueftung_spezifisch.Wert = ImportDec(Wert)
                                    Case "Endenergiebedarf-Energietraeger-Gesamtgebaeude-spezifisch"
                                        .Endenergiebedarf_Energietraeger_Gesamtgebaeude_spezifisch.Wert = ImportDec(Wert)
                                End Select
                                '-----------------------------------------------------
                            End With
                            '-----------------------------------------------------
                        End If
                        '--------------------------------------------------
                    Case "Angaben-erneuerbare-Energien-65EE-Regel"
                        '--------------------------------------------------
                        If EE_65EE_Regel_Nr > 0 Then
                            '-----------------------------------------------------
                            With .EE_65EE_Regel(EE_65EE_Regel_Nr)
                                '-----------------------------------------------------
                                Select Case Bezeichnung
                                    Case "Waermeerzeuger-Bauweise-18599"
                                        .Waermeerzeuger_Bauweise_18599.Wert = Wert
                                    Case "Art-der-Nutzung-erneuerbaren-Energie"
                                        .Art_der_Nutzung_erneuerbaren_Energie.Wert = Wert
                                    Case "Deckungsanteil"
                                        .Deckungsanteil.Wert = ImportInt(Wert)
                                    Case "Anteil-der-Pflichterfuellung-Anlage"
                                        .Anteil_der_Pflichterfuellung_Anlage.Wert = ImportInt(Wert)
                                    Case "Anteil-der-Pflichterfuellung-Gesamt"
                                        .Anteil_der_Pflichterfuellung_Gesamt.Wert = ImportInt(Wert)
                                End Select
                                '-----------------------------------------------------
                            End With
                            '-----------------------------------------------------
                        End If
                        '--------------------------------------------------
                    Case "Angaben-erneuerbare-Energien-keine-65EE-Regel"
                        '--------------------------------------------------
                        If EE_65EE_keine_Regel_Nr > 0 Then
                            '-----------------------------------------------------
                            With .EE_65EE_keine_Regel(EE_65EE_keine_Regel_Nr)
                                '-----------------------------------------------------
                                Select Case Bezeichnung
                                    Case "Waermeerzeuger-Bauweise-18599"
                                        .Waermeerzeuger_Bauweise_18599.Wert = Wert
                                    Case "Art-der-Nutzung-erneuerbaren-Energie"
                                        .Art_der_Nutzung_erneuerbaren_Energie.Wert = Wert
                                    Case "Anteil-EE-Anlage"
                                        .Anteil_EE_Anlage.Wert = ImportInt(Wert)
                                End Select
                                '-----------------------------------------------------
                            End With
                            '-----------------------------------------------------
                        End If
                        '--------------------------------------------------
                    Case "Zeitraum-Strom"
                        '--------------------------------------------------
                        If Zeitraum_Strom_Daten_Nr > 0 Then
                            '-----------------------------------------------------
                            With .Zeitraum_Strom_Daten(Zeitraum_Strom_Daten_Nr)
                                '-----------------------------------------------------
                                Select Case Bezeichnung
                                    Case "Startdatum"
                                        .Startdatum.Wert = CDate(ImportDate(Wert))
                                    Case "Enddatum"
                                        .Enddatum.Wert = CDate(ImportDate(Wert))
                                    Case "Energieverbrauch-Strom"
                                        .Energieverbrauch_Strom.Wert = ImportInt(Wert)
                                    Case "Energieverbrauchsanteil-elektrisch-erzeugte-Kaelte"
                                        .Energieverbrauchsanteil_elektrisch_erzeugte_Kaelte.Wert = ImportInt(Wert)
                                End Select
                                '-----------------------------------------------------
                            End With
                            '-----------------------------------------------------
                        End If
                        '--------------------------------------------------
                    Case "Nutzung-Gebaeudekategorie"
                        '--------------------------------------------------
                        If Nutzung_Gebaeudekategorie_Nr > 0 Then
                            '-----------------------------------------------------
                            With .Nutzung_Gebaeudekategorie(Nutzung_Gebaeudekategorie_Nr)
                                '-----------------------------------------------------
                                Select Case Bezeichnung
                                    Case "Gebaeudekategorie"
                                        .Gebaeudekategorie.Wert = ImportHauptnutzung(Wert)
                                    Case "Flaechenanteil-Nutzung"
                                        .Flaechenanteil_Nutzung.Wert = ImportInt(Wert)
                                    Case "Vergleichswert-Waerme"
                                        .Vergleichswert_Waerme.Wert = ImportDec(Wert)
                                    Case "Vergleichswert-Strom"
                                        .Vergleichswert_Strom.Wert = ImportDec(Wert)
                                End Select
                                '-----------------------------------------------------
                            End With
                            '-----------------------------------------------------
                        End If
                        '--------------------------------------------------
                    Case "Modernisierungsempfehlungen"
                        '--------------------------------------------------
                        If Modernisierungsempfehlungen_Nr > 0 Then
                            '-----------------------------------------------------
                            With .Modernisierungsempfehlungen(Modernisierungsempfehlungen_Nr)
                                '-----------------------------------------------------
                                Select Case Bezeichnung
                                    Case "Nummer"
                                        .Nummer.Wert = Wert
                                    Case "Bauteil-Anlagenteil"
                                        .Bauteil_Anlagenteil.Wert = Wert
                                    Case "Massnahmenbeschreibung"
                                        .Massnahmenbeschreibung.Wert = Wert
                                    Case "Modernisierungskombination"
                                        .Modernisierungskombination.Wert = Wert
                                    Case "Amortisation"
                                        .Amortisation.Wert = Wert
                                    Case "spezifische-Kosten"
                                        .spezifische_Kosten.Wert = Wert
                                End Select
                                '-----------------------------------------------------
                            End With
                            '-----------------------------------------------------
                        End If
                        '--------------------------------------------------
                    Case "Leerstandszuschlag-Heizung"
                        '--------------------------------------------------
                        With .Leerstandszuschlag_Heizung
                            '-----------------------------------------------------
                            Select Case Bezeichnung
                                Case "kein-Leerstand"
                                    .kein_Leerstand.Wert = Wert
                                Case "Leerstandsfaktor"
                                    .Leerstandsfaktor.Wert = ImportDec(Wert)
                                Case "Startdatum"
                                    .Startdatum.Wert = CDate(ImportDate(Wert))
                                Case "Enddatum"
                                    .Enddatum.Wert = CDate(ImportDate(Wert))
                                Case "Leerstandszuschlag-kWh"
                                    .Leerstandszuschlag_kWh.Wert = ImportInt(Wert)
                                Case "Primaerenergiefaktor"
                                    .Primaerenergiefaktor.Wert = ImportDec(Wert)
                                Case "Zuschlagsfaktor"
                                    .Zuschlagsfaktor.Wert = ImportDec(Wert)
                                Case "witterungsbereinigter-Endenergieverbrauchsanteil-fuer-Heizung"
                                    .witterungsbereinigter_Endenergieverbrauchsanteil_fuer_Heizung.Wert = ImportDec(Wert)
                            End Select
                            '-----------------------------------------------------
                        End With
                        '--------------------------------------------------
                    Case "Leerstandszuschlag-Warmwasser"
                        '--------------------------------------------------
                        With .Leerstandszuschlag_Warmwasser
                            '-----------------------------------------------------
                            Select Case Bezeichnung
                                Case "keine-Nutzung-von-WW"
                                    .keine_Nutzung_von_WW.Wert = ImportBool(Wert)
                                Case "kein-Leerstand"
                                    .kein_Leerstand.Wert = Wert
                                Case "Leerstandsfaktor"
                                    .Leerstandsfaktor.Wert = ImportDec(Wert)
                                Case "Startdatum"
                                    .Startdatum.Wert = CDate(ImportDate(Wert))
                                Case "Enddatum"
                                    .Enddatum.Wert = CDate(ImportDate(Wert))
                                Case "Leerstandszuschlag-kWh"
                                    .Leerstandszuschlag_kWh.Wert = ImportInt(Wert)
                                Case "Primaerenergiefaktor"
                                    .Primaerenergiefaktor.Wert = ImportDec(Wert)
                            End Select
                            '-----------------------------------------------------
                        End With
                        '--------------------------------------------------
                    Case "Leerstandszuschlag-thermisch-erzeugte-Kaelte"
                        '--------------------------------------------------
                        With .Leerstandszuschlag_thermisch_erzeugte_Kaelte
                            '-----------------------------------------------------
                            Select Case Bezeichnung
                                Case "kein-Leerstand"
                                    .kein_Leerstand.Wert = Wert
                                Case "Leerstandsfaktor"
                                    .Leerstandsfaktor.Wert = ImportDec(Wert)
                                Case "Startdatum"
                                    .Startdatum.Wert = CDate(ImportDate(Wert))
                                Case "Enddatum"
                                    .Enddatum.Wert = CDate(ImportDate(Wert))
                                Case "Leerstandszuschlag-kWh"
                                    .Leerstandszuschlag_kWh.Wert = ImportInt(Wert)
                                Case "Primaerenergiefaktor"
                                    .Primaerenergiefaktor.Wert = ImportDec(Wert)
                            End Select
                            '-----------------------------------------------------
                        End With
                        '--------------------------------------------------
                    Case "Leerstandszuschlag-Strom"
                        '--------------------------------------------------
                        With .Leerstandszuschlag_Strom
                            '-----------------------------------------------------
                            Select Case Bezeichnung
                                Case "kein-Leerstand"
                                    .kein_Leerstand.Wert = Wert
                                Case "Leerstandsfaktor"
                                    .Leerstandsfaktor.Wert = ImportDec(Wert)
                                Case "Startdatum"
                                    .Startdatum.Wert = CDate(ImportDate(Wert))
                                Case "Enddatum"
                                    .Enddatum.Wert = CDate(ImportDate(Wert))
                                Case "Leerstandszuschlag-kWh"
                                    .Leerstandszuschlag_kWh.Wert = ImportInt(Wert)
                                Case "Primaerenergiefaktor"
                                    .Primaerenergiefaktor.Wert = ImportDec(Wert)
                            End Select
                            '-----------------------------------------------------
                        End With
                        '--------------------------------------------------
                    Case "Warmwasserzuschlag"
                        '--------------------------------------------------
                        With .Warmwasserzuschlag
                            '-----------------------------------------------------
                            Select Case Bezeichnung
                                Case "Startdatum"
                                    .Startdatum.Wert = CDate(ImportDate(Wert))
                                Case "Enddatum"
                                    .Enddatum.Wert = CDate(ImportDate(Wert))
                                Case "Primaerenergiefaktor"
                                    .Primaerenergiefaktor.Wert = ImportDec(Wert)
                                Case "Warmwasserzuschlag-kWh"
                                    .Warmwasserzuschlag_kWh.Wert = ImportInt(Wert)
                            End Select
                            '-----------------------------------------------------
                        End With
                        '--------------------------------------------------
                    Case "Kuehlzuschlag"
                        '--------------------------------------------------
                        With .Kuehlzuschlag
                            '-----------------------------------------------------
                            Select Case Bezeichnung
                                Case "Startdatum"
                                    .Startdatum.Wert = CDate(ImportDate(Wert))
                                Case "Enddatum"
                                    .Enddatum.Wert = CDate(ImportDate(Wert))
                                Case "Gebaeudenutzflaeche-gekuehlt"
                                    .Gebaeudenutzflaeche_gekuehlt.Wert = ImportInt(Wert)
                                Case "Primaerenergiefaktor"
                                    .Primaerenergiefaktor.Wert = ImportDec(Wert)
                                Case "Kuehlzuschlag-kWh"
                                    .Kuehlzuschlag_kWh.Wert = ImportInt(Wert)
                            End Select
                            '-----------------------------------------------------
                        End With
                        '--------------------------------------------------
                    Case "Strom-Daten"
                        '--------------------------------------------------
                        With .Energieausweis_Daten.Nichtwohngebaeude.Verbrauchswerte_NWG.Strom_Daten
                            '-----------------------------------------------------
                            Select Case Bezeichnung
                                Case "Stromverbrauch-enthaelt-Zusatzheizung"
                                    .Stromverbrauch_enthaelt_Zusatzheizung.Wert = ImportBool(Wert)
                                Case "Stromverbrauch-enthaelt-Warmwasser"
                                    .Stromverbrauch_enthaelt_Warmwasser.Wert = ImportBool(Wert)
                                Case "Stromverbrauch-enthaelt-Lueftung"
                                    .Stromverbrauch_enthaelt_Lueftung.Wert = ImportBool(Wert)
                                Case "Stromverbrauch-enthaelt-Beleuchtung"
                                    .Stromverbrauch_enthaelt_Beleuchtung.Wert = ImportBool(Wert)
                                Case "Stromverbrauch-enthaelt-Kuehlung"
                                    .Stromverbrauch_enthaelt_Kuehlung.Wert = ImportBool(Wert)
                                Case "Stromverbrauch-enthaelt-Sonstiges"
                                    .Stromverbrauch_enthaelt_Sonstiges.Wert = ImportBool(Wert)
                            End Select
                            '-----------------------------------------------------
                        End With
                        '--------------------------------------------------
                    Case "Energietraeger"
                        '--------------------------------------------------
                        If Energietraeger_Daten_Nr > 0 Then
                            '-----------------------------------------------------
                            With .Energietraeger_Daten(Energietraeger_Daten_Nr)
                                '-----------------------------------------------------
                                Select Case Bezeichnung
                                    Case "Energietraeger-Verbrauch", "Energietraegerbezeichnung"
                                        .Energietraeger_Verbrauch.Wert = Wert
                                    Case "Sonstiger-Energietraeger-Verbrauch"
                                        .Sonstiger_Energietraeger_Verbrauch.Wert = Wert
                                    Case "Unterer-Heizwert", "Umrechnungsfaktor"
                                        .Unterer_Heizwert.Wert = ImportDec(Wert)
                                    Case "Primaerenergiefaktor"
                                        .Primaerenergiefaktor.Wert = ImportDec(Wert)
                                    Case "Emissionsfaktor"
                                        .Emissionsfaktor.Wert = ImportInt(Wert)
                                End Select
                                '-----------------------------------------------------
                                Select Case Ebene_5
                                    Case "Zeitraum", "Verbrauchsperiode"
                                        '-----------------------------------------------------
                                        If Zeitraum_Daten_Nr > 0 Then
                                            '-----------------------------------------------------
                                            With .Zeitraum_Daten(Zeitraum_Daten_Nr)
                                                '-----------------------------------------------------
                                                Select Case Bezeichnung
                                                    Case "Startdatum"
                                                        .Startdatum.Wert = CDate(ImportDate(Wert))
                                                    Case "Enddatum"
                                                        .Enddatum.Wert = CDate(ImportDate(Wert))
                                                    Case "Verbrauchte-Menge", "Verbrauchswert"
                                                        .Verbrauchte_Menge.Wert = ImportInt(Wert)
                                                    Case "Energieverbrauch", "Verbrauchswert-kWh-Heizwert"
                                                        .Energieverbrauch.Wert = ImportInt(Wert)
                                                    Case "Energieverbrauchsanteil-Warmwasser-zentral", "Warmwasserwert"
                                                        .Energieverbrauchsanteil_Warmwasser_zentral.Wert = ImportInt(Wert)
                                                    Case "Warmwasserwertermittlung"
                                                        .Warmwasserwertermittlung.Wert = Wert
                                                    Case "Energieverbrauchsanteil-thermisch-erzeugte-Kaelte"
                                                        .Energieverbrauchsanteil_thermisch_erzeugte_Kaelte.Wert = ImportInt(Wert)
                                                    Case "Energieverbrauchsanteil-Heizung"
                                                        .Energieverbrauchsanteil_Heizung.Wert = ImportInt(Wert)
                                                    Case "Klimafaktor", "Witterungskorrekturfaktor"
                                                        .Klimafaktor.Wert = ImportDec(Wert)
                                                End Select
                                                '-----------------------------------------------------
                                            End With
                                            '-----------------------------------------------------
                                        End If
                                        '-----------------------------------------------------
                                End Select
                                '-----------------------------------------------------
                            End With
                            '-----------------------------------------------------
                        End If
                        '--------------------------------------------------
                End Select
                '--------------------------------------------------
            End With
            '--------------------------------------------------
        Catch ex As Exception
            Fehlerfenster(ex)
        End Try
        '--------------------------------------------------
    End Sub
    '--------------------------------------------------
#End Region
    '--------------------------------------------------
#Region "XML Kontrolldatei bearbeiten und speichern"
    '--------------------------------------------------
    ''' <summary>
    ''' In dieser Anweisung wird die vorhandene Kontrolldatei gemäß den vorgaben des DIBt bearbeitet.
    ''' </summary>
    Sub XML_Kontrolldatei_bearbeiten_und_speichern()
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            Dim myFile As New FileInfo(Variable_Steuerung.Importdatei)
            Dim Extension As String = myFile.Extension
            Dim Name_ohne_Extension As String = myFile.FullName.Remove(myFile.FullName.Length - Extension.Length, Extension.Length)
            Variable_Steuerung.Kontrolldatei = Name_ohne_Extension & "_kontrolldatei_" & Variable_Steuerung.Registriernummer & Extension
            '--------------------------------------------------
            Dim Doc As New XmlDocument
            Doc.Load(Variable_Steuerung.Importdatei)
            '--------------------------------------------------
            Dim XML_Inhalt As String = Doc.InnerXml
            '--------------------------------------------------
            Doc.InnerXml = XML_Replace(Doc.InnerXml)
            '--------------------------------------------------
            If Variable_Steuerung.Datenregistratur_Fehler_Anzahl = 0 Then
                Doc.InnerXml = XML_Kontrolldatei_Registriernummer_eintragen(Doc.InnerXml, Variable_Steuerung.Registriernummer)
            End If
            '--------------------------------------------------
            Doc.InnerXml = XML_Kontrolldatei_Knoten_loeschen(Doc.InnerXml, "Gebaeudebezogene-Daten")
            '--------------------------------------------------
            Doc.InnerXml = XML_Kontrolldatei_Knoten_loeschen(Doc.InnerXml, "Signature")
            '--------------------------------------------------
            Doc.Save(Variable_Steuerung.Kontrolldatei)
            '--------------------------------------------------
        Catch ex As Exception
            Fehlerfenster(ex)
        End Try
        '--------------------------------------------------
    End Sub
    '--------------------------------------------------
    ''' <summary>
    ''' In dieser Anweisung werden die Besonderheiten aus alten xsd Schema Dateien ausgetauscht.
    ''' </summary>
    Function XML_Replace(ByVal XML_Datei As String) As String
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            XML_Replace = XML_Datei
            '--------------------------------------------------
            If Variable_Steuerung.XML_Kontrolldatei_Prefix <> "" Then
                '--------------------------------------------------
                XML_Replace = Replace(XML_Replace, "xmlns:" & Variable_Steuerung.XML_Kontrolldatei_Prefix & "=", "xmlns=")
                '--------------------------------------------------
                XML_Replace = Replace(XML_Replace, "<" & Variable_Steuerung.XML_Kontrolldatei_Prefix & ":", "<")
                '--------------------------------------------------
                XML_Replace = Replace(XML_Replace, "</" & Variable_Steuerung.XML_Kontrolldatei_Prefix & ":", "</")
                '--------------------------------------------------
            End If
            '--------------------------------------------------
        Catch ex As Exception
            Fehlerfenster(ex)
        End Try
        '--------------------------------------------------
    End Function
    '--------------------------------------------------
    ''' <summary>
    ''' In dieser Anweisung wird die Registriernummer in die vorhandene Kontrolldatei eingetragen.
    ''' </summary>
    Function XML_Kontrolldatei_Registriernummer_eintragen(ByVal XML_Datei As String, ByVal Registriernummer As String) As String
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            XML_Kontrolldatei_Registriernummer_eintragen = XML_Datei
                '--------------------------------------------------
                Dim Zeichenanzahl_bis_Eintrag As Integer
                Dim Zeichenanzahl_ab_Eintrag As Integer
                Dim Zeichenanzahl_Gesamt As Integer
                Dim Zeichenfolge_bis_Eintrag As String
                Dim Zeichenfolge_ab_Eintrag As String
                '--------------------------------------------------
                Zeichenanzahl_bis_Eintrag = XML_Kontrolldatei_Registriernummer_eintragen.IndexOf("<Energieausweis-Daten")
                Zeichenanzahl_bis_Eintrag = XML_Kontrolldatei_Registriernummer_eintragen.IndexOf(">", Zeichenanzahl_bis_Eintrag)
                Zeichenanzahl_ab_Eintrag = XML_Kontrolldatei_Registriernummer_eintragen.IndexOf("<Ausstellungsdatum>")
                Zeichenanzahl_Gesamt = XML_Kontrolldatei_Registriernummer_eintragen.Length
                '--------------------------------------------------
                Zeichenfolge_bis_Eintrag = String_Left(XML_Kontrolldatei_Registriernummer_eintragen, Zeichenanzahl_bis_Eintrag + 1)
                Zeichenfolge_ab_Eintrag = String_Right(XML_Kontrolldatei_Registriernummer_eintragen, Zeichenanzahl_Gesamt - Zeichenanzahl_ab_Eintrag)
                '--------------------------------------------------
                XML_Kontrolldatei_Registriernummer_eintragen = Zeichenfolge_bis_Eintrag & "<Registriernummer>" & Registriernummer & "</Registriernummer>" & Zeichenfolge_ab_Eintrag
            '--------------------------------------------------
        Catch ex As Exception
            Fehlerfenster(ex)
        End Try
        '--------------------------------------------------
    End Function
    '--------------------------------------------------
    ''' <summary>
    ''' In dieser Anweisung werden aus der Kontrolldatei Knoten gelöscht.
    ''' </summary>
    Function XML_Kontrolldatei_Knoten_loeschen(ByVal XML_Datei As String, ByVal Knoten As String) As String
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            XML_Kontrolldatei_Knoten_loeschen = XML_Datei
            '--------------------------------------------------
            Dim Zeichenanzahl_bis_Eintrag As Integer = XML_Kontrolldatei_Knoten_loeschen.IndexOf("<" & Knoten)
            Dim Zeichenanzahl_ab_Eintrag As Integer = XML_Kontrolldatei_Knoten_loeschen.IndexOf("</" & Knoten & ">")
            Dim Zeichenanzahl_Gesamt As Integer = XML_Kontrolldatei_Knoten_loeschen.Length
            Dim Zeichenanzahl_Knoten As Integer = Knoten.Length
            '--------------------------------------------------
            Dim Zeichenfolge_bis_Eintrag As String
            Dim Zeichenfolge_ab_Eintrag As String
            '--------------------------------------------------
            If Zeichenanzahl_bis_Eintrag = -1 Or Zeichenanzahl_ab_Eintrag = -1 Then
                '--------------------------------------------------
                XML_Kontrolldatei_Knoten_loeschen = XML_Datei
                '--------------------------------------------------
            Else
                '--------------------------------------------------
                Zeichenfolge_bis_Eintrag = String_Left(XML_Kontrolldatei_Knoten_loeschen, Zeichenanzahl_bis_Eintrag)
                '--------------------------------------------------
                Zeichenfolge_ab_Eintrag = String_Right(XML_Kontrolldatei_Knoten_loeschen, Zeichenanzahl_Gesamt - (Zeichenanzahl_ab_Eintrag + Zeichenanzahl_Knoten + 3))
                '--------------------------------------------------
                XML_Kontrolldatei_Knoten_loeschen = Zeichenfolge_bis_Eintrag & Zeichenfolge_ab_Eintrag
                '--------------------------------------------------
            End If
            '--------------------------------------------------
        Catch ex As Exception
            Fehlerfenster(ex)
        End Try
        '--------------------------------------------------
    End Function
    '--------------------------------------------------
#End Region
    '--------------------------------------------------
#Region "XML_Datei_Datenregistratur_auswerten"
    '--------------------------------------------------
    ''' <summary>
    ''' In dieser Anweisung wird das Ergebnis aus der Anforderung (Datenregistratur) beim Webservice ausgelesen.
    ''' </summary>
    Sub XML_Datei_Datenregistratur_auswerten(ByVal Dateiinhalt As String)
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            If Dateiinhalt <> "" Then
                '--------------------------------------------------
                Dim XMLReader As XmlReader = XmlReader.Create(New StringReader(Dateiinhalt))
                Dim Bezeichnung As String = ""
                Dim Wert As String = ""
                Dim Anzahl_ID As Integer = 0
                Dim Anzahl_Text As Integer = 0
                Variable_Steuerung.Datenregistratur_Fehler_Anzahl = 0
                '--------------------------------------------------
                With XMLReader
                    '--------------------------------------------------
                    Do While .Read
                        '--------------------------------------------------
                        Select Case .NodeType
                        '--------------------------------------------------
                            Case Xml.XmlNodeType.Element
                                '--------------------------------------------------
                                If .Name <> "" Then
                                    Bezeichnung = .Name
                                End If
                            '--------------------------------------------------
                            Case Xml.XmlNodeType.Text
                                '--------------------------------------------------
                                If .Value <> "" Then
                                    '--------------------------------------------------
                                    Select Case Bezeichnung
                                        Case "Registriernummer"
                                            Variable_Steuerung.Registriernummer = .Value
                                        Case "Restkontingent"
                                            Variable_Steuerung.Restkontingent_Anzahl = CInt(.Value)
                                        Case "Datendatei"
                                            Variable_Steuerung.Datenregistratur_Pruefung_Datendatei = CInt(.Value)
                                        Case "Bemerkungen"
                                            Variable_Steuerung.Datenregistratur_Pruefung_Datendatei_Bemerkungen = .Value
                                        Case "id"
                                            Variable_Steuerung.Datenregistratur_Fehler_ID(Anzahl_ID) = CInt(.Value)
                                            Anzahl_ID += 1
                                            If CInt(.Value) <> 0 Then Variable_Steuerung.Datenregistratur_Fehler_Anzahl += 1
                                        Case "value"
                                            Variable_Steuerung.Datenregistratur_Fehler_Text(Anzahl_Text) = .Value
                                            Anzahl_Text += 1
                                    End Select
                                    '--------------------------------------------------
                                End If
                                '--------------------------------------------------
                        End Select
                        '--------------------------------------------------
                    Loop
                    '--------------------------------------------------
                    .Close()
                    '--------------------------------------------------
                End With
                '--------------------------------------------------
            End If
            '--------------------------------------------------
        Catch ex As Exception
            Fehlerfenster(ex)
        End Try
        '--------------------------------------------------
    End Sub
    '--------------------------------------------------
#End Region
    '--------------------------------------------------
#Region "XML_Datei_ZusatzdatenErfassung_auswerten"
    '--------------------------------------------------
    ''' <summary>
    ''' In dieser Anweisung wird das Ergebnis aus der Anforderung (ZusatzdatenErfassung) beim Webservice ausgelesen.
    ''' </summary>
    Sub XML_Datei_ZusatzdatenErfassung_auswerten(ByVal Dateiinhalt As String)
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            If Dateiinhalt <> "" Then
                '--------------------------------------------------
                Dim XMLReader As XmlReader = XmlReader.Create(New StringReader(Dateiinhalt))
                Dim Bezeichnung As String = ""
                Dim Wert As String = ""
                Dim Anzahl_ID As Integer = 0
                Dim Anzahl_Text As Integer = 0
                Variable_Steuerung.ZusatzdatenErfassung_Fehler_Anzahl = 0
                '--------------------------------------------------
                With XMLReader
                    '--------------------------------------------------
                    Do While .Read
                        '--------------------------------------------------
                        Select Case .NodeType
                        '--------------------------------------------------
                            Case Xml.XmlNodeType.Element
                                '--------------------------------------------------
                                If .Name <> "" Then
                                    Bezeichnung = .Name
                                End If
                            '--------------------------------------------------
                            Case Xml.XmlNodeType.Text
                                '--------------------------------------------------
                                If .Value <> "" Then
                                    '--------------------------------------------------
                                    Select Case Bezeichnung
                                        Case "id"
                                            Variable_Steuerung.ZusatzdatenErfassung_Fehler_ID(Anzahl_ID) = CInt(.Value)
                                            Anzahl_ID += 1
                                            If CInt(.Value) <> 0 Then Variable_Steuerung.ZusatzdatenErfassung_Fehler_Anzahl += 1
                                        Case "value"
                                            Variable_Steuerung.ZusatzdatenErfassung_Fehler_Text(Anzahl_Text) = .Value
                                            Anzahl_Text += 1
                                    End Select
                                    '--------------------------------------------------
                                End If
                                '--------------------------------------------------
                        End Select
                        '--------------------------------------------------
                    Loop
                    '--------------------------------------------------
                    .Close()
                    '--------------------------------------------------
                End With
                '--------------------------------------------------
            End If
            '--------------------------------------------------
        Catch ex As Exception
            Fehlerfenster(ex)
        End Try
        '--------------------------------------------------
    End Sub
    '--------------------------------------------------
#End Region
    '--------------------------------------------------
#Region "XML_Import_Funktionen"
    '--------------------------------------------------
    ''' <summary>
    ''' Dieser Anweisung dient als Hilfsanweisung für den XML Import. 
    ''' </summary>
    Function ImportDec(ByVal Wert As String) As Decimal

        Try
            '--------------------------------------------------
            Dim oAssembly As System.Reflection.AssemblyName = System.Reflection.Assembly.GetExecutingAssembly().GetName
            Dim DecimalSeparator As String = oAssembly.CultureInfo.NumberFormat.CurrencyDecimalSeparator
            '--------------------------------------------------
            If DecimalSeparator = "." Then
                ImportDec = Replace(Wert, ".", ",")
            End If
            '--------------------------------------------------
            If DecimalSeparator = "," Then
                ImportDec = Replace(Wert, ",", ".")
            End If
            '--------------------------------------------------
            ImportDec = CDec(ImportDec)
            '--------------------------------------------------
        Catch ex As Exception
            Return 0
        End Try

    End Function
    '--------------------------------------------------
    ''' <summary>
    ''' Dieser Anweisung dient als Hilfsanweisung für den XML Import. 
    ''' </summary>
    Function ImportInt(ByVal Wert As String) As Integer

        Try
            '--------------------------------------------------
            ImportInt = CInt(Wert)
            '--------------------------------------------------
        Catch ex As Exception
            Return 0
        End Try

    End Function
    '--------------------------------------------------
    ''' <summary>
    ''' Dieser Anweisung dient als Hilfsanweisung für den XML Import. 
    ''' </summary>
    Function ImportBool(ByVal Wert As String) As Boolean

        Try
            '--------------------------------------------------
            If Wert = "true" Then Wert = "True"
            If Wert = "false" Then Wert = "False"
            '--------------------------------------------------
            ImportBool = CBool(Wert)
            '--------------------------------------------------
        Catch ex As Exception
            Return "False"
        End Try

    End Function
    '--------------------------------------------------
    ''' <summary>
    ''' Dieser Anweisung dient als Hilfsanweisung für den XML Import. 
    ''' </summary>
    Function ImportDate(ByVal Wert As String) As Date

        Try
            '--------------------------------------------------
            ImportDate = CDate(Wert)
            '--------------------------------------------------
        Catch ex As Exception
            Return ""
        End Try

    End Function
    '--------------------------------------------------
    ''' <summary>
    ''' Dieser Anweisung dient als Hilfsanweisung für den XML Import. 
    ''' </summary>
    Function ImportHauptnutzung(ByVal Wert As String) As String
        '--------------------------------------------------
        Dim Lage_Zeichen As Integer
        Dim Zeichenanzahl_Gesamt As Integer
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            Lage_Zeichen = Wert.IndexOf(":")
            Zeichenanzahl_Gesamt = Wert.Length
            ImportHauptnutzung = String_Right(Wert, Zeichenanzahl_Gesamt - Lage_Zeichen - 1)
            '--------------------------------------------------
        Catch ex As Exception
            Return ""
        End Try

    End Function
    '--------------------------------------------------
    ''' <summary>
    ''' Dieser Anweisung dient als Hilfsanweisung für den XML Import. 
    ''' </summary>
    Function Import_XML_Zusatz(ByVal Wert As String, ByVal Prefix As String) As String
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            Import_XML_Zusatz = Wert
            '--------------------------------------------------
            If Prefix = String.Empty Then

            Else
                Import_XML_Zusatz = Replace(Wert, Prefix & ":", "")
            End If
            '--------------------------------------------------
        Catch ex As Exception
            Fehlerfenster(ex)
        End Try
        '--------------------------------------------------
    End Function
    '--------------------------------------------------
    ''' <summary>
    ''' Dieser Anweisung dient als Hilfsanweisung für den XML Import. 
    ''' </summary>
    Function String_Left(ByVal sText As String, ByVal nLen As Integer) As String
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            If nLen > sText.Length Then nLen = sText.Length
            Return (sText.Substring(0, nLen))
            '--------------------------------------------------
        Catch ex As Exception
            Fehlerfenster(ex)
        End Try
        '--------------------------------------------------
    End Function
    '--------------------------------------------------
    ''' <summary>
    ''' Dieser Anweisung dient als Hilfsanweisung für den XML Import. 
    ''' </summary>
    Function String_Right(ByVal sText As String, ByVal nLen As Integer) As String
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            If nLen > sText.Length Then nLen = sText.Length
            Return (sText.Substring(sText.Length - nLen))
            '--------------------------------------------------
        Catch ex As Exception
            Fehlerfenster(ex)
        End Try
        '--------------------------------------------------
    End Function
    '--------------------------------------------------
#End Region
    '--------------------------------------------------
#Region "XML Datei schreiben"
    '----------------------------------------------------------------------------------
    ''' <summary>
    ''' Dieser Anweisung dient als Hilfsanweisung für den XML Import. 
    ''' </summary>
    Sub XML_Schreiben(ByVal XML As XmlWriter, ByVal name As String, ByVal Wert As String)
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            XML.WriteStartElement(name)
            XML.WriteString(ReplaceX(Wert, True))
            XML.WriteEndElement()
            '--------------------------------------------------
        Catch ex As Exception
            Fehlerfenster(ex)
        End Try
        '--------------------------------------------------
    End Sub
    '----------------------------------------------------------------------------------
    ''' <summary>
    ''' Dieser Anweisung dient als Hilfsanweisung für den XML Import. 
    ''' </summary>
    Sub XML_Schreiben(ByVal XML As XmlWriter, ByVal name As String, ByVal Wert As String, ByVal Attribut1_name As String, ByVal Attribut1_Wert As String)
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            XML.WriteStartElement(name)
            XML.WriteAttributeString(Attribut1_name, Attribut1_Wert)
            XML.WriteString(ReplaceX(Wert, True))
            XML.WriteEndElement()
            '--------------------------------------------------
        Catch ex As Exception
            Fehlerfenster(ex)
        End Try
        '--------------------------------------------------
    End Sub
    '----------------------------------------------------------------------------------
    ''' <summary>
    ''' Dieser Anweisung dient als Hilfsanweisung für den XML Import. 
    ''' </summary>
    Sub XML_Schreiben(ByVal XML As XmlWriter, ByVal name As String, ByVal Wert As String, ByVal Attribut1_name As String, ByVal Attribut1_Wert As String, ByVal Attribut2_name As String, ByVal Attribut2_Wert As String)
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            XML.WriteStartElement(name)
            XML.WriteAttributeString(Attribut1_name, Attribut1_Wert)
            XML.WriteAttributeString(Attribut2_name, Attribut2_Wert)
            XML.WriteString(ReplaceX(Wert, True))
            XML.WriteEndElement()
            '--------------------------------------------------
        Catch ex As Exception
            Fehlerfenster(ex)
        End Try
        '--------------------------------------------------
    End Sub
    '----------------------------------------------------------------------------------
    ''' <summary>
    ''' Dieser Anweisung dient als Hilfsanweisung für den XML Import. 
    ''' </summary>
    Public Function ReplaceX(ByVal Wert As String, ByVal Richtung As Boolean)
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            Select Case Richtung
                Case True 'Schreiben
                    Wert = Replace(Wert, "<", "¬")
                    Wert = Replace(Wert, ">", "¯")
                Case False 'Lesen
                    Wert = Replace(Wert, "¬", "<")
                    Wert = Replace(Wert, "¯", ">")
            End Select

            ReplaceX = Wert
            '--------------------------------------------------
        Catch ex As Exception
            Fehlerfenster(ex)
        End Try
        '--------------------------------------------------
    End Function
    '----------------------------------------------------------------------------------
#End Region
    '--------------------------------------------------
End Module
