'--------------------------------------------------
#Region "Imports"
Imports System.IO
Imports System.Windows.Forms
#End Region
'--------------------------------------------------
Module Modul_Ablauf_Funktionen
    '--------------------------------------------------
#Region "Funktionen Ablauf"
    '--------------------------------------------------
    ''' <summary>
    ''' In dieser Anweisung werden die Argumente die beim Start gesetzt wurden ausgelesen. Wenn keine oder Fehlerhafte Argumente gesetzt wurden, wird eine Fehlermeldung ausgegeben.
    ''' </summary>
    Sub Startparameter_auslesen()
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            Dim Param As String
            Dim Nummer As Integer = 0
            Dim Anzahl As Integer = 0
            '--------------------------------------------------
            For Each Param In Environment.GetCommandLineArgs
                '--------------------------------------------------
                Select Case Nummer
                    Case 0
                        '--------------------------------------------------
                        Dim myFile As New FileInfo(Param)
                        Variable_Parameter.Arbeitsverzeichnis = myFile.DirectoryName
                        '--------------------------------------------------
                        Anzahl = 1
                        '--------------------------------------------------
                    Case 1
                        '--------------------------------------------------
                        Variable_Parameter.Steuerungsdatei = Param
                        '--------------------------------------------------
                        Anzahl = 2
                        '--------------------------------------------------
                    Case 2
                        '--------------------------------------------------
                        Variable_Parameter.Aussteller_ID_DIBT = Param
                        '--------------------------------------------------
                        Anzahl = 3
                        '--------------------------------------------------
                    Case 3
                        '--------------------------------------------------
                        Variable_Parameter.Aussteller_PWD_DIBT = Param
                        '--------------------------------------------------
                        Anzahl = 4
                        '--------------------------------------------------
                End Select
                '--------------------------------------------------
                Nummer += 1
                '--------------------------------------------------
            Next
            '--------------------------------------------------
            If Anzahl <> 4 Then
                MsgBox("Die Übergabeparameter sind fehlerhaft.")
                Variable_Steuerung.Ablauf_Fehler = True
            End If
            '--------------------------------------------------
        Catch ex As Exception
            Fehlerfenster(ex)
        End Try
        '--------------------------------------------------
    End Sub
    '--------------------------------------------------
    ''' <summary>
    ''' In dieser Anweisung wird das Auslesen der Steuerungsdatei gestartet, wenn keine Steuerungsdatei gefunden wird, wird eine Fehlermeldung ausgegeben.
    ''' </summary>
    Sub Steuerungsdatei_lesen()
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            If Variable_Steuerung.Ablauf_Fehler = False Then
                '--------------------------------------------------
                If Not File.Exists(Variable_Parameter.Steuerungsdatei) Then
                    '--------------------------------------------------
                    MsgBox("Steuerungsdatei nicht vorhanden!")
                    '--------------------------------------------------
                    Variable_Steuerung.Ablauf_Fehler = True
                    '--------------------------------------------------
                Else
                    '--------------------------------------------------
                    Importvariablen_Steuerungsdatei_zuruecksetzen()
                    '--------------------------------------------------
                    XML_Steuerungsdatei_einlesen(Variable_Parameter.Steuerungsdatei)
                    '--------------------------------------------------
                End If
                '--------------------------------------------------
            Else
                '--------------------------------------------------
                'Funktion wurde nicht gestartet, weil es einen Fehler gab.
                '--------------------------------------------------
            End If
            '--------------------------------------------------
        Catch ex As Exception
            Fehlerfenster(ex)
        End Try
        '--------------------------------------------------
    End Sub
    '--------------------------------------------------
    ''' <summary>
    ''' In dieser Anweisung werden die Variablen der Steuerungsdatei zurückgesetzt.
    ''' </summary>
    Sub Importvariablen_Steuerungsdatei_zuruecksetzen()
        '--------------------------------------------------
        Dim Anzahl_Fehlermeldung As Integer = 200
        Dim Anzahl_Ausweise As Integer = 1000
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            With Variable_Steuerung
                '--------------------------------------------------
                ReDim .Datenregistratur_Fehler_ID(Anzahl_Fehlermeldung)
                ReDim .Datenregistratur_Fehler_Text(Anzahl_Fehlermeldung)
                '--------------------------------------------------
                ReDim .KontrolldateiPruefen_Fehler_ID(Anzahl_Fehlermeldung)
                ReDim .KontrolldateiPruefen_Fehler_Kurztext(Anzahl_Fehlermeldung)
                ReDim .KontrolldateiPruefen_Fehler_Langtext(Anzahl_Fehlermeldung)
                '--------------------------------------------------
                ReDim .ZusatzdatenErfassung_Fehler_ID(Anzahl_Fehlermeldung)
                ReDim .ZusatzdatenErfassung_Fehler_Text(Anzahl_Fehlermeldung)
                '--------------------------------------------------
                ReDim .OffeneKontrolldateien_Ausweis_Registriernummer(Anzahl_Ausweise)
                ReDim .OffeneKontrolldateien_Ausweis_NummerErzeugtAm(Anzahl_Ausweise)
                ReDim .OffeneKontrolldateien_Ausweis_Aussteller(Anzahl_Ausweise)
                '--------------------------------------------------
                .Datenregistratur = False
                .KontrolldateiPruefen = False
                .Restkontingent = False
                .ZusatzdatenErfassung = False
                .OffeneKontrolldateien = False
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
    ''' In dieser Anweisung werden die Variablen der Kontrolldatei zurückgesetzt.
    ''' </summary>
    Sub Importvariablen_Kontrolldatei_zuruecksetzen()
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            With Variable_XML_Import
                '--------------------------------------------------
                ReDim .Energieausweis_Daten.Nichtwohngebaeude.Zone(200)
                '--------------------------------------------------
                ReDim .Energietraeger(200)
                '--------------------------------------------------
                ReDim .Energietraeger_Daten(8)
                '--------------------------------------------------
                Dim WertXX As Integer
                '--------------------------------------------------
                For WertXX = 1 To 8
                    ReDim .Energietraeger_Daten(WertXX).Zeitraum_Daten(40)
                Next
                ReDim .Zeitraum_Strom_Daten(40)
                '--------------------------------------------------
                ReDim .Nutzung_Gebaeudekategorie(5)
                '--------------------------------------------------
                ReDim .Verbrauchserfassung(500)
                '--------------------------------------------------
                ReDim .Modernisierungsempfehlungen(30)
                '--------------------------------------------------
                ReDim .EE_65EE_Regel(200)
                ReDim .EE_65EE_keine_Regel(200)
                '--------------------------------------------------
                ReDim .BauteilOpak(10000)
                ReDim .BauteilTransparent(10000)
                ReDim .BauteilDach(10000)
                '--------------------------------------------------
                ReDim .Heizungsanlage(200)
                ReDim .Trinkwarmwasseranlage(200)
                ReDim .Kaelteanlage(200)
                ReDim .RLT_System(200)
                '--------------------------------------------------
                ReDim .Bilddateien(40)
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
    ''' In dieser Anweisung wird das Auslesen der Importdatei gestartet, wenn keine Importdatei gefunden wird, wird eine Fehlermeldung ausgegeben.
    ''' </summary>
    Sub Kontrolldatei_lesen()
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            If Variable_Steuerung.Ablauf_Fehler = False Then
                '--------------------------------------------------
                If Not File.Exists(Variable_Steuerung.Importdatei) Then
                    '--------------------------------------------------
                    MsgBox("Kontrolldatei ist nicht vorhanden!")
                    '--------------------------------------------------
                    Variable_Steuerung.Ablauf_Fehler = True
                    '--------------------------------------------------
                Else
                    '--------------------------------------------------
                    Importvariablen_Kontrolldatei_zuruecksetzen()
                    '--------------------------------------------------
                    XML_Kontrolldatei_einlesen(Variable_Steuerung.Importdatei)
                    '--------------------------------------------------
                    Testdaten_anlegen()
                    '--------------------------------------------------
                End If
                '--------------------------------------------------
            Else
                '--------------------------------------------------
                'Funktion wurde nicht gestartet, weil es einen Fehler gab.
                '--------------------------------------------------
            End If
            '--------------------------------------------------
        Catch ex As Exception
            Fehlerfenster(ex)
        End Try
        '--------------------------------------------------
    End Sub
    '--------------------------------------------------
    ''' <summary>
    ''' In dieser Anweisung wird das Schreiben der Steuerungsdatei gestartet.
    ''' </summary>
    Sub Steuerungsdatei_erstellen()
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            If Not File.Exists(Variable_Parameter.Steuerungsdatei) Then
                '--------------------------------------------------
                Kill(Variable_Parameter.Steuerungsdatei)
                '--------------------------------------------------
                XML_Steuerungsdatei_schreiben(Variable_Parameter.Steuerungsdatei)
                '--------------------------------------------------
            Else
                '--------------------------------------------------
                XML_Steuerungsdatei_schreiben(Variable_Parameter.Steuerungsdatei)
                '--------------------------------------------------
            End If
            '--------------------------------------------------
            If Variable_Steuerung.Ablauf_Fehler = False Then
                Environment.ExitCode = 0
            Else
                Environment.ExitCode = 1
            End If
            '--------------------------------------------------
        Catch ex As Exception
            Fehlerfenster(ex)
        End Try
        '--------------------------------------------------
    End Sub
    '--------------------------------------------------
#End Region
    '-------------------------------------------------------------------
#Region "String in Base64 umwandeln"
    '-------------------------------------------------------------------
    ''' <summary>
    ''' String in Base64 umwandeln
    ''' </summary>
    Function String_To_Base64(ByVal sText As String) As String
        Dim nBytes() As Byte = System.Text.Encoding.Default.GetBytes(sText)
        Return System.Convert.ToBase64String(nBytes)
    End Function
    '-------------------------------------------------------------------
    ''' <summary>
    ''' Base64 in String umwandeln
    ''' </summary>
    Function Base64_to_String(ByVal sText As String) As String
        Dim nBytes() As Byte = System.Convert.FromBase64String(sText)
        Return System.Text.Encoding.Default.GetString(nBytes)
    End Function
    '-------------------------------------------------------------------
#End Region
    '--------------------------------------------------
#Region "Testdaten anlegen"
    '--------------------------------------------------
    ''' <summary>
    ''' In dieser Anweisung werden Testdaten geschrieben. Diese Funktion wird nur für einen internen Test benötigt.
    ''' </summary>
    Sub Testdaten_anlegen()
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            Dim Wert_X As Integer
            '--------------------------------------------------
            With Variable_XML_Import
                '--------------------------------------------------
                If Variable_Steuerung.Testdaten = True Then
                    '--------------------------------------------------
                    '--------------------------------------------------
                    For Wert_X = 1 To 200
                        '--------------------------------------------------
                        With .Energieausweis_Daten.Nichtwohngebaeude.Zone(Wert_X)
                            '--------------------------------------------------
                            .Zonenbezeichnung.Wert = "Zone " & Wert_X
                            .Nettogrundflaeche_Zone.Wert = Wert_X * 10
                            '--------------------------------------------------
                        End With
                        '--------------------------------------------------
                    Next
                    '--------------------------------------------------
                    '--------------------------------------------------
                    .Energieausweis_Daten.Empfehlungen_moeglich.Wert = True
                    For Wert_X = 1 To 30
                        '--------------------------------------------------
                        With .Modernisierungsempfehlungen(Wert_X)
                            '--------------------------------------------------
                            .Nummer.Wert = CInt(Wert_X)
                            .Bauteil_Anlagenteil.Wert = "Bauteil " & Wert_X.ToString
                            .Massnahmenbeschreibung.Wert = "Maßnahme " & Wert_X.ToString
                            .Modernisierungskombination.Wert = "als Einzelmaßnahme"
                            .Amortisation.Wert = (Wert_X * 1.45).ToString
                            .spezifische_Kosten.Wert = (Wert_X * 2.88).ToString
                            '--------------------------------------------------
                        End With
                        '--------------------------------------------------
                    Next
                    '--------------------------------------------------
                    '--------------------------------------------------
                    For Wert_X = 1 To 200
                        '--------------------------------------------------
                        With .Energietraeger(Wert_X)
                            '--------------------------------------------------
                            .Energietraegerbezeichnung.Wert = "Energieträger " & Wert_X
                            .Endenergiebedarf_Heizung_spezifisch.Wert = Wert_X * 1
                            .Endenergiebedarf_Trinkwarmwasser_spezifisch.Wert = Wert_X * 2
                            .Endenergiebedarf_Beleuchtung_spezifisch.Wert = Wert_X * 3
                            .Endenergiebedarf_Lueftung_spezifisch.Wert = Wert_X * 4
                            .Endenergiebedarf_Kuehlung_Befeuchtung_spezifisch.Wert = Wert_X * 5
                            .Endenergiebedarf_Energietraeger_Gesamtgebaeude_spezifisch.Wert = Wert_X * 15
                            '--------------------------------------------------
                        End With
                        '--------------------------------------------------
                    Next
                    '--------------------------------------------------
                    '--------------------------------------------------
                    For Wert_X = 1 To 8
                        '--------------------------------------------------
                        Dim WertXX As Integer
                        '--------------------------------------------------
                        For WertXX = 1 To 40
                            '--------------------------------------------------
                            With .Energietraeger_Daten(Wert_X)
                                '--------------------------------------------------
                                .Zeitraum_Daten(WertXX).Startdatum.Wert = "01.01.2022"
                                .Zeitraum_Daten(WertXX).Enddatum.Wert = "01.01.2022"
                                .Energietraeger_Verbrauch.Wert = "Energieträger " & Wert_X
                                .Zeitraum_Daten(WertXX).Energieverbrauch.Wert = CInt(Int((6 * Rnd()) * 11))
                                .Primaerenergiefaktor.Wert = CInt(Int((6 * Rnd()) + 1))
                                .Zeitraum_Daten(WertXX).Energieverbrauch.Wert = CInt(Int((6 * Rnd()) + 1))
                                .Zeitraum_Daten(WertXX).Energieverbrauchsanteil_Warmwasser_zentral.Wert = CInt(Int((6 * Rnd()) + 1))
                                .Zeitraum_Daten(WertXX).Energieverbrauchsanteil_thermisch_erzeugte_Kaelte.Wert = CInt(Int((6 * Rnd()) + 1))
                                .Zeitraum_Daten(WertXX).Energieverbrauchsanteil_Heizung.Wert = CInt(Int((6 * Rnd()) + 1))
                                .Zeitraum_Daten(WertXX).Klimafaktor.Wert = Rnd()
                                '--------------------------------------------------
                            End With
                            '--------------------------------------------------
                        Next
                        '--------------------------------------------------
                    Next
                    '--------------------------------------------------
                    '--------------------------------------------------
                    For Wert_X = 1 To 40
                        '--------------------------------------------------
                        With .Zeitraum_Strom_Daten(Wert_X)
                            '--------------------------------------------------
                            .Startdatum.Wert = "01.01.2022"
                            .Enddatum.Wert = "01.01.2022"
                            .Energieverbrauch_Strom.Wert = CInt(Int((6 * Rnd()) + 1))
                            '--------------------------------------------------
                        End With
                        '--------------------------------------------------
                    Next
                    '--------------------------------------------------
                    '--------------------------------------------------
                    For Wert_X = 1 To 200
                        '--------------------------------------------------
                        With .EE_65EE_Regel(Wert_X)
                            '--------------------------------------------------
                            .Waermeerzeuger_Bauweise_18599.Wert = "Wärmeerzeuger " & Wert_X
                            .Art_der_Nutzung_erneuerbaren_Energie.Wert = "PV-Anlage"
                            .Deckungsanteil.Wert = CInt(Int((6 * Rnd()) + 1))
                            .Anteil_der_Pflichterfuellung_Anlage.Wert = CInt(Int((4 * Rnd()) + 1))
                            .Anteil_der_Pflichterfuellung_Gesamt.Wert = CInt(Int((6 * Rnd()) + 1))
                            '--------------------------------------------------
                        End With
                        '--------------------------------------------------
                    Next
                    '--------------------------------------------------
                    '--------------------------------------------------
                    For Wert_X = 1 To 200
                        '--------------------------------------------------
                        With .EE_65EE_keine_Regel(Wert_X)
                            '--------------------------------------------------
                            .Waermeerzeuger_Bauweise_18599.Wert = "Wärmeerzeuger " & Wert_X
                            .Art_der_Nutzung_erneuerbaren_Energie.Wert = "PV-Anlage"
                            .Anteil_EE_Anlage.Wert = CInt(Int((6 * Rnd()) + 1))
                            '--------------------------------------------------
                        End With
                        '--------------------------------------------------
                    Next
                    '--------------------------------------------------
                End If
                '--------------------------------------------------
            End With
        Catch ex As Exception
            Fehlerfenster(ex)
        End Try
        '--------------------------------------------------
    End Sub
    '--------------------------------------------------
#End Region
    '--------------------------------------------------
End Module
