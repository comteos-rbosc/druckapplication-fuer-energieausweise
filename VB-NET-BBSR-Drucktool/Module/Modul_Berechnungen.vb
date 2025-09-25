Module Modul_Berechnungen
    '--------------------------------------------------
#Region "Zusammenstellung der Werte für die Verbrauchserfassung"
    '--------------------------------------------------
    ''' <summary>
    ''' Zusammenstellung der Werte für die Verbrauchserfassung - Wohnbau.
    ''' </summary>
    Sub Zusammenstellung_WB_Verbrauchserfassung()
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            Dim Wert_X As Integer
            Dim Wert_Y As Integer
            '--------------------------------------------------
            With Variable_XML_Import
                '--------------------------------------------------
                Dim Maximale_Anzahl As Integer = 324
                '--------------------------------------------------
                ReDim .Verbrauchserfassung(Maximale_Anzahl)
                '--------------------------------------------------
                .Verbrauchserfassung_Anzahl = 0
                '--------------------------------------------------
                'Auswertung Energieträger
                '--------------------------------------------------
                For Wert_X = 1 To 8 'Energietraeger_Daten()
                    '--------------------------------------------------
                    If .Energietraeger_Daten(Wert_X).Energietraeger_Verbrauch.Wert <> "" Or .Energietraeger_Daten(Wert_X).Sonstiger_Energietraeger_Verbrauch.Wert <> "" Then
                        '--------------------------------------------------
                        For Wert_Y = 1 To 40 '.Zeitraum_Daten()
                            '--------------------------------------------------
                            If .Energietraeger_Daten(Wert_X).Zeitraum_Daten(Wert_Y).Energieverbrauch.Wert <> 0 Then
                                '--------------------------------------------------
                                .Verbrauchserfassung_Anzahl += 1
                                '--------------------------------------------------
                                If .Energietraeger_Daten(Wert_X).Energietraeger_Verbrauch.Wert <> "" Then
                                    .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Energietraeger_Verbrauch.Wert = .Energietraeger_Daten(Wert_X).Energietraeger_Verbrauch.Wert
                                Else
                                    .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Energietraeger_Verbrauch.Wert = .Energietraeger_Daten(Wert_X).Sonstiger_Energietraeger_Verbrauch.Wert
                                End If
                                '--------------------------------------------------
                                .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Primaerenergiefaktor.Wert = .Energietraeger_Daten(Wert_X).Primaerenergiefaktor.Wert
                                '--------------------------------------------------
                                .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Startdatum.Wert = .Energietraeger_Daten(Wert_X).Zeitraum_Daten(Wert_Y).Startdatum.Wert
                                '--------------------------------------------------
                                .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Enddatum.Wert = .Energietraeger_Daten(Wert_X).Zeitraum_Daten(Wert_Y).Enddatum.Wert
                                '--------------------------------------------------
                                .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Energieverbrauch.Wert = .Energietraeger_Daten(Wert_X).Zeitraum_Daten(Wert_Y).Energieverbrauch.Wert
                                '--------------------------------------------------
                                .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Energieverbrauchsanteil_Warmwasser_zentral.Wert = .Energietraeger_Daten(Wert_X).Zeitraum_Daten(Wert_Y).Energieverbrauchsanteil_Warmwasser_zentral.Wert
                                '--------------------------------------------------
                                .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Energieverbrauchsanteil_Heizung.Wert = .Energietraeger_Daten(Wert_X).Zeitraum_Daten(Wert_Y).Energieverbrauchsanteil_Heizung.Wert
                                '--------------------------------------------------
                                .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Klimafaktor.Wert = .Energietraeger_Daten(Wert_X).Zeitraum_Daten(Wert_Y).Klimafaktor.Wert
                                '--------------------------------------------------
                            End If
                            '--------------------------------------------------
                        Next
                        '--------------------------------------------------
                    End If
                    '--------------------------------------------------
                Next
                '--------------------------------------------------
                'Leerstandszuschlag Heizung
                '--------------------------------------------------
                If .Leerstandszuschlag_Heizung.Leerstandszuschlag_kWh.Wert <> 0 Then
                    '--------------------------------------------------
                    .Verbrauchserfassung_Anzahl += 1
                    '--------------------------------------------------
                    .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Energietraeger_Verbrauch.Wert = "Leerstandszuschlag Heizung"
                    '--------------------------------------------------
                    .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Primaerenergiefaktor.Wert = .Leerstandszuschlag_Heizung.Primaerenergiefaktor.Wert
                    '--------------------------------------------------
                    .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Startdatum.Wert = .Leerstandszuschlag_Heizung.Startdatum.Wert
                    '--------------------------------------------------
                    .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Enddatum.Wert = .Leerstandszuschlag_Heizung.Enddatum.Wert
                    '--------------------------------------------------
                    .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Energieverbrauch.Wert = .Leerstandszuschlag_Heizung.Leerstandszuschlag_kWh.Wert
                    '--------------------------------------------------
                    .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Energieverbrauchsanteil_Heizung.Wert = .Leerstandszuschlag_Heizung.Leerstandszuschlag_kWh.Wert
                    '--------------------------------------------------
                End If
                '--------------------------------------------------
                'Leerstandszuschlag Warmwasser
                '--------------------------------------------------
                If .Leerstandszuschlag_Warmwasser.Leerstandszuschlag_kWh.Wert <> 0 Then
                    '--------------------------------------------------
                    .Verbrauchserfassung_Anzahl += 1
                    '--------------------------------------------------
                    .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Energietraeger_Verbrauch.Wert = "Leerstandszuschlag Warmwasser"
                    '--------------------------------------------------
                    .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Primaerenergiefaktor.Wert = .Leerstandszuschlag_Warmwasser.Primaerenergiefaktor.Wert
                    '--------------------------------------------------
                    .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Startdatum.Wert = .Leerstandszuschlag_Warmwasser.Startdatum.Wert
                    '--------------------------------------------------
                    .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Enddatum.Wert = .Leerstandszuschlag_Warmwasser.Enddatum.Wert
                    '--------------------------------------------------
                    .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Energieverbrauch.Wert = .Leerstandszuschlag_Warmwasser.Leerstandszuschlag_kWh.Wert
                    '--------------------------------------------------
                    .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Energieverbrauchsanteil_Warmwasser_zentral.Wert = .Leerstandszuschlag_Warmwasser.Leerstandszuschlag_kWh.Wert
                    '--------------------------------------------------
                End If
                '--------------------------------------------------
                'Warmwasserzuschlag
                '--------------------------------------------------
                If .Warmwasserzuschlag.Warmwasserzuschlag_kWh.Wert <> 0 Then
                    '--------------------------------------------------
                    .Verbrauchserfassung_Anzahl += 1
                    '--------------------------------------------------
                    .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Energietraeger_Verbrauch.Wert = "Warmwasserzuschlag"
                    '--------------------------------------------------
                    .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Primaerenergiefaktor.Wert = .Warmwasserzuschlag.Primaerenergiefaktor.Wert
                    '--------------------------------------------------
                    .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Startdatum.Wert = .Warmwasserzuschlag.Startdatum.Wert
                    '--------------------------------------------------
                    .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Enddatum.Wert = .Warmwasserzuschlag.Enddatum.Wert
                    '--------------------------------------------------
                    .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Energieverbrauch.Wert = .Warmwasserzuschlag.Warmwasserzuschlag_kWh.Wert
                    '--------------------------------------------------
                    .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Energieverbrauchsanteil_Warmwasser_zentral.Wert = .Warmwasserzuschlag.Warmwasserzuschlag_kWh.Wert
                    '--------------------------------------------------
                End If
                '--------------------------------------------------
                'Kuehlzuschlag
                '--------------------------------------------------
                If .Kuehlzuschlag.Kuehlzuschlag_kWh.Wert <> 0 Then
                    '--------------------------------------------------
                    .Verbrauchserfassung_Anzahl += 1
                    '--------------------------------------------------
                    .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Energietraeger_Verbrauch.Wert = "Kühlzuschlag"
                    '--------------------------------------------------
                    .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Primaerenergiefaktor.Wert = .Kuehlzuschlag.Primaerenergiefaktor.Wert
                    '--------------------------------------------------
                    .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Startdatum.Wert = .Kuehlzuschlag.Startdatum.Wert
                    '--------------------------------------------------
                    .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Enddatum.Wert = .Kuehlzuschlag.Enddatum.Wert
                    '--------------------------------------------------
                    .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Energieverbrauch.Wert = .Kuehlzuschlag.Kuehlzuschlag_kWh.Wert
                    '--------------------------------------------------
                End If
                '--------------------------------------------------
                '--------------------------------------------------
                If .Verbrauchserfassung_Anzahl > Maximale_Anzahl Then
                    .Verbrauchserfassung_Anzahl = Maximale_Anzahl
                End If
                '--------------------------------------------------
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
    ''' Zusammenstellung der Werte für die Verbrauchserfassung - Nichtwohnbau.
    ''' </summary> 
    Sub Zusammenstellung_NWB_Verbrauchserfassung()
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            Dim Wert_X As Integer
            Dim Wert_Y As Integer
            '--------------------------------------------------
            With Variable_XML_Import
                '--------------------------------------------------
                Dim Maximale_Anzahl As Integer = 364
                '--------------------------------------------------
                ReDim .Verbrauchserfassung(Maximale_Anzahl)
                '--------------------------------------------------
                .Verbrauchserfassung_Anzahl = 0
                '--------------------------------------------------
                'Auswertung Energieträger Wärme
                '--------------------------------------------------
                For Wert_X = 1 To 8 'Energietraeger_Daten()
                    '--------------------------------------------------
                    If .Energietraeger_Daten(Wert_X).Energietraeger_Verbrauch.Wert <> "" Or .Energietraeger_Daten(Wert_X).Sonstiger_Energietraeger_Verbrauch.Wert <> "" Then
                        '--------------------------------------------------
                        For Wert_Y = 1 To 40 '.Zeitraum_Daten()
                            '--------------------------------------------------
                            If .Energietraeger_Daten(Wert_X).Zeitraum_Daten(Wert_Y).Energieverbrauch.Wert <> 0 Then
                                '--------------------------------------------------
                                .Verbrauchserfassung_Anzahl += 1
                                '--------------------------------------------------
                                If .Energietraeger_Daten(Wert_X).Energietraeger_Verbrauch.Wert <> "" Then
                                    .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Energietraeger_Verbrauch.Wert = .Energietraeger_Daten(Wert_X).Energietraeger_Verbrauch.Wert
                                Else
                                    .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Energietraeger_Verbrauch.Wert = .Energietraeger_Daten(Wert_X).Sonstiger_Energietraeger_Verbrauch.Wert
                                End If
                                '--------------------------------------------------
                                .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Primaerenergiefaktor.Wert = .Energietraeger_Daten(Wert_X).Primaerenergiefaktor.Wert
                                '--------------------------------------------------
                                .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Startdatum.Wert = .Energietraeger_Daten(Wert_X).Zeitraum_Daten(Wert_Y).Startdatum.Wert
                                '--------------------------------------------------
                                .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Enddatum.Wert = .Energietraeger_Daten(Wert_X).Zeitraum_Daten(Wert_Y).Enddatum.Wert
                                '--------------------------------------------------
                                .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Energieverbrauch.Wert = .Energietraeger_Daten(Wert_X).Zeitraum_Daten(Wert_Y).Energieverbrauch.Wert
                                '--------------------------------------------------
                                .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Energieverbrauchsanteil_Warmwasser_zentral.Wert = .Energietraeger_Daten(Wert_X).Zeitraum_Daten(Wert_Y).Energieverbrauchsanteil_Warmwasser_zentral.Wert
                                '--------------------------------------------------
                                .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Energieverbrauchsanteil_thermisch_erzeugte_Kaelte.Wert = .Energietraeger_Daten(Wert_X).Zeitraum_Daten(Wert_Y).Energieverbrauchsanteil_thermisch_erzeugte_Kaelte.Wert
                                '--------------------------------------------------
                                .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Energieverbrauchsanteil_Heizung.Wert = .Energietraeger_Daten(Wert_X).Zeitraum_Daten(Wert_Y).Energieverbrauchsanteil_Heizung.Wert
                                '--------------------------------------------------
                                .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Klimafaktor.Wert = .Energietraeger_Daten(Wert_X).Zeitraum_Daten(Wert_Y).Klimafaktor.Wert
                                '--------------------------------------------------
                            End If
                            '--------------------------------------------------
                        Next
                        '--------------------------------------------------
                    End If
                    '--------------------------------------------------
                Next
                '--------------------------------------------------
                'Auswertung Energieträger Strom
                '--------------------------------------------------
                For Wert_Y = 1 To 40 '.Zeitraum_Daten()
                    '--------------------------------------------------
                    If .Zeitraum_Strom_Daten(Wert_Y).Energieverbrauch_Strom.Wert <> 0 Then
                        '--------------------------------------------------
                        .Verbrauchserfassung_Anzahl += 1
                        '--------------------------------------------------
                        .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Energietraeger_Verbrauch.Wert = "allgemeiner Strommix"
                        '--------------------------------------------------
                        .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Primaerenergiefaktor.Wert = 1.8
                        '--------------------------------------------------
                        .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Startdatum.Wert = .Zeitraum_Strom_Daten(Wert_Y).Startdatum.Wert
                        '--------------------------------------------------
                        .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Enddatum.Wert = .Zeitraum_Strom_Daten(Wert_Y).Enddatum.Wert
                        '--------------------------------------------------
                        .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Energieverbrauchsanteil_elektrisch_erzeugte_Kaelte.Wert = .Zeitraum_Strom_Daten(Wert_Y).Energieverbrauchsanteil_elektrisch_erzeugte_Kaelte.Wert
                        '--------------------------------------------------
                        .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Energieverbrauch_Strom.Wert = .Zeitraum_Strom_Daten(Wert_Y).Energieverbrauch_Strom.Wert
                        '--------------------------------------------------
                    End If
                    '--------------------------------------------------
                Next
                '--------------------------------------------------
                'Leerstandszuschlag Heizung
                '--------------------------------------------------
                If .Leerstandszuschlag_Heizung.Leerstandszuschlag_kWh.Wert <> 0 Then
                    '--------------------------------------------------
                    .Verbrauchserfassung_Anzahl += 1
                    '--------------------------------------------------
                    .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Energietraeger_Verbrauch.Wert = "Leerstandszuschlag Heizung"
                    '--------------------------------------------------
                    .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Primaerenergiefaktor.Wert = .Leerstandszuschlag_Heizung.Primaerenergiefaktor.Wert
                    '--------------------------------------------------
                    .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Startdatum.Wert = .Leerstandszuschlag_Heizung.Startdatum.Wert
                    '--------------------------------------------------
                    .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Enddatum.Wert = .Leerstandszuschlag_Heizung.Enddatum.Wert
                    '--------------------------------------------------
                    .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Energieverbrauch.Wert = .Leerstandszuschlag_Heizung.Leerstandszuschlag_kWh.Wert
                    '--------------------------------------------------
                    .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Energieverbrauchsanteil_Heizung.Wert = .Leerstandszuschlag_Heizung.Leerstandszuschlag_kWh.Wert
                    '--------------------------------------------------
                End If
                '--------------------------------------------------
                'Leerstandszuschlag Warmwasser
                '--------------------------------------------------
                If .Leerstandszuschlag_Warmwasser.Leerstandszuschlag_kWh.Wert <> 0 Then
                    '--------------------------------------------------
                    .Verbrauchserfassung_Anzahl += 1
                    '--------------------------------------------------
                    .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Energietraeger_Verbrauch.Wert = "Leerstandszuschlag Warmwasser"
                    '--------------------------------------------------
                    .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Primaerenergiefaktor.Wert = .Leerstandszuschlag_Warmwasser.Primaerenergiefaktor.Wert
                    '--------------------------------------------------
                    .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Startdatum.Wert = .Leerstandszuschlag_Warmwasser.Startdatum.Wert
                    '--------------------------------------------------
                    .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Enddatum.Wert = .Leerstandszuschlag_Warmwasser.Enddatum.Wert
                    '--------------------------------------------------
                    .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Energieverbrauch.Wert = .Leerstandszuschlag_Warmwasser.Leerstandszuschlag_kWh.Wert
                    '--------------------------------------------------
                    .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Energieverbrauchsanteil_Warmwasser_zentral.Wert = .Leerstandszuschlag_Warmwasser.Leerstandszuschlag_kWh.Wert
                    '--------------------------------------------------
                End If
                '--------------------------------------------------
                'Leerstandszuschlag-thermisch-erzeugte-Kaelte
                '--------------------------------------------------
                If .Leerstandszuschlag_thermisch_erzeugte_Kaelte.Leerstandszuschlag_kWh.Wert <> 0 Then
                    '--------------------------------------------------
                    .Verbrauchserfassung_Anzahl += 1
                    '--------------------------------------------------
                    .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Energietraeger_Verbrauch.Wert = "Leerstandszuschlag thermisch erzeugte Kälte"
                    '--------------------------------------------------
                    .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Primaerenergiefaktor.Wert = .Leerstandszuschlag_thermisch_erzeugte_Kaelte.Primaerenergiefaktor.Wert
                    '--------------------------------------------------
                    .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Startdatum.Wert = .Leerstandszuschlag_thermisch_erzeugte_Kaelte.Startdatum.Wert
                    '--------------------------------------------------
                    .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Enddatum.Wert = .Leerstandszuschlag_thermisch_erzeugte_Kaelte.Enddatum.Wert
                    '--------------------------------------------------
                    .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Energieverbrauch.Wert = .Leerstandszuschlag_thermisch_erzeugte_Kaelte.Leerstandszuschlag_kWh.Wert
                    '--------------------------------------------------
                    .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Energieverbrauchsanteil_Warmwasser_zentral.Wert = .Leerstandszuschlag_thermisch_erzeugte_Kaelte.Leerstandszuschlag_kWh.Wert
                    '--------------------------------------------------
                End If
                '--------------------------------------------------
                'Leerstandszuschlag-Strom
                '--------------------------------------------------
                If .Leerstandszuschlag_Strom.Leerstandszuschlag_kWh.Wert <> 0 Then
                    '--------------------------------------------------
                    .Verbrauchserfassung_Anzahl += 1
                    '--------------------------------------------------
                    .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Energietraeger_Verbrauch.Wert = "Leerstandszuschlag Strom"
                    '--------------------------------------------------
                    .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Primaerenergiefaktor.Wert = .Leerstandszuschlag_Strom.Primaerenergiefaktor.Wert
                    '--------------------------------------------------
                    .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Startdatum.Wert = .Leerstandszuschlag_Strom.Startdatum.Wert
                    '--------------------------------------------------
                    .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Enddatum.Wert = .Leerstandszuschlag_Strom.Enddatum.Wert
                    '--------------------------------------------------
                    .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Energieverbrauch.Wert = .Leerstandszuschlag_Strom.Leerstandszuschlag_kWh.Wert
                    '--------------------------------------------------
                    .Verbrauchserfassung(.Verbrauchserfassung_Anzahl).Energieverbrauchsanteil_Warmwasser_zentral.Wert = .Leerstandszuschlag_Strom.Leerstandszuschlag_kWh.Wert
                    '--------------------------------------------------
                End If
                '--------------------------------------------------
                '--------------------------------------------------
                If .Verbrauchserfassung_Anzahl > Maximale_Anzahl Then
                    .Verbrauchserfassung_Anzahl = Maximale_Anzahl
                End If
                '--------------------------------------------------
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
#Region "Zusammenstellung der Werte für die Modernisierungsempfehlung"
    '--------------------------------------------------
    ''' <summary>
    ''' Zusammenstellung der Werte für die Modernisierungsempfehlung.
    ''' </summary>
    Function Anzahl_Modernisierungsempfehlung() As Integer
        '--------------------------------------------------
        Anzahl_Modernisierungsempfehlung = 0
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            Dim Wert_X As Integer
            Dim Anzahl As Integer = 0
            '--------------------------------------------------
            Dim Maximale_Anzahl As Integer = 30
            '--------------------------------------------------
            With Variable_XML_Import
                '--------------------------------------------------
                For Wert_X = 1 To Maximale_Anzahl
                    '--------------------------------------------------
                    If .Modernisierungsempfehlungen(Wert_X).Bauteil_Anlagenteil.Wert <> "" Then
                        '--------------------------------------------------
                        Anzahl += 1
                        '--------------------------------------------------
                    End If
                    '--------------------------------------------------
                Next
                '--------------------------------------------------
                Anzahl_Modernisierungsempfehlung = Anzahl
                '--------------------------------------------------
            End With
            '--------------------------------------------------
        Catch ex As Exception
            Fehlerfenster(ex)
        End Try
        '--------------------------------------------------
    End Function
    '--------------------------------------------------
#End Region
    '--------------------------------------------------
#Region "Zusammenstellung der Werte für die Gebäudenutzung"
    '--------------------------------------------------
    ''' <summary>
    ''' Zusammenstellung der Werte für die Gebäudenutzung.
    ''' </summary>
    Function Anzahl_Gebaeudenutzung() As Integer
        '--------------------------------------------------
        Anzahl_Gebaeudenutzung = 0
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            Dim Wert_X As Integer
            Dim Anzahl As Integer = 0
            '--------------------------------------------------
            Dim Maximale_Anzahl As Integer = 5
            '--------------------------------------------------
            With Variable_XML_Import
                '--------------------------------------------------
                For Wert_X = 1 To Maximale_Anzahl
                    '--------------------------------------------------
                    If .Nutzung_Gebaeudekategorie(Wert_X).Gebaeudekategorie.Wert <> "" Then
                        '--------------------------------------------------
                        Anzahl += 1
                        '--------------------------------------------------
                    End If
                    '--------------------------------------------------
                Next
                '--------------------------------------------------
                Anzahl_Gebaeudenutzung = Anzahl
                '--------------------------------------------------
            End With
            '--------------------------------------------------
        Catch ex As Exception
            Fehlerfenster(ex)
        End Try
        '--------------------------------------------------
    End Function
    '--------------------------------------------------
#End Region
    '--------------------------------------------------
#Region "Zusammenstellung der Werte für den Energiebedarf"
    '--------------------------------------------------
    ''' <summary>
    ''' Zusammenstellung der Werte für den Energiebedarf.
    ''' </summary>
    Function Anzahl_Energiebedarf() As Integer
        '--------------------------------------------------
        Anzahl_Energiebedarf = 0
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            Dim Wert_X As Integer
            Dim Anzahl As Integer = 0
            '--------------------------------------------------
            Dim Maximale_Anzahl As Integer = 200
            '--------------------------------------------------
            With Variable_XML_Import
                '--------------------------------------------------
                For Wert_X = 1 To Maximale_Anzahl
                    '--------------------------------------------------
                    If .Energietraeger(Wert_X).Energietraegerbezeichnung.Wert <> "" Then
                        '--------------------------------------------------
                        Anzahl += 1
                        '--------------------------------------------------
                    End If
                    '--------------------------------------------------
                Next
                '--------------------------------------------------
                Anzahl_Energiebedarf = Anzahl
                '--------------------------------------------------
            End With
            '--------------------------------------------------
        Catch ex As Exception
            Fehlerfenster(ex)
        End Try
        '--------------------------------------------------
    End Function
    '--------------------------------------------------
#End Region
    '--------------------------------------------------
#Region "Zusammenstellung der Werte für die Zonen"
    '--------------------------------------------------
    ''' <summary>
    ''' Zusammenstellung der Werte für die Zonen.
    ''' </summary>
    Function Anzahl_Zonen() As Integer
        '--------------------------------------------------
        Anzahl_Zonen = 0
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            Dim Wert_X As Integer
            Dim Anzahl As Integer = 0
            '--------------------------------------------------
            Dim Maximale_Anzahl As Integer = 200
            '--------------------------------------------------
            With Variable_XML_Import
                '--------------------------------------------------
                For Wert_X = 1 To Maximale_Anzahl
                    '--------------------------------------------------
                    If .Energieausweis_Daten.Nichtwohngebaeude.Zone(Wert_X).Zonenbezeichnung.Wert <> "" Then
                        '--------------------------------------------------
                        Anzahl += 1
                        '--------------------------------------------------
                    End If
                    '--------------------------------------------------
                Next
                '--------------------------------------------------
                Anzahl_Zonen = Anzahl
                '--------------------------------------------------
            End With
            '--------------------------------------------------
        Catch ex As Exception
            Fehlerfenster(ex)
        End Try
        '--------------------------------------------------
    End Function
    '--------------------------------------------------
#End Region
    '--------------------------------------------------
#Region "Zusammenstellung Erneuerbare_Energien_65EE_Regel"
    '--------------------------------------------------
    ''' <summary>
    ''' Zusammenstellung Erneuerbare_Energien_65EE_Regel - Anzahl.
    ''' </summary>
    Function Anzahl_Erneuerbare_Energien_65EE_Regel() As Integer
        '--------------------------------------------------
        Anzahl_Erneuerbare_Energien_65EE_Regel = 0
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            Dim Wert_X As Integer
            Dim Anzahl As Integer = 0
            '--------------------------------------------------
            Dim Maximale_Anzahl As Integer = 200
            '--------------------------------------------------
            With Variable_XML_Import
                '--------------------------------------------------
                For Wert_X = 1 To Maximale_Anzahl
                    '--------------------------------------------------
                    If .EE_65EE_Regel(Wert_X).Art_der_Nutzung_erneuerbaren_Energie.Wert <> "" Then
                        '--------------------------------------------------
                        Anzahl += 1
                        '--------------------------------------------------
                    End If
                    '--------------------------------------------------
                Next
                '--------------------------------------------------
                Anzahl_Erneuerbare_Energien_65EE_Regel = Anzahl
                '--------------------------------------------------
            End With
            '--------------------------------------------------
        Catch ex As Exception
            Fehlerfenster(ex)
        End Try
        '--------------------------------------------------
    End Function
    '--------------------------------------------------
    ''' <summary>
    ''' Zusammenstellung Erneuerbare_Energien_65EE_Regel - Gesamtsumme.
    ''' </summary>
    Function Gesamtsumme_Erneuerbare_Energien_65EE_Regel() As Integer
        '--------------------------------------------------
        Gesamtsumme_Erneuerbare_Energien_65EE_Regel = 0
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            Dim Wert_X As Integer
            Dim Summe As Integer = 0
            '--------------------------------------------------
            Dim Maximale_Anzahl As Integer = 200
            '--------------------------------------------------
            With Variable_XML_Import
                '--------------------------------------------------
                For Wert_X = 1 To Maximale_Anzahl
                    '--------------------------------------------------
                    If .EE_65EE_Regel(Wert_X).Anteil_der_Pflichterfuellung_Gesamt.Wert <> 0 Then
                        '--------------------------------------------------
                        Summe += .EE_65EE_Regel(Wert_X).Anteil_der_Pflichterfuellung_Gesamt.Wert
                        '--------------------------------------------------
                    End If
                    '--------------------------------------------------
                Next
                '--------------------------------------------------
                Gesamtsumme_Erneuerbare_Energien_65EE_Regel = Summe
                '--------------------------------------------------
            End With
            '--------------------------------------------------
        Catch ex As Exception
            Fehlerfenster(ex)
        End Try
        '--------------------------------------------------
    End Function
    '--------------------------------------------------
#End Region
    '--------------------------------------------------
#Region "Zusammenstellung Erneuerbare_Energien_65EE_keine_Regel"
    '--------------------------------------------------
    ''' <summary>
    ''' Zusammenstellung Erneuerbare_Energien_65EE_keine_Regel - Anzahl.
    ''' </summary>
    Function Anzahl_Erneuerbare_Energien_65EE_keine_Regel() As Integer
        '--------------------------------------------------
        Anzahl_Erneuerbare_Energien_65EE_keine_Regel = 0
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            Dim Wert_X As Integer
            Dim Anzahl As Integer = 0
            '--------------------------------------------------
            Dim Maximale_Anzahl As Integer = 200
            '--------------------------------------------------
            With Variable_XML_Import
                '--------------------------------------------------
                For Wert_X = 1 To Maximale_Anzahl
                    '--------------------------------------------------
                    If .EE_65EE_keine_Regel(Wert_X).Art_der_Nutzung_erneuerbaren_Energie.Wert <> "" Then
                        '--------------------------------------------------
                        Anzahl += 1
                        '--------------------------------------------------
                    End If
                    '--------------------------------------------------
                Next
                '--------------------------------------------------
                Anzahl_Erneuerbare_Energien_65EE_keine_Regel = Anzahl
                '--------------------------------------------------
            End With
            '--------------------------------------------------
        Catch ex As Exception
            Fehlerfenster(ex)
        End Try
        '--------------------------------------------------
    End Function
    '--------------------------------------------------
    ''' <summary>
    ''' Zusammenstellung Erneuerbare_Energien_65EE_keine_Regel - Gesamtsumme.
    ''' </summary>
    Function Gesamtsumme_Erneuerbare_Energien_65EE_keine_Regel() As Integer
        '--------------------------------------------------
        Gesamtsumme_Erneuerbare_Energien_65EE_keine_Regel = 0
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            Dim Wert_X As Integer
            Dim Summe As Integer = 0
            '--------------------------------------------------
            Dim Maximale_Anzahl As Integer = 200
            '--------------------------------------------------
            With Variable_XML_Import
                '--------------------------------------------------
                For Wert_X = 1 To Maximale_Anzahl
                    '--------------------------------------------------
                    If .EE_65EE_keine_Regel(Wert_X).Anteil_EE_Anlage.Wert <> 0 Then
                        '--------------------------------------------------
                        Summe += .EE_65EE_keine_Regel(Wert_X).Anteil_EE_Anlage.Wert
                        '--------------------------------------------------
                    End If
                    '--------------------------------------------------
                Next
                '--------------------------------------------------
                Gesamtsumme_Erneuerbare_Energien_65EE_keine_Regel = Summe
                '--------------------------------------------------
            End With
            '--------------------------------------------------
        Catch ex As Exception
            Fehlerfenster(ex)
        End Try
        '--------------------------------------------------
    End Function
    '--------------------------------------------------
#End Region
    '--------------------------------------------------
#Region "Aufteilung des Energiebedarfs"
    '--------------------------------------------------
    ''' <summary>
    ''' Aufteilung des Energiebedarfs - maximal Energiebedarf.
    ''' </summary>
    Function Berechnung_des_maximal_Energiebedarf() As Integer
        '--------------------------------------------------
        Berechnung_des_maximal_Energiebedarf = 0
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            Dim Nutzenergiebedarf As Decimal = 0
            Dim Endenergiebedarf As Decimal = 0
            Dim Primaerenergiebedarf As Decimal = 0
            '--------------------------------------------------
            Dim Maximalwert As Integer
            '--------------------------------------------------
            With Variable_XML_Import.Gebaeudebezogene_Daten.NWG_Aushang_Daten
                '--------------------------------------------------
                Nutzenergiebedarf += .Nutzenergiebedarf_Heizung_Diagramm.Wert
                '--------------------------------------------------
                Nutzenergiebedarf += .Nutzenergiebedarf_Trinkwarmwasser_Diagramm.Wert
                '--------------------------------------------------
                Nutzenergiebedarf += .Nutzenergiebedarf_Beleuchtung_Diagramm.Wert
                '--------------------------------------------------
                Nutzenergiebedarf += .Nutzenergiebedarf_Lueftung_Diagramm.Wert
                '--------------------------------------------------
                Nutzenergiebedarf += .Nutzenergiebedarf_Kuehlung_Befeuchtung_Diagramm.Wert
                '--------------------------------------------------
                '--------------------------------------------------
                Endenergiebedarf += .Endenergiebedarf_Heizung_Diagramm.Wert
                '--------------------------------------------------
                Endenergiebedarf += .Endenergiebedarf_Trinkwarmwasser_Diagramm.Wert
                '--------------------------------------------------
                Endenergiebedarf += .Endenergiebedarf_Beleuchtung_Diagramm.Wert
                '--------------------------------------------------
                Endenergiebedarf += .Endenergiebedarf_Lueftung_Diagramm.Wert
                '--------------------------------------------------
                Endenergiebedarf += .Endenergiebedarf_Kuehlung_Befeuchtung_Diagramm.Wert
                '--------------------------------------------------
                '--------------------------------------------------
                Primaerenergiebedarf += .Primaerenergiebedarf_Heizung_Diagramm.Wert
                '--------------------------------------------------
                Primaerenergiebedarf += .Primaerenergiebedarf_Trinkwarmwasser_Diagramm.Wert
                '--------------------------------------------------
                Primaerenergiebedarf += .Primaerenergiebedarf_Beleuchtung_Diagramm.Wert
                '--------------------------------------------------
                Primaerenergiebedarf += .Primaerenergiebedarf_Lueftung_Diagramm.Wert
                '--------------------------------------------------
                Primaerenergiebedarf += .Primaerenergiebedarf_Kuehlung_Befeuchtung_Diagramm.Wert
                '--------------------------------------------------
            End With
            '--------------------------------------------------
            Maximalwert = Math.Max(Nutzenergiebedarf, Math.Max(Endenergiebedarf, Primaerenergiebedarf))
            '--------------------------------------------------
            Select Case Maximalwert
                Case < 10
                    Berechnung_des_maximal_Energiebedarf = 10
                Case < 50
                    Berechnung_des_maximal_Energiebedarf = 50
                Case < 100
                    Berechnung_des_maximal_Energiebedarf = 100
                Case < 200
                    Berechnung_des_maximal_Energiebedarf = 200
                Case < 500
                    Berechnung_des_maximal_Energiebedarf = 500
                Case < 1000
                    Berechnung_des_maximal_Energiebedarf = 1000
                Case < 2000
                    Berechnung_des_maximal_Energiebedarf = 2000
                Case < 5000
                    Berechnung_des_maximal_Energiebedarf = 5000
                Case < 10000
                    Berechnung_des_maximal_Energiebedarf = 10000
                Case < 20000
                    Berechnung_des_maximal_Energiebedarf = 20000
                Case < 50000
                    Berechnung_des_maximal_Energiebedarf = 50000
                Case Else
                    Berechnung_des_maximal_Energiebedarf = 0
            End Select
            '--------------------------------------------------
        Catch ex As Exception
            Fehlerfenster(ex)
        End Try
        '--------------------------------------------------
    End Function
    '--------------------------------------------------
    ''' <summary>
    ''' Aufteilung des Energiebedarfs - Skalenteilung.
    ''' </summary>
    Function Berechnung_der_Skalenteilung_Energiebedarf(ByVal maximaler_Energiebedarf As Integer) As Integer
        '--------------------------------------------------
        Berechnung_der_Skalenteilung_Energiebedarf = 0
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            Select Case maximaler_Energiebedarf
                Case 10
                    Berechnung_der_Skalenteilung_Energiebedarf = 2
                Case 50
                    Berechnung_der_Skalenteilung_Energiebedarf = 10
                Case 100
                    Berechnung_der_Skalenteilung_Energiebedarf = 20
                Case 200
                    Berechnung_der_Skalenteilung_Energiebedarf = 40
                Case 500
                    Berechnung_der_Skalenteilung_Energiebedarf = 100
                Case 1000
                    Berechnung_der_Skalenteilung_Energiebedarf = 200
                Case 2000
                    Berechnung_der_Skalenteilung_Energiebedarf = 400
                Case 5000
                    Berechnung_der_Skalenteilung_Energiebedarf = 1000
                Case 10000
                    Berechnung_der_Skalenteilung_Energiebedarf = 2000
                Case 20000
                    Berechnung_der_Skalenteilung_Energiebedarf = 4000
                Case 50000
                    Berechnung_der_Skalenteilung_Energiebedarf = 10000
                Case Else
                    Berechnung_der_Skalenteilung_Energiebedarf = 0
            End Select
            '--------------------------------------------------
        Catch ex As Exception
            Fehlerfenster(ex)
        End Try
        '--------------------------------------------------
    End Function
    '--------------------------------------------------
#End Region
    '--------------------------------------------------
End Module
