'--------------------------------------------------
#Region "Imports"
Imports System.IO
Imports System.Windows.Forms
#End Region
'--------------------------------------------------
Module Modul_Ablauf_ohne_Oberflaeche
    '--------------------------------------------------
#Region "Ablauf ohne Oberflaeche"
    '--------------------------------------------------
    ''' <summary>
    ''' In dieser Anweisung wird der Ablauf ohne Oberflaeche durchgeführt.
    ''' </summary>
    Public Sub Main()
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            Select Case Variable_Steuerung.Methode
                Case "Datenregistratur"
                    '--------------------------------------------------
                    Webservice_Datenregistratur_ansteuern_und_auswerten()
                    '--------------------------------------------------
                    Steuerungsdatei_erstellen()
                    '--------------------------------------------------
                Case "Energieausweis-Vorschau"
                    '--------------------------------------------------
                    Kontrolldatei_lesen()
                    '--------------------------------------------------
                    XML_Kontrolldatei_bearbeiten_und_speichern()
                    '--------------------------------------------------
                    Webservice_KontrolldateiPruefen_ansteuern_und_auswerten()
                    '--------------------------------------------------
                    Variable_Steuerung.Entwurf = True
                    '--------------------------------------------------
                    Energeiausweis_erstellen()
                    '--------------------------------------------------
                    Steuerungsdatei_erstellen()
                    '--------------------------------------------------
                Case "Energieausweis-Vorschau-ohne-Pruefung"
                    '--------------------------------------------------
                    Kontrolldatei_lesen()
                    '--------------------------------------------------
                    XML_Kontrolldatei_bearbeiten_und_speichern()
                    '--------------------------------------------------
                    Variable_Steuerung.Entwurf = True
                    '--------------------------------------------------
                    Energeiausweis_erstellen()
                    '--------------------------------------------------
                    Steuerungsdatei_erstellen()
                    '--------------------------------------------------
                Case "Energieausweis-Endgueltig"
                    '--------------------------------------------------
                    Variable_Steuerung.Entwurf = False
                    '--------------------------------------------------
                    Kontrolldatei_lesen()
                    '--------------------------------------------------
                    XML_Kontrolldatei_bearbeiten_und_speichern()
                    '--------------------------------------------------
                    Webservice_KontrolldateiPruefen_ansteuern_und_auswerten()
                    '--------------------------------------------------
                    Webservice_Zusatzdatenerfassung_ansteuern_und_auswerten()
                    '--------------------------------------------------
                    Energeiausweis_erstellen()
                    '--------------------------------------------------
                    Steuerungsdatei_erstellen()
                    '--------------------------------------------------
                Case "Energieausweis-Endgueltig-ohne-Zusatzdatenerfassung"
                    '--------------------------------------------------
                    Variable_Steuerung.Entwurf = False
                    '--------------------------------------------------
                    Kontrolldatei_lesen()
                    '--------------------------------------------------
                    XML_Kontrolldatei_bearbeiten_und_speichern()
                    '--------------------------------------------------
                    Webservice_KontrolldateiPruefen_ansteuern_und_auswerten()
                    '--------------------------------------------------
                    Energeiausweis_erstellen()
                    '--------------------------------------------------
                    Steuerungsdatei_erstellen()
                    '--------------------------------------------------
                Case "Energieausweis-Endgueltig-ohne-Pruefung"
                    '--------------------------------------------------
                    Variable_Steuerung.Entwurf = False
                    '--------------------------------------------------
                    Kontrolldatei_lesen()
                    '--------------------------------------------------
                    XML_Kontrolldatei_bearbeiten_und_speichern()
                    '--------------------------------------------------
                    Webservice_Zusatzdatenerfassung_ansteuern_und_auswerten()
                    '--------------------------------------------------
                    Energeiausweis_erstellen()
                    '--------------------------------------------------
                    Steuerungsdatei_erstellen()
                    '--------------------------------------------------
                Case "Energieausweis-Endgueltig-ohne-Pruefung-ohne-Zusatzdatenerfassung"
                    '--------------------------------------------------
                    Variable_Steuerung.Entwurf = False
                    '--------------------------------------------------
                    Kontrolldatei_lesen()
                    '--------------------------------------------------
                    XML_Kontrolldatei_bearbeiten_und_speichern()
                    '--------------------------------------------------
                    Energeiausweis_erstellen()
                    '--------------------------------------------------
                    Steuerungsdatei_erstellen()
                    '--------------------------------------------------
                Case "Energieausweis-Komplett"
                    '--------------------------------------------------
                    Variable_Steuerung.Entwurf = False
                    '--------------------------------------------------
                    Webservice_Datenregistratur_ansteuern_und_auswerten()
                    '--------------------------------------------------
                    Kontrolldatei_lesen()
                    '--------------------------------------------------
                    XML_Kontrolldatei_bearbeiten_und_speichern()
                    '--------------------------------------------------
                    Webservice_KontrolldateiPruefen_ansteuern_und_auswerten()
                    '--------------------------------------------------
                    Webservice_Zusatzdatenerfassung_ansteuern_und_auswerten()
                    '--------------------------------------------------
                    Energeiausweis_erstellen()
                    '--------------------------------------------------
                    Steuerungsdatei_erstellen()
                    '--------------------------------------------------
                Case "Energieausweis-Komplett-ohne-Zusatzdatenerfassung"
                    '--------------------------------------------------
                    Variable_Steuerung.Entwurf = False
                    '--------------------------------------------------
                    Webservice_Datenregistratur_ansteuern_und_auswerten()
                    '--------------------------------------------------
                    Kontrolldatei_lesen()
                    '--------------------------------------------------
                    XML_Kontrolldatei_bearbeiten_und_speichern()
                    '--------------------------------------------------
                    Webservice_KontrolldateiPruefen_ansteuern_und_auswerten()
                    '--------------------------------------------------
                    Energeiausweis_erstellen()
                    '--------------------------------------------------
                    Steuerungsdatei_erstellen()
                    '--------------------------------------------------
                Case "Energieausweis-Komplett-ohne-Pruefung"
                    '--------------------------------------------------
                    Variable_Steuerung.Entwurf = False
                    '--------------------------------------------------
                    Webservice_Datenregistratur_ansteuern_und_auswerten()
                    '--------------------------------------------------
                    Kontrolldatei_lesen()
                    '--------------------------------------------------
                    XML_Kontrolldatei_bearbeiten_und_speichern()
                    '--------------------------------------------------
                    Webservice_Zusatzdatenerfassung_ansteuern_und_auswerten()
                    '--------------------------------------------------
                    Energeiausweis_erstellen()
                    '--------------------------------------------------
                    Steuerungsdatei_erstellen()
                    '--------------------------------------------------
                Case "Energieausweis-Komplett-ohne-Pruefung-ohne-Zusatzdatenerfassung"
                    '--------------------------------------------------
                    Variable_Steuerung.Entwurf = False
                    '--------------------------------------------------
                    Webservice_Datenregistratur_ansteuern_und_auswerten()
                    '--------------------------------------------------
                    Kontrolldatei_lesen()
                    '--------------------------------------------------
                    XML_Kontrolldatei_bearbeiten_und_speichern()
                    '--------------------------------------------------
                    Energeiausweis_erstellen()
                    '--------------------------------------------------
                    Steuerungsdatei_erstellen()
                    '--------------------------------------------------
                Case "Restkontingent"
                    '--------------------------------------------------
                    Webservice_Restkontingent_ansteuern_und_auswerten()
                    '--------------------------------------------------
                    Steuerungsdatei_erstellen()
                    '--------------------------------------------------
                Case "KontrolldateiPruefen"
                    '--------------------------------------------------
                    Kontrolldatei_lesen()
                    '--------------------------------------------------
                    XML_Kontrolldatei_bearbeiten_und_speichern()
                    '--------------------------------------------------
                    Webservice_KontrolldateiPruefen_ansteuern_und_auswerten()
                    '--------------------------------------------------
                    Steuerungsdatei_erstellen()
                    '--------------------------------------------------
                Case "Zusatzdatenerfassung"
                    '--------------------------------------------------
                    Kontrolldatei_lesen()
                    '--------------------------------------------------
                    XML_Kontrolldatei_bearbeiten_und_speichern()
                    '--------------------------------------------------
                    Webservice_KontrolldateiPruefen_ansteuern_und_auswerten()
                    '--------------------------------------------------
                    Webservice_Zusatzdatenerfassung_ansteuern_und_auswerten()
                    '--------------------------------------------------
                    Steuerungsdatei_erstellen()
                    '--------------------------------------------------
                Case "OffeneKontrolldateien"
                    '--------------------------------------------------
                    Webservice_OffeneKontrolldateien_ansteuern_und_auswerten()
                    '--------------------------------------------------
                    Steuerungsdatei_erstellen()
                    '--------------------------------------------------
            End Select
            '--------------------------------------------------
            If Variable_XML_Import.Bildanzahl > 0 Then
                '--------------------------------------------------
                PDF_erzeugen(Variable_Steuerung.PDF_oeffnen)
                '--------------------------------------------------
            End If
            '--------------------------------------------------
            Application.Exit()
            '--------------------------------------------------
            Windows.Forms.Cursor.Current = Cursors.Default
            '--------------------------------------------------
        Catch ex As Exception
            Fehlerfenster(ex)
        End Try
        '--------------------------------------------------
    End Sub
    '-------------------------------------------------------
    ''' <summary>
    ''' In dieser Anweisung wird das ansteuern des Webservice und das Auslesen der Ergebnisdatei gesteuert.
    ''' </summary>
    Private Sub Webservice_Datenregistratur_ansteuern_und_auswerten()
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            Variable_Steuerung.Datenregistratur = True
            '--------------------------------------------------
            If Variable_Steuerung.Ablauf_Fehler = False Then
                '--------------------------------------------------
                Dim XML_Datei_Webservice As String = Webservice_Datenregistratur()
                '--------------------------------------------------
                If Variable_Steuerung.Datenregistratur_Fehler_Anzahl = 0 Then

                Else
                    '--------------------------------------------------
                    Variable_Steuerung.Ablauf_Fehler = True
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
    ''' In dieser Anweisung wird das ansteuern des Webservice und das Anfragen des Restkontingent gesteuert.
    ''' </summary>
    Private Sub Webservice_Restkontingent_ansteuern_und_auswerten()
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            Variable_Steuerung.Restkontingent = True
            '--------------------------------------------------
            If Variable_Steuerung.Ablauf_Fehler = False Then
                '--------------------------------------------------
                Dim Ergebnis_Webservice_Restkontingent As Integer = Webservice_Restkontingent()
                '--------------------------------------------------
                If Variable_Steuerung.Restkontingent_Fehlermeldung = "" Then

                Else
                    '--------------------------------------------------
                    Variable_Steuerung.Ablauf_Fehler = True
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
    ''' In dieser Anweisung wird das ansteuern des Webservice und das Prüfen der Kontrolldatei gesteuert.
    ''' </summary>
    Private Sub Webservice_KontrolldateiPruefen_ansteuern_und_auswerten()
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            Variable_Steuerung.KontrolldateiPruefen = True
            '--------------------------------------------------
            If Variable_Steuerung.Ablauf_Fehler = False Then
                '--------------------------------------------------
                If Webservice_KontrolldateiPruefen() = False Then
                    '--------------------------------------------------
                    Variable_Steuerung.Entwurf = False
                    '--------------------------------------------------
                Else
                    '--------------------------------------------------
                    Variable_Steuerung.Ablauf_Fehler = True
                    '--------------------------------------------------
                    Variable_Steuerung.Entwurf = True
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
    ''' In dieser Anweisung wird das ansteuern des Webservice und das Hochladen der Kontrolldatei gesteuert.
    ''' </summary>
    Private Sub Webservice_Zusatzdatenerfassung_ansteuern_und_auswerten()
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            Variable_Steuerung.ZusatzdatenErfassung = True
            '--------------------------------------------------
            If Variable_Steuerung.Ablauf_Fehler = False Then
                '--------------------------------------------------
                Dim XML_Datei_Webservice As String = Webservice_ZusatzdatenErfassung()
                '--------------------------------------------------
                If Variable_Steuerung.ZusatzdatenErfassung_Fehler_Anzahl = 0 Then
                    '--------------------------------------------------
                    Variable_Steuerung.Entwurf = False
                    '--------------------------------------------------
                Else
                    '--------------------------------------------------
                    Variable_Steuerung.Ablauf_Fehler = True
                    '--------------------------------------------------
                    Variable_Steuerung.Entwurf = True
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
    ''' In dieser Anweisung wird das ansteuern des Webservice und das lesen der Offenen Kontrolldateien gesteuert.
    ''' </summary>
    Private Sub Webservice_OffeneKontrolldateien_ansteuern_und_auswerten()
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            Variable_Steuerung.OffeneKontrolldateien = True
            '--------------------------------------------------
            If Variable_Steuerung.Ablauf_Fehler = False Then
                '--------------------------------------------------
                Dim Ergebnis_Webservice_OffeneKontrolldateien As Integer = Webservice_OffeneKontrolldateien()
                '--------------------------------------------------
                If Variable_Steuerung.OffeneKontrolldateien_Fehlermeldung = "" Then

                Else
                    '--------------------------------------------------
                    Variable_Steuerung.Ablauf_Fehler = True
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
    ''' In dieser Anweisung wird die PDF Erstellung des Energieausweises gestartet.
    ''' </summary>
    Private Sub Energeiausweis_erstellen()
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            Dim Me_PictureBox As New PictureBox
            Dim Me_ListView As New ListView
            Dim Me_ImageList As New ImageList
            '--------------------------------------------------
            If Variable_Steuerung.Ablauf_Fehler = False Then
                '--------------------------------------------------
                Select Case Variable_Steuerung.Gesetzesgrundlage
                    Case "ENEV-2016"
                        '--------------------------------------------------
                        Select Case Variable_Steuerung.Gebaeudeart
                            Case "Wohngebäude", "Wohnteil gemischt genutztes Gebäude"
                                '--------------------------------------------------
                                Dim EnergieausweisWB2016 As New Klasse_Energieausweis_WB_2016
                                EnergieausweisWB2016.Energieausweis_erzeugen(False, Me_PictureBox, Me_ListView, Me_ImageList)
                                '--------------------------------------------------
                            Case "Nichtwohngebäude", "Nichtwohnteil gemischt genutztes Gebäude"
                                '--------------------------------------------------
                                Dim EnergieausweisNWB2016 As New Klasse_Energieausweis_NWB_2016
                                EnergieausweisNWB2016.Energieausweis_erzeugen(False, Me_PictureBox, Me_ListView, Me_ImageList)
                                '--------------------------------------------------
                        End Select
                        '--------------------------------------------------
                    Case "GEG-2020"
                        '--------------------------------------------------
                        Select Case Variable_Steuerung.Gebaeudeart
                            Case "Wohngebäude", "Wohnteil gemischt genutztes Gebäude"
                                '--------------------------------------------------
                                Dim EnergieausweisWB2020 As New Klasse_Energieausweis_WB_2020
                                EnergieausweisWB2020.Energieausweis_erzeugen(False, Me_PictureBox, Me_ListView, Me_ImageList)
                                '--------------------------------------------------
                            Case "Nichtwohngebäude", "Nichtwohnteil gemischt genutztes Gebäude"
                                '--------------------------------------------------
                                Dim EnergieausweisNWB2020 As New Klasse_Energieausweis_NWB_2020
                                EnergieausweisNWB2020.Energieausweis_erzeugen(False, Me_PictureBox, Me_ListView, Me_ImageList)
                                '--------------------------------------------------
                        End Select
                        '--------------------------------------------------
                    Case "GEG-2023"
                        '--------------------------------------------------
                        Select Case Variable_Steuerung.Gebaeudeart
                            Case "Wohngebäude", "Wohnteil gemischt genutztes Gebäude"
                                '--------------------------------------------------
                                Dim EnergieausweisWB2023 As New Klasse_Energieausweis_WB_2023
                                EnergieausweisWB2023.Energieausweis_erzeugen(False, Me_PictureBox, Me_ListView, Me_ImageList)
                                '--------------------------------------------------
                            Case "Nichtwohngebäude", "Nichtwohnteil gemischt genutztes Gebäude"
                                '--------------------------------------------------
                                Dim EnergieausweisNWB2023 As New Klasse_Energieausweis_NWB_2023
                                EnergieausweisNWB2023.Energieausweis_erzeugen(False, Me_PictureBox, Me_ListView, Me_ImageList)
                                '--------------------------------------------------
                        End Select
                        '--------------------------------------------------
                    Case "GEG-2024"
                        '--------------------------------------------------
                        Select Case Variable_Steuerung.Gebaeudeart
                            Case "Wohngebäude", "Wohnteil gemischt genutztes Gebäude"
                                '--------------------------------------------------
                                Dim EnergieausweisWB2024 As New Klasse_Energieausweis_WB_2024
                                EnergieausweisWB2024.Energieausweis_erzeugen(False, Me_PictureBox, Me_ListView, Me_ImageList)
                                '--------------------------------------------------
                            Case "Nichtwohngebäude", "Nichtwohnteil gemischt genutztes Gebäude"
                                '--------------------------------------------------
                                Dim EnergieausweisNWB2024 As New Klasse_Energieausweis_NWB_2024
                                EnergieausweisNWB2024.Energieausweis_erzeugen(False, Me_PictureBox, Me_ListView, Me_ImageList)
                                '--------------------------------------------------
                        End Select
                        '--------------------------------------------------
                    Case "GEG-20xx"
                        '--------------------------------------------------

                        '--------------------------------------------------
                End Select
                '--------------------------------------------------
            Else
                '--------------------------------------------------
                'Ausdruck wird nicht erstellt, weil es einen Fehler gab.
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
End Module
