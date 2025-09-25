#Region "Imports"
'--------------------------------------------------
'--------------------------------------------------
#End Region
'--------------------------------------------------
Module Modul_Webdienst
    '--------------------------------------------------
#Region "Funktionen Webdienst"
    '--------------------------------------------------
    ''' <summary>
    ''' In dieser Funktion wird der Webservice angesteuert und eine Regriernummer angefordert.
    ''' </summary>
    Function Webservice_Datenregistratur() As String
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            With Variable_Steuerung
                '--------------------------------------------------
                Dim DLL_Datenregistratur As New BBSR_Webservice.Klasse_Webservice(Variable_Parameter.Aussteller_ID_DIBT, Variable_Parameter.Aussteller_PWD_DIBT, .Ausstellungsdatum, .Bundesland, .Postleitzahl, .Gesetzesgrundlage, .Gebaeudeart, .Berechnungsart, .Neubau, Variable_Steuerung.Sandbox)
                '--------------------------------------------------
                DLL_Datenregistratur.Webservice_Datenregistratur()
                '--------------------------------------------------
                Webservice_Datenregistratur = DLL_Datenregistratur.DLL_Ergebnis_Datenregistratur
                '--------------------------------------------------
                XML_Datei_Datenregistratur_auswerten(DLL_Datenregistratur.DLL_Ergebnis_Datenregistratur)
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
    ''' In dieser Anweisung wird der Webservice angesteuert um das Restkontingent abzufragen.
    ''' </summary>
    Function Webservice_Restkontingent() As Integer
        '--------------------------------------------------
        Webservice_Restkontingent = 0
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            With Variable_Steuerung
                '--------------------------------------------------
                Dim DLL_Restkontingent As New BBSR_Webservice.Klasse_Webservice(Variable_Parameter.Aussteller_ID_DIBT, Variable_Parameter.Aussteller_PWD_DIBT, Variable_Steuerung.Sandbox)
                '--------------------------------------------------
                DLL_Restkontingent.Webservice_Restkontingent()
                '--------------------------------------------------
                Webservice_Restkontingent = DLL_Restkontingent.DLL_Ergebnis_Restkontingent_Anzahl
                '--------------------------------------------------
                Variable_Steuerung.Restkontingent_Anzahl = DLL_Restkontingent.DLL_Ergebnis_Restkontingent_Anzahl
                Variable_Steuerung.Restkontingent_Fehlermeldung = DLL_Restkontingent.DLL_Ergebnis_Restkontingent_Fehlermeldung
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
    ''' In dieser Anweisung wird der Webservice angesteuert um die Kontrolldatei zu prüfen.
    ''' </summary>
    Function Webservice_KontrolldateiPruefen() As Boolean
        '--------------------------------------------------
        Webservice_KontrolldateiPruefen = False
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            With Variable_Steuerung
                '--------------------------------------------------
                Dim DLL_Validierung_Kontrolldatei As New BBSR_Webservice.Klasse_Webservice(Variable_Parameter.Aussteller_ID_DIBT, Variable_Parameter.Aussteller_PWD_DIBT, Variable_Steuerung.Registriernummer, Variable_Steuerung.Kontrolldatei, Variable_Steuerung.Sandbox)
                '--------------------------------------------------
                DLL_Validierung_Kontrolldatei.Webservice_KontrolldateiPruefen()
                '--------------------------------------------------
                Webservice_KontrolldateiPruefen = DLL_Validierung_Kontrolldatei.DLL_Ergebnis_KontrolldateiPruefen
                '--------------------------------------------------
                Variable_Steuerung.KontrolldateiPruefen_Erfolgreich = DLL_Validierung_Kontrolldatei.DLL_Ergebnis_KontrolldateiPruefen
                Variable_Steuerung.KontrolldateiPruefen_Fehler_Anzahl = DLL_Validierung_Kontrolldatei.DLL_Ergebnis_KontrolldateiPruefen_Fehlerliste_Anzahl
                Variable_Steuerung.KontrolldateiPruefen_Fehler_ID = DLL_Validierung_Kontrolldatei.DLL_Ergebnis_KontrolldateiPruefen_Fehlerliste_ID
                Variable_Steuerung.KontrolldateiPruefen_Fehler_Kurztext = DLL_Validierung_Kontrolldatei.DLL_Ergebnis_KontrolldateiPruefen_Fehlerliste_Kurztext
                Variable_Steuerung.KontrolldateiPruefen_Fehler_Langtext = DLL_Validierung_Kontrolldatei.DLL_Ergebnis_KontrolldateiPruefen_Fehlerliste_Langtext
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
    ''' In dieser Funktion wird der Webservice angesteuert um die Zusatzdaten zu erfassen.
    ''' </summary>
    Function Webservice_ZusatzdatenErfassung() As String
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            With Variable_Steuerung
                '--------------------------------------------------
                Dim DLL_ZusatzdatenErfassung As New BBSR_Webservice.Klasse_Webservice(Variable_Parameter.Aussteller_ID_DIBT, Variable_Parameter.Aussteller_PWD_DIBT, Variable_Steuerung.Registriernummer, Variable_Steuerung.Kontrolldatei, Variable_Steuerung.Sandbox)
                '--------------------------------------------------
                DLL_ZusatzdatenErfassung.Webservice_ZusatzdatenErfassung()
                '--------------------------------------------------
                Webservice_ZusatzdatenErfassung = DLL_ZusatzdatenErfassung.DLL_Ergebnis_ZusatzdatenErfassung
                '--------------------------------------------------
                XML_Datei_ZusatzdatenErfassung_auswerten(DLL_ZusatzdatenErfassung.DLL_Ergebnis_ZusatzdatenErfassung)
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
    ''' In dieser Anweisung wird der Webservice angesteuert um die Liste der offenen Kontrolldateien abzufragen.
    ''' </summary>
    Function Webservice_OffeneKontrolldateien() As Integer
        '--------------------------------------------------
        Webservice_OffeneKontrolldateien = 0
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            With Variable_Steuerung
                '--------------------------------------------------
                Dim DLL_OffeneKontrolldateien As New BBSR_Webservice.Klasse_Webservice(Variable_Parameter.Aussteller_ID_DIBT, Variable_Parameter.Aussteller_PWD_DIBT, Variable_Steuerung.Sandbox)
                '--------------------------------------------------
                DLL_OffeneKontrolldateien.Webservice_OffeneKontrolldateien()
                '--------------------------------------------------
                Webservice_OffeneKontrolldateien = DLL_OffeneKontrolldateien.DLL_Ergebnis_OffeneKontrolldateien_Ausweis_Anzahl
                '--------------------------------------------------
                Variable_Steuerung.OffeneKontrolldateien_Anzahl = DLL_OffeneKontrolldateien.DLL_Ergebnis_OffeneKontrolldateien_Ausweis_Anzahl
                Variable_Steuerung.OffeneKontrolldateien_Ausweis_Registriernummer = DLL_OffeneKontrolldateien.DLL_Ergebnis_OffeneKontrolldateien_Ausweis_Registriernummer
                Variable_Steuerung.OffeneKontrolldateien_Ausweis_NummerErzeugtAm = DLL_OffeneKontrolldateien.DLL_Ergebnis_OffeneKontrolldateien_Ausweis_NummerErzeugtAm
                Variable_Steuerung.OffeneKontrolldateien_Ausweis_Aussteller = DLL_OffeneKontrolldateien.DLL_Ergebnis_OffeneKontrolldateien_Ausweis_Aussteller
                Variable_Steuerung.OffeneKontrolldateien_Fehlermeldung = DLL_OffeneKontrolldateien.DLL_Ergebnis_OffeneKontrolldateien_Fehler
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
    ''' In dieser Anweisung wird geprüft ob eine Verbindung zum Internet besteht.
    ''' </summary>
    Public Function Webservice_Internettest() As Boolean
        '--------------------------------------------------
        Try
            Dim oRequest As Net.HttpWebRequest = Net.HttpWebRequest.Create("https://energieausweis.dibt.de")
            With oRequest
                .Proxy = System.Net.WebRequest.DefaultWebProxy
                .Credentials = System.Net.CredentialCache.DefaultCredentials
            End With
            Dim oResponse As Net.WebResponse = oRequest.GetResponse
            oResponse.Close()
            Return False
        Catch ex As Exception
            Return True
        End Try
        '--------------------------------------------------
    End Function
    '--------------------------------------------------
#End Region
    '--------------------------------------------------
End Module
