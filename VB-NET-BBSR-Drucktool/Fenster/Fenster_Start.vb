'--------------------------------------------------
#Region "Imports"
Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms
#End Region
'--------------------------------------------------
Public Class Fenster_Start
    '--------------------------------------------------
#Region "Variablen"
    '--------------------------------------------------
    Dim Inhalt_Tabelle_Ergebnis As String
    '--------------------------------------------------
    Dim ErgebnisDIBt2025 As New WebBrowser
    '--------------------------------------------------
#End Region
    '--------------------------------------------------
#Region "Seitenstart"
    '--------------------------------------------------
    ''' <summary>
    ''' Startsequenz, es wird ein zeitversetzter Timer gestartet. 
    ''' Die Versionsnummer wird ausgelesen und in die Oberfläche 
    ''' geschrieben. 
    ''' </summary>
    Public Sub New()
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            InitializeComponent()
            '--------------------------------------------------
            Me.Tab2.Controls.Add(ErgebnisDIBt2025)
            ErgebnisDIBt2025.Dock = DockStyle.Fill
            ErgebnisDIBt2025.AllowNavigation = False
            ErgebnisDIBt2025.WebBrowserShortcutsEnabled = False
            '--------------------------------------------------
            Dim oAssembly As System.Reflection.AssemblyName = System.Reflection.Assembly.GetExecutingAssembly().GetName
            Dim sVersion As String = oAssembly.Version.ToString()
            Dim sMajor As String = oAssembly.Version.Major.ToString()
            Dim sMinor As String = oAssembly.Version.Minor.ToString()
            Dim sBuild As String = oAssembly.Version.Build.ToString()
            '--------------------------------------------------
            Dim Zusatz As String
            If Environment.Is64BitProcess = True Then
                '--------------------------------------------------
                Zusatz = "x64"
                '--------------------------------------------------
            Else
                '--------------------------------------------------
                Zusatz = "x86"
                '--------------------------------------------------
            End If
            '--------------------------------------------------
            Me.Text = My.Settings.Title
            StatusLabel.Text = My.Settings.Title & " - Version " & sMajor & "." & sMinor & "." & sBuild & " - " & "[" & Zusatz & "]"
            '--------------------------------------------------
            Label_Software_Name.Text = Me.Text
            Label_Produkt_Version.Text = "Produkt-Version: " & sMajor & "." & sMinor & "." & sBuild & " - " & "[" & Zusatz & "]"
            '--------------------------------------------------
            Timer1.Start()
            '--------------------------------------------------
        Catch ex As Exception
            Fehlerfenster(ex)
        End Try
        '--------------------------------------------------
    End Sub
    '--------------------------------------------------
#End Region
    '--------------------------------------------------
#Region "Start der Druckapplikation"
    '--------------------------------------------------
    ''' <summary>
    ''' Diese Funktion startet die gesamten Ablauf. Innerhalb des Ablaufes werden die unterschiedlichen Abläufe gestartet.
    ''' </summary>
    Private Sub Kompletter_Ablauf(sender As Object, e As EventArgs) Handles Timer1.Tick
        '--------------------------------------------------
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        '--------------------------------------------------
        Timer1.Stop()
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            Inhalt_Tabelle_Ergebnis = ""
            '--------------------------------------------------
            If Variable_Steuerung.Internetverbindung_Fehler = True Then
                Inhalt_Tabelle_Ergebnis += HTML_Tabelle_Ueberschrift_Rot("Eine Verbindung zum Internet konnte nicht hergestellt werden.")
                Inhalt_Tabelle_Ergebnis += HTML_Tabelle_Zeile_Leerzeile()
            Else
                Inhalt_Tabelle_Ergebnis += HTML_Tabelle_Ueberschrift_Gruen("Eine Verbindung zum Internet konnte hergestellt werden.")
                Inhalt_Tabelle_Ergebnis += HTML_Tabelle_Zeile_Leerzeile()
            End If
            '--------------------------------------------------
            Inhalt_Tabelle_Ergebnis += HTML_Tabelle_Abschnitt("Rückmeldung vom DIBt Server")
            '--------------------------------------------------
            Inhalt_Tabelle_Ergebnis += HTML_Tabelle_Zeile_Leerzeile()
            '--------------------------------------------------
            ProgressBar.Visible = True
            '--------------------------------------------------
            ProgressBar.Value = 0
            '--------------------------------------------------
            PictureBox.Image = Image.FromFile(Variable_Parameter.Arbeitsverzeichnis & "/Image/Druckapplikation-2025.png")
            '--------------------------------------------------
            ProgressBar.Value = 20
            '--------------------------------------------------
            Select Case Variable_Steuerung.Methode
                Case "Datenregistratur"
                    '--------------------------------------------------
                    Eigenschaften_Fenster_Start_festlegen(0)
                    '--------------------------------------------------
                    ProgressBar.Value = 30
                    '--------------------------------------------------
                    Webservice_Datenregistratur_ansteuern_und_auswerten()
                    '--------------------------------------------------
                    ProgressBar.Value = 60
                    '--------------------------------------------------
                    Ergebnisprotokoll_erstellen()
                    '--------------------------------------------------
                    Steuerungsdatei_erstellen()
                    '--------------------------------------------------
                    ProgressBar.Value = 90
                    '--------------------------------------------------
                Case "Energieausweis-Vorschau"
                    '--------------------------------------------------
                    Eigenschaften_Fenster_Start_festlegen(1)
                    '--------------------------------------------------
                    Kontrolldatei_lesen()
                    '--------------------------------------------------
                    ProgressBar.Value = 30
                    '--------------------------------------------------
                    XML_Kontrolldatei_bearbeiten_und_speichern()
                    '--------------------------------------------------
                    ProgressBar.Value = 50
                    '--------------------------------------------------
                    Webservice_KontrolldateiPruefen_ansteuern_und_auswerten()
                    '--------------------------------------------------
                    ProgressBar.Value = 70
                    '--------------------------------------------------
                    Variable_Steuerung.Entwurf = True
                    '--------------------------------------------------
                    Energeiausweis_erstellen()
                    '--------------------------------------------------
                    ProgressBar.Value = 80
                    '--------------------------------------------------
                    Ergebnisprotokoll_erstellen()
                    '--------------------------------------------------
                    Steuerungsdatei_erstellen()
                    '--------------------------------------------------
                    ProgressBar.Value = 90
                    '--------------------------------------------------
                Case "Energieausweis-Vorschau-ohne-Pruefung"
                    '--------------------------------------------------
                    Eigenschaften_Fenster_Start_festlegen(1)
                    '--------------------------------------------------
                    Kontrolldatei_lesen()
                    '--------------------------------------------------
                    ProgressBar.Value = 30
                    '--------------------------------------------------
                    XML_Kontrolldatei_bearbeiten_und_speichern()
                    '--------------------------------------------------
                    ProgressBar.Value = 50
                    '--------------------------------------------------
                    Variable_Steuerung.Entwurf = True
                    '--------------------------------------------------
                    Energeiausweis_erstellen()
                    '--------------------------------------------------
                    ProgressBar.Value = 80
                    '--------------------------------------------------
                    Ergebnisprotokoll_erstellen()
                    '--------------------------------------------------
                    Steuerungsdatei_erstellen()
                    '--------------------------------------------------
                    ProgressBar.Value = 90
                    '--------------------------------------------------
                Case "Energieausweis-Endgueltig"
                    '--------------------------------------------------
                    Eigenschaften_Fenster_Start_festlegen(1)
                    '--------------------------------------------------
                    Variable_Steuerung.Entwurf = False
                    '--------------------------------------------------
                    Kontrolldatei_lesen()
                    '--------------------------------------------------
                    ProgressBar.Value = 20
                    '--------------------------------------------------
                    XML_Kontrolldatei_bearbeiten_und_speichern()
                    '--------------------------------------------------
                    ProgressBar.Value = 40
                    '--------------------------------------------------
                    Webservice_KontrolldateiPruefen_ansteuern_und_auswerten()
                    '--------------------------------------------------
                    ProgressBar.Value = 60
                    '--------------------------------------------------
                    Webservice_Zusatzdatenerfassung_ansteuern_und_auswerten()
                    '--------------------------------------------------
                    ProgressBar.Value = 70
                    '--------------------------------------------------
                    Energeiausweis_erstellen()
                    '--------------------------------------------------
                    ProgressBar.Value = 80
                    '--------------------------------------------------
                    Ergebnisprotokoll_erstellen()
                    '--------------------------------------------------
                    Steuerungsdatei_erstellen()
                    '--------------------------------------------------
                    ProgressBar.Value = 90
                    '--------------------------------------------------
                Case "Energieausweis-Endgueltig-ohne-Zusatzdatenerfassung"
                    '--------------------------------------------------
                    Eigenschaften_Fenster_Start_festlegen(1)
                    '--------------------------------------------------
                    Variable_Steuerung.Entwurf = False
                    '--------------------------------------------------
                    Kontrolldatei_lesen()
                    '--------------------------------------------------
                    ProgressBar.Value = 20
                    '--------------------------------------------------
                    XML_Kontrolldatei_bearbeiten_und_speichern()
                    '--------------------------------------------------
                    ProgressBar.Value = 40
                    '--------------------------------------------------
                    Webservice_KontrolldateiPruefen_ansteuern_und_auswerten()
                    '--------------------------------------------------
                    ProgressBar.Value = 60
                    '--------------------------------------------------
                    Energeiausweis_erstellen()
                    '--------------------------------------------------
                    ProgressBar.Value = 80
                    '--------------------------------------------------
                    Ergebnisprotokoll_erstellen()
                    '--------------------------------------------------
                    Steuerungsdatei_erstellen()
                    '--------------------------------------------------
                    ProgressBar.Value = 90
                    '--------------------------------------------------
                Case "Energieausweis-Endgueltig-ohne-Pruefung"
                    '--------------------------------------------------
                    Eigenschaften_Fenster_Start_festlegen(1)
                    '--------------------------------------------------
                    Variable_Steuerung.Entwurf = False
                    '--------------------------------------------------
                    Kontrolldatei_lesen()
                    '--------------------------------------------------
                    ProgressBar.Value = 20
                    '--------------------------------------------------
                    XML_Kontrolldatei_bearbeiten_und_speichern()
                    '--------------------------------------------------
                    ProgressBar.Value = 50
                    '--------------------------------------------------
                    Webservice_Zusatzdatenerfassung_ansteuern_und_auswerten()
                    '--------------------------------------------------
                    ProgressBar.Value = 70
                    '--------------------------------------------------
                    Energeiausweis_erstellen()
                    '--------------------------------------------------
                    ProgressBar.Value = 80
                    '--------------------------------------------------
                    Ergebnisprotokoll_erstellen()
                    '--------------------------------------------------
                    Steuerungsdatei_erstellen()
                    '--------------------------------------------------
                    ProgressBar.Value = 90
                    '--------------------------------------------------
                Case "Energieausweis-Endgueltig-ohne-Pruefung-ohne-Zusatzdatenerfassung"
                    '--------------------------------------------------
                    Eigenschaften_Fenster_Start_festlegen(1)
                    '--------------------------------------------------
                    Variable_Steuerung.Entwurf = False
                    '--------------------------------------------------
                    Kontrolldatei_lesen()
                    '--------------------------------------------------
                    ProgressBar.Value = 20
                    '--------------------------------------------------
                    XML_Kontrolldatei_bearbeiten_und_speichern()
                    '--------------------------------------------------
                    ProgressBar.Value = 50
                    '--------------------------------------------------
                    Energeiausweis_erstellen()
                    '--------------------------------------------------
                    ProgressBar.Value = 80
                    '--------------------------------------------------
                    Ergebnisprotokoll_erstellen()
                    '--------------------------------------------------
                    Steuerungsdatei_erstellen()
                    '--------------------------------------------------
                    ProgressBar.Value = 90
                    '--------------------------------------------------
                Case "Energieausweis-Komplett"
                    '--------------------------------------------------
                    Eigenschaften_Fenster_Start_festlegen(1)
                    '--------------------------------------------------
                    Webservice_Datenregistratur_ansteuern_und_auswerten()
                    '--------------------------------------------------
                    ProgressBar.Value = 20
                    '--------------------------------------------------
                    Kontrolldatei_lesen()
                    '--------------------------------------------------
                    ProgressBar.Value = 30
                    '--------------------------------------------------
                    XML_Kontrolldatei_bearbeiten_und_speichern()
                    '--------------------------------------------------
                    ProgressBar.Value = 40
                    '--------------------------------------------------
                    Webservice_KontrolldateiPruefen_ansteuern_und_auswerten()
                    '--------------------------------------------------
                    ProgressBar.Value = 50
                    '--------------------------------------------------
                    Webservice_Zusatzdatenerfassung_ansteuern_und_auswerten()
                    '--------------------------------------------------
                    ProgressBar.Value = 60
                    '--------------------------------------------------
                    Energeiausweis_erstellen()
                    '--------------------------------------------------
                    ProgressBar.Value = 70
                    '--------------------------------------------------
                    Ergebnisprotokoll_erstellen()
                    '--------------------------------------------------
                    Steuerungsdatei_erstellen()
                    '--------------------------------------------------
                    ProgressBar.Value = 90
                    '--------------------------------------------------
                Case "Energieausweis-Komplett-ohne-Zusatzdatenerfassung"
                    '--------------------------------------------------
                    Eigenschaften_Fenster_Start_festlegen(1)
                    '--------------------------------------------------
                    Webservice_Datenregistratur_ansteuern_und_auswerten()
                    '--------------------------------------------------
                    ProgressBar.Value = 20
                    '--------------------------------------------------
                    Kontrolldatei_lesen()
                    '--------------------------------------------------
                    ProgressBar.Value = 30
                    '--------------------------------------------------
                    XML_Kontrolldatei_bearbeiten_und_speichern()
                    '--------------------------------------------------
                    ProgressBar.Value = 40
                    '--------------------------------------------------
                    Webservice_KontrolldateiPruefen_ansteuern_und_auswerten()
                    '--------------------------------------------------
                    ProgressBar.Value = 50
                    '--------------------------------------------------
                    Energeiausweis_erstellen()
                    '--------------------------------------------------
                    ProgressBar.Value = 70
                    '--------------------------------------------------
                    Ergebnisprotokoll_erstellen()
                    '--------------------------------------------------
                    Steuerungsdatei_erstellen()
                    '--------------------------------------------------
                    ProgressBar.Value = 90
                    '--------------------------------------------------
                Case "Energieausweis-Komplett-ohne-Pruefung"
                    '--------------------------------------------------
                    Eigenschaften_Fenster_Start_festlegen(1)
                    '--------------------------------------------------
                    Webservice_Datenregistratur_ansteuern_und_auswerten()
                    '--------------------------------------------------
                    ProgressBar.Value = 20
                    '--------------------------------------------------
                    Kontrolldatei_lesen()
                    '--------------------------------------------------
                    ProgressBar.Value = 30
                    '--------------------------------------------------
                    XML_Kontrolldatei_bearbeiten_und_speichern()
                    '--------------------------------------------------
                    ProgressBar.Value = 50
                    '--------------------------------------------------
                    Webservice_Zusatzdatenerfassung_ansteuern_und_auswerten()
                    '--------------------------------------------------
                    ProgressBar.Value = 60
                    '--------------------------------------------------
                    Energeiausweis_erstellen()
                    '--------------------------------------------------
                    ProgressBar.Value = 70
                    '--------------------------------------------------
                    Ergebnisprotokoll_erstellen()
                    '--------------------------------------------------
                    Steuerungsdatei_erstellen()
                    '--------------------------------------------------
                    ProgressBar.Value = 90
                    '--------------------------------------------------
                Case "Energieausweis-Komplett-ohne-Pruefung-ohne-Zusatzdatenerfassung"
                    '--------------------------------------------------
                    Eigenschaften_Fenster_Start_festlegen(1)
                    '--------------------------------------------------
                    Webservice_Datenregistratur_ansteuern_und_auswerten()
                    '--------------------------------------------------
                    ProgressBar.Value = 20
                    '--------------------------------------------------
                    Kontrolldatei_lesen()
                    '--------------------------------------------------
                    ProgressBar.Value = 30
                    '--------------------------------------------------
                    XML_Kontrolldatei_bearbeiten_und_speichern()
                    '--------------------------------------------------
                    ProgressBar.Value = 50
                    '--------------------------------------------------
                    Energeiausweis_erstellen()
                    '--------------------------------------------------
                    ProgressBar.Value = 70
                    '--------------------------------------------------
                    Ergebnisprotokoll_erstellen()
                    '--------------------------------------------------
                    Steuerungsdatei_erstellen()
                    '--------------------------------------------------
                    ProgressBar.Value = 90
                    '--------------------------------------------------
                Case "Restkontingent"
                    '--------------------------------------------------
                    Eigenschaften_Fenster_Start_festlegen(0)
                    '--------------------------------------------------
                    Webservice_Restkontingent_ansteuern_und_auswerten()
                    '--------------------------------------------------
                    ProgressBar.Value = 50
                    '--------------------------------------------------
                    Ergebnisprotokoll_erstellen()
                    '--------------------------------------------------
                    Steuerungsdatei_erstellen()
                    '--------------------------------------------------
                    ProgressBar.Value = 90
                    '--------------------------------------------------
                Case "KontrolldateiPruefen"
                    '--------------------------------------------------
                    Eigenschaften_Fenster_Start_festlegen(0)
                    '--------------------------------------------------
                    Kontrolldatei_lesen()
                    '--------------------------------------------------
                    ProgressBar.Value = 20
                    '--------------------------------------------------
                    XML_Kontrolldatei_bearbeiten_und_speichern()
                    '--------------------------------------------------
                    ProgressBar.Value = 50
                    '--------------------------------------------------
                    Webservice_KontrolldateiPruefen_ansteuern_und_auswerten()
                    '--------------------------------------------------
                    ProgressBar.Value = 70
                    '--------------------------------------------------
                    Ergebnisprotokoll_erstellen()
                    '--------------------------------------------------
                    Steuerungsdatei_erstellen()
                    '--------------------------------------------------
                    ProgressBar.Value = 90
                    '--------------------------------------------------
                Case "Zusatzdatenerfassung"
                    '--------------------------------------------------
                    Eigenschaften_Fenster_Start_festlegen(0)
                    '--------------------------------------------------
                    Kontrolldatei_lesen()
                    '--------------------------------------------------
                    ProgressBar.Value = 20
                    '--------------------------------------------------
                    XML_Kontrolldatei_bearbeiten_und_speichern()
                    '--------------------------------------------------
                    ProgressBar.Value = 30
                    '--------------------------------------------------
                    Webservice_KontrolldateiPruefen_ansteuern_und_auswerten()
                    '--------------------------------------------------
                    ProgressBar.Value = 50
                    '--------------------------------------------------
                    Webservice_Zusatzdatenerfassung_ansteuern_und_auswerten()
                    '--------------------------------------------------
                    ProgressBar.Value = 70
                    '--------------------------------------------------
                    Ergebnisprotokoll_erstellen()
                    '--------------------------------------------------
                    Steuerungsdatei_erstellen()
                    '--------------------------------------------------
                    ProgressBar.Value = 90
                    '--------------------------------------------------
                Case "OffeneKontrolldateien"
                    '--------------------------------------------------
                    Eigenschaften_Fenster_Start_festlegen(0)
                    '--------------------------------------------------
                    Webservice_OffeneKontrolldateien_ansteuern_und_auswerten()
                    '--------------------------------------------------
                    ProgressBar.Value = 50
                    '--------------------------------------------------
                    Ergebnisprotokoll_erstellen()
                    '--------------------------------------------------
                    Steuerungsdatei_erstellen()
                    '--------------------------------------------------
                    ProgressBar.Value = 90
                    '--------------------------------------------------
            End Select
            '--------------------------------------------------
            If Variable_Steuerung.Ablauf_Fehler = True Then
                '--------------------------------------------------
                TabControl.SelectedIndex = 1
                '--------------------------------------------------
            End If
            '--------------------------------------------------
            ProgressBar.Value = 100
            '--------------------------------------------------
            If Variable_Steuerung.PDF_erzeugen = True Then
                '--------------------------------------------------
                If ListView.Items.Count > 0 Then
                    '--------------------------------------------------
                    Schaltflaeche_PDFErzeugen.Visible = False
                    PDF_erzeugen(Variable_Steuerung.PDF_oeffnen)
                    '--------------------------------------------------
                End If
                '--------------------------------------------------
            End If
            '--------------------------------------------------
            If Variable_Steuerung.Anwendung_beenden = True Then
                '--------------------------------------------------
                Me.Close()
                '--------------------------------------------------
            End If
            '--------------------------------------------------
            Windows.Forms.Cursor.Current = Cursors.Default
            '--------------------------------------------------
            ProgressBar.Visible = False
            '--------------------------------------------------
        Catch ex As Exception
            Fehlerfenster(ex)
        End Try
        '--------------------------------------------------
    End Sub
    '--------------------------------------------------
    ''' <summary>
    ''' In dieser Anweisung wird das Ergebnisprotokoll geschrieben.
    ''' </summary>
    Private Sub Ergebnisprotokoll_erstellen()
        '--------------------------------------------------
        If Variable_Steuerung.Sandbox = True Then
            '--------------------------------------------------
            Inhalt_Tabelle_Ergebnis += HTML_Tabelle_Zeile_Leerzeile()
            Inhalt_Tabelle_Ergebnis += HTML_Tabelle_Zeile_Leerzeile()
            Inhalt_Tabelle_Ergebnis += HTML_Tabelle_Ueberschrift_Rot("Informationen zur Sandbox:")
            Inhalt_Tabelle_Ergebnis += HTML_Tabelle_Zeile("Arbeitsverzeichnis:", Variable_Parameter.Arbeitsverzeichnis)
            Inhalt_Tabelle_Ergebnis += HTML_Tabelle_Zeile("Steuerungsdatei:", Variable_Parameter.Steuerungsdatei)
            Inhalt_Tabelle_Ergebnis += HTML_Tabelle_Zeile("ID-DIBt:", Variable_Parameter.Aussteller_ID_DIBT)
            Inhalt_Tabelle_Ergebnis += HTML_Tabelle_Zeile("PW-DIBt:", Variable_Parameter.Aussteller_PWD_DIBT)
            Inhalt_Tabelle_Ergebnis += HTML_Tabelle_Zeile_Leerzeile()
            '--------------------------------------------------
        End If
        '--------------------------------------------------
        HTML_Ergebnisdarstellung(HTML_Start() + HTML_Tabelle_Start() + Inhalt_Tabelle_Ergebnis + HTML_Tabelle_Ende() + HTML_Ende())
        '--------------------------------------------------
        Variable_Steuerung.Ergebnisprotokoll_Text = ErgebnisDIBt2025.Document.Body.InnerHtml
        '--------------------------------------------------
        Variable_Steuerung.Ergebnisprotokoll_Base64 = String_To_Base64(ErgebnisDIBt2025.Document.Body.InnerHtml)
        '--------------------------------------------------
    End Sub
    '--------------------------------------------------
    ''' <summary>
    ''' In dieser Anweisung wird die Fensterdarstellung festgelegt.
    ''' </summary>
    Private Sub Eigenschaften_Fenster_Start_festlegen(ByVal Art As Integer)
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            Select Case Art
                Case 0
                    '--------------------------------------------------
                    Me.Width = 677
                    Me.Height = 1024
                    Me.StartPosition = FormStartPosition.CenterScreen
                    '--------------------------------------------------
                    TabControl.Width = 638
                    TabControl.SelectedIndex = 1
                    '--------------------------------------------------
                    TabControl.Visible = True
                    ListView.Visible = False
                    ProgressBar.Visible = True
                    MenuStrip.Visible = True
                    StatusStrip.Visible = True
                    '--------------------------------------------------
                    Schaltflaeche_PDFErzeugen.Visible = False
                '--------------------------------------------------
                Case 1
                    '--------------------------------------------------
                    Me.Width = 850
                    Me.Height = 1024
                    Me.StartPosition = FormStartPosition.CenterScreen
                    '--------------------------------------------------
                    TabControl.Width = 638
                    TabControl.SelectedIndex = 0
                    '--------------------------------------------------
                    TabControl.Visible = True
                    ListView.Visible = True
                    ProgressBar.Visible = True
                    MenuStrip.Visible = True
                    StatusStrip.Visible = True
                    '--------------------------------------------------
            End Select
            '--------------------------------------------------
        Catch ex As Exception
            Fehlerfenster(ex)
        End Try
        '--------------------------------------------------
    End Sub
    '--------------------------------------------------
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
                Inhalt_Tabelle_Ergebnis += HTML_Tabelle_Zeile_Abschnitt("Rückmeldung DIBt Server - Datenregistratur")
                '--------------------------------------------------
                If Variable_Steuerung.Datenregistratur_Fehler_Anzahl = 0 Then
                    '--------------------------------------------------
                    Inhalt_Tabelle_Ergebnis += HTML_Tabelle_Ueberschrift("Ergebnisse:")
                    Inhalt_Tabelle_Ergebnis += HTML_Tabelle_Zeile("Registriernummer:", Variable_Steuerung.Registriernummer)
                    Inhalt_Tabelle_Ergebnis += HTML_Tabelle_Zeile("Restkontingent:", Variable_Steuerung.Restkontingent_Anzahl & " Stk.")
                    Inhalt_Tabelle_Ergebnis += HTML_Tabelle_Zeile_Leerzeile()
                    '--------------------------------------------------
                    If Variable_Steuerung.Datenregistratur_Pruefung_Datendatei > 0 Then
                        Inhalt_Tabelle_Ergebnis += HTML_Tabelle_Ueberschrift("Datendatei:")
                        Inhalt_Tabelle_Ergebnis += HTML_Tabelle_Zeile("", Variable_Steuerung.Datenregistratur_Pruefung_Datendatei_Bemerkungen)
                        Inhalt_Tabelle_Ergebnis += HTML_Tabelle_Zeile_Leerzeile()
                    End If
                    '--------------------------------------------------
                    Inhalt_Tabelle_Ergebnis += HTML_Tabelle_Ueberschrift("Mitteilung:")
                    Inhalt_Tabelle_Ergebnis += HTML_Tabelle_Zeile("", "Die Datenregistratur wurde ohne Fehler durchgeführt.")
                    Inhalt_Tabelle_Ergebnis += HTML_Tabelle_Zeile_Leerzeile()
                    '--------------------------------------------------
                Else
                    '--------------------------------------------------
                    Variable_Steuerung.Ablauf_Fehler = True
                    '--------------------------------------------------
                    Inhalt_Tabelle_Ergebnis += HTML_Tabelle_Ueberschrift("Fehlermeldungen:")
                    '--------------------------------------------------
                    Dim wertX As Integer
                    '--------------------------------------------------
                    For wertX = 1 To Variable_Steuerung.Datenregistratur_Fehler_Anzahl
                        '--------------------------------------------------
                        Inhalt_Tabelle_Ergebnis += HTML_Tabelle_Zeile_ID("ID " & Variable_Steuerung.Datenregistratur_Fehler_ID(wertX - 1), Variable_Steuerung.Datenregistratur_Fehler_Text(wertX - 1))
                        '--------------------------------------------------
                    Next
                    '--------------------------------------------------
                    Inhalt_Tabelle_Ergebnis += HTML_Tabelle_Zeile_Leerzeile()
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
                Inhalt_Tabelle_Ergebnis += HTML_Tabelle_Zeile_Abschnitt("Rückmeldung DIBt Server - Restkontingent")
                '--------------------------------------------------
                If Variable_Steuerung.Restkontingent_Fehlermeldung = "" Then
                    '--------------------------------------------------
                    Inhalt_Tabelle_Ergebnis += HTML_Tabelle_Ueberschrift("Ergebnis:")
                    Inhalt_Tabelle_Ergebnis += HTML_Tabelle_Zeile("Restkontingent:", Variable_Steuerung.Restkontingent_Anzahl & " Stk.")
                    Inhalt_Tabelle_Ergebnis += HTML_Tabelle_Zeile_Leerzeile()
                    '--------------------------------------------------
                Else
                    '--------------------------------------------------
                    Variable_Steuerung.Ablauf_Fehler = True
                    '--------------------------------------------------
                    Inhalt_Tabelle_Ergebnis += HTML_Tabelle_Ueberschrift("Fehlermeldung:")
                    '--------------------------------------------------
                    Inhalt_Tabelle_Ergebnis += HTML_Tabelle_Zeile("", Variable_Steuerung.Restkontingent_Fehlermeldung)
                    '--------------------------------------------------
                    Inhalt_Tabelle_Ergebnis += HTML_Tabelle_Zeile_Leerzeile()
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
                Inhalt_Tabelle_Ergebnis += HTML_Tabelle_Zeile_Abschnitt("Rückmeldung DIBt Server - Kontrolldatei Prüfen")
                '--------------------------------------------------
                If Webservice_KontrolldateiPruefen() = False Then
                    '--------------------------------------------------
                    Inhalt_Tabelle_Ergebnis += HTML_Tabelle_Ueberschrift("Mitteilung:")
                    Inhalt_Tabelle_Ergebnis += HTML_Tabelle_Zeile("", "Prüfung erfolgreich durchgeführt.")
                    Inhalt_Tabelle_Ergebnis += HTML_Tabelle_Zeile_Leerzeile()
                    '--------------------------------------------------
                    Variable_Steuerung.Entwurf = False
                    '--------------------------------------------------
                Else
                    '--------------------------------------------------
                    Variable_Steuerung.Ablauf_Fehler = True
                    '--------------------------------------------------
                    Inhalt_Tabelle_Ergebnis += HTML_Tabelle_Ueberschrift("Mitteilung:")
                    Inhalt_Tabelle_Ergebnis += HTML_Tabelle_Zeile("", "Prüfung nicht erfolgreich.")
                    '--------------------------------------------------
                    Inhalt_Tabelle_Ergebnis += HTML_Tabelle_Ueberschrift("Fehlermeldungen:")
                    '--------------------------------------------------
                    Dim WertX As Integer
                    '--------------------------------------------------
                    For WertX = 0 To (Variable_Steuerung.KontrolldateiPruefen_Fehler_Anzahl - 1)
                        '--------------------------------------------------
                        Inhalt_Tabelle_Ergebnis += HTML_Tabelle_Zeile_ID("ID " & Variable_Steuerung.KontrolldateiPruefen_Fehler_ID(WertX), Variable_Steuerung.KontrolldateiPruefen_Fehler_Kurztext(WertX))
                        Inhalt_Tabelle_Ergebnis += HTML_Tabelle_Zeile_ID("", Variable_Steuerung.KontrolldateiPruefen_Fehler_Langtext(WertX), 12)
                        '--------------------------------------------------
                    Next
                    '--------------------------------------------------
                    Inhalt_Tabelle_Ergebnis += HTML_Tabelle_Zeile_Leerzeile()
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
                Inhalt_Tabelle_Ergebnis += HTML_Tabelle_Zeile_Abschnitt("Rückmeldung DIBt Server - Kontrolldatei hochladen")
                '--------------------------------------------------
                If Variable_Steuerung.ZusatzdatenErfassung_Fehler_Anzahl = 0 Then
                    '--------------------------------------------------
                    Inhalt_Tabelle_Ergebnis += HTML_Tabelle_Ueberschrift("Mitteilung:")
                    Inhalt_Tabelle_Ergebnis += HTML_Tabelle_Zeile("", "Die Kontrolldatei wurde beim DIBt Server erfolgreich hochgeladen.")
                    Inhalt_Tabelle_Ergebnis += HTML_Tabelle_Zeile_Leerzeile()
                    '--------------------------------------------------
                    Variable_Steuerung.Entwurf = False
                    '--------------------------------------------------
                Else
                    '--------------------------------------------------
                    Variable_Steuerung.Ablauf_Fehler = True
                    '--------------------------------------------------
                    Inhalt_Tabelle_Ergebnis += HTML_Tabelle_Ueberschrift("Fehlermeldungen:")
                    Inhalt_Tabelle_Ergebnis += HTML_Tabelle_Zeile("", "Die Kontrolldatei wurde beim DIBt Server <strong><u>nicht</u></strong> erfolgreich hochgeladen.")
                    '--------------------------------------------------
                    Dim WertX As Integer
                    '--------------------------------------------------
                    For WertX = 1 To Variable_Steuerung.ZusatzdatenErfassung_Fehler_Anzahl
                        '--------------------------------------------------
                        Inhalt_Tabelle_Ergebnis += HTML_Tabelle_Zeile_ID("ID " & Variable_Steuerung.ZusatzdatenErfassung_Fehler_ID(WertX - 1), Variable_Steuerung.ZusatzdatenErfassung_Fehler_Text(WertX - 1))
                        '--------------------------------------------------
                    Next
                    '--------------------------------------------------
                    Inhalt_Tabelle_Ergebnis += HTML_Tabelle_Zeile_Leerzeile()
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
                Inhalt_Tabelle_Ergebnis += HTML_Tabelle_Zeile_Abschnitt("Rückmeldung DIBt Server - Offene Kontrolldateien")
                '--------------------------------------------------
                If Variable_Steuerung.OffeneKontrolldateien_Fehlermeldung = "" Then
                    '--------------------------------------------------
                    Inhalt_Tabelle_Ergebnis += HTML_Tabelle_Ueberschrift("Liste der offenen Ausweise:")
                    '--------------------------------------------------
                    Inhalt_Tabelle_Ergebnis += HTML_Tabelle_Zeile("Anzahl:", Variable_Steuerung.OffeneKontrolldateien_Anzahl & " Stk.")
                    '--------------------------------------------------
                    Dim WertX As Integer
                    '--------------------------------------------------
                    For WertX = 0 To (Variable_Steuerung.OffeneKontrolldateien_Anzahl - 1)
                        '--------------------------------------------------
                        Inhalt_Tabelle_Ergebnis += HTML_Tabelle_Zeile_ID(FormatDateTime(Variable_Steuerung.OffeneKontrolldateien_Ausweis_NummerErzeugtAm(WertX), DateFormat.ShortDate), Variable_Steuerung.OffeneKontrolldateien_Ausweis_Registriernummer(WertX))
                        '--------------------------------------------------
                    Next
                    '--------------------------------------------------
                    Inhalt_Tabelle_Ergebnis += HTML_Tabelle_Zeile_Leerzeile()
                    '--------------------------------------------------
                Else
                    '--------------------------------------------------
                    Variable_Steuerung.Ablauf_Fehler = True
                    '--------------------------------------------------
                    Inhalt_Tabelle_Ergebnis += HTML_Tabelle_Ueberschrift("Fehlermeldung:")
                    '--------------------------------------------------
                    Inhalt_Tabelle_Ergebnis += HTML_Tabelle_Zeile("", Variable_Steuerung.OffeneKontrolldateien_Fehlermeldung)
                    '--------------------------------------------------
                    Inhalt_Tabelle_Ergebnis += HTML_Tabelle_Zeile_Leerzeile()
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
            If Variable_Steuerung.Ablauf_Fehler = False Then
                '--------------------------------------------------
                Select Case Variable_Steuerung.Gesetzesgrundlage
                    Case "ENEV-2016"
                        '--------------------------------------------------
                        Select Case Variable_Steuerung.Gebaeudeart
                            Case "Wohngebäude", "Wohnteil gemischt genutztes Gebäude"
                                '--------------------------------------------------
                                Dim EnergieausweisWB2016 As New Klasse_Energieausweis_WB_2016
                                EnergieausweisWB2016.Energieausweis_erzeugen(True, PictureBox, ListView, ImageList)
                                '--------------------------------------------------
                            Case "Nichtwohngebäude", "Nichtwohnteil gemischt genutztes Gebäude"
                                '--------------------------------------------------
                                Dim EnergieausweisNWB2016 As New Klasse_Energieausweis_NWB_2016
                                EnergieausweisNWB2016.Energieausweis_erzeugen(True, PictureBox, ListView, ImageList)
                                '--------------------------------------------------
                        End Select
                        '--------------------------------------------------
                    Case "GEG-2020"
                        '--------------------------------------------------
                        Select Case Variable_Steuerung.Gebaeudeart
                            Case "Wohngebäude", "Wohnteil gemischt genutztes Gebäude"
                                '--------------------------------------------------
                                Dim EnergieausweisWB2020 As New Klasse_Energieausweis_WB_2020
                                EnergieausweisWB2020.Energieausweis_erzeugen(True, PictureBox, ListView, ImageList)
                                '--------------------------------------------------
                            Case "Nichtwohngebäude", "Nichtwohnteil gemischt genutztes Gebäude"
                                '--------------------------------------------------
                                Dim EnergieausweisNWB2020 As New Klasse_Energieausweis_NWB_2020
                                EnergieausweisNWB2020.Energieausweis_erzeugen(True, PictureBox, ListView, ImageList)
                                '--------------------------------------------------
                        End Select
                        '--------------------------------------------------
                    Case "GEG-2023"
                        '--------------------------------------------------
                        Select Case Variable_Steuerung.Gebaeudeart
                            Case "Wohngebäude", "Wohnteil gemischt genutztes Gebäude"
                                '--------------------------------------------------
                                Dim EnergieausweisWB2023 As New Klasse_Energieausweis_WB_2023
                                EnergieausweisWB2023.Energieausweis_erzeugen(True, PictureBox, ListView, ImageList)
                                '--------------------------------------------------
                            Case "Nichtwohngebäude", "Nichtwohnteil gemischt genutztes Gebäude"
                                '--------------------------------------------------
                                Dim EnergieausweisNWB2023 As New Klasse_Energieausweis_NWB_2023
                                EnergieausweisNWB2023.Energieausweis_erzeugen(True, PictureBox, ListView, ImageList)
                                '--------------------------------------------------
                        End Select
                        '--------------------------------------------------
                    Case "GEG-2024"
                        '--------------------------------------------------
                        Select Case Variable_Steuerung.Gebaeudeart
                            Case "Wohngebäude", "Wohnteil gemischt genutztes Gebäude"
                                '--------------------------------------------------
                                Dim EnergieausweisWB2024 As New Klasse_Energieausweis_WB_2024
                                EnergieausweisWB2024.Energieausweis_erzeugen(True, PictureBox, ListView, ImageList)
                                '--------------------------------------------------
                            Case "Nichtwohngebäude", "Nichtwohnteil gemischt genutztes Gebäude"
                                '--------------------------------------------------
                                Dim EnergieausweisNWB2024 As New Klasse_Energieausweis_NWB_2024
                                EnergieausweisNWB2024.Energieausweis_erzeugen(True, PictureBox, ListView, ImageList)
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
            If ListView.Items.Count >= 1 Then
                ListView.Items(0).Selected = True
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
#Region "Ergebnisse eintragen"
    '-------------------------------------------------------------------
    ''' <summary>
    ''' HTML Ergebnisdarstellung wird an das WebBrowser Steuerelement übergeben.
    ''' </summary>
    Private Sub HTML_Ergebnisdarstellung(ByVal HtmlText As String) 'ByVal WebBrowser As WebBrowser,
        '-------------------------------------------------------------------
        With ErgebnisDIBt2025
            '-------------------------------------------------------------------
            If IsNothing(.Url) OrElse .Url.AbsoluteUri <> "about:blank" Then
                .Navigate("about:blank")
                Application.DoEvents()
                '-------------------------------------------------------------------
                .Document.Body.InnerHtml = HtmlText
                '-------------------------------------------------------------------
            End If
            '-------------------------------------------------------------------
        End With
        '-------------------------------------------------------------------
    End Sub
    '-------------------------------------------------------------------
#End Region
    '--------------------------------------------------
#Region "Schaltflächen auswerten"
    '--------------------------------------------------
    ''' <summary>
    ''' Diese Schaltfläche beendet die Druckapplikation.
    ''' </summary>
    Private Sub Schaltflaeche_Beenden_Click(sender As Object, e As EventArgs) Handles Schaltflaeche_Beenden.Click
        Me.Close()
    End Sub
    '--------------------------------------------------
    ''' <summary>
    ''' In dieser Anweisung wird die PDF Datei geschrieben und geöffnet.
    ''' </summary>
    Private Sub Schaltflaeche_PDFErzeugen_Click(sender As Object, e As EventArgs) Handles Schaltflaeche_PDFErzeugen.Click
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            If ListView.Items.Count > 0 Then
                '--------------------------------------------------
                Schaltflaeche_PDFErzeugen.Visible = False
                PDF_erzeugen(True)
                '--------------------------------------------------
            Else
            End If
            '--------------------------------------------------
        Catch ex As Exception
            Fehlerfenster(ex)
        End Try
        '--------------------------------------------------
    End Sub
    '---------------------------------------
    '--------------------------------------------------
    ''' <summary>
    ''' In dieser Anweisung öffnet das Infofenster.
    ''' </summary>
    Private Sub Schaltflaeche_Info_Click(sender As Object, e As EventArgs) Handles Schaltflaeche_Info.Click
        Text_Haftung.SelectionStart = 0
        Text_Haftung.SelectionLength = 0
        Panel_Info.Visible = True
        Panel_Info_Schatten.Visible = True
    End Sub
    '--------------------------------------------------
    ''' <summary>
    ''' In dieser Anweisung schließt das Infofenster.
    ''' </summary>
    Private Sub Schaltflaeche_OK_Click(sender As Object, e As EventArgs) Handles Schaltflaeche_OK.Click
        Panel_Info.Visible = False
        Panel_Info_Schatten.Visible = False
    End Sub
    '---------------------------------------
#End Region
    '---------------------------------------
#Region "ListView auswerten"
    '---------------------------------------
    ''' <summary>
    ''' In dieser Anweisung wertet die Seitenvorschau aus.
    ''' </summary>
    Private Sub ListView_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListView.SelectedIndexChanged
        '---------------------------------------
        Try
            '---------------------------------------
            If ListView.SelectedItems IsNot Nothing AndAlso ListView.SelectedItems.Count > 0 Then
                '---------------------------------------
                For Each LV As Windows.Forms.ListViewItem In ListView.SelectedItems
                    '---------------------------------------
                    PictureBox.Image = String_to_Image(Variable_XML_Import.Bilddateien(LV.Index + 1))
                    '---------------------------------------
                    TabControl.SelectedIndex = 0
                    '---------------------------------------
                Next
                '---------------------------------------
            End If
            '---------------------------------------
        Catch ex As Exception
            Fehlerfenster(ex)
        End Try
        '---------------------------------------
    End Sub
    '---------------------------------------
#End Region
    '---------------------------------------
End Class