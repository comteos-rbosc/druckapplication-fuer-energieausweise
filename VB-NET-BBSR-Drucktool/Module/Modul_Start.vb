'--------------------------------------------------
#Region "Imports"
Imports System.IO
Imports System.Windows.Forms
#End Region
'--------------------------------------------------
Module Modul_Start
    '--------------------------------------------------
    ''' <summary>
    ''' In dieser Anweisung wird der Start mit oder ohne Oberflaeche durchgeführt.
    ''' </summary>
    Public Sub Main()
        '--------------------------------------------------
        AddHandler Application.ThreadException, AddressOf GeneralErrorHandler
        '--------------------------------------------------
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        '--------------------------------------------------
        Variable_Steuerung.Ablauf_Fehler = False
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            Variable_Steuerung.Ablauf_Fehler = Webservice_Internettest()
            Variable_Steuerung.Internetverbindung_Fehler = Webservice_Internettest()
            '--------------------------------------------------
            If Variable_Steuerung.Internetverbindung_Fehler = True Then
                '--------------------------------------------------
                MsgBox("Eine Verbindung zum Internet konnte nicht hergestellt werden. Die Anwendung wird beendet.",, "Fehlermeldung")
                '--------------------------------------------------
            Else
                '--------------------------------------------------
                Startparameter_auslesen()
                '--------------------------------------------------
                Steuerungsdatei_lesen()
                '--------------------------------------------------
                If Variable_Steuerung.Anwendung_minimieren = True Then
                    Modul_Ablauf_ohne_Oberflaeche.Main()
                Else
                    Application.Run(New Fenster_Start)
                End If
                '--------------------------------------------------
            End If
        Catch ex As Exception
            Fehlerfenster(ex)
        End Try
        '--------------------------------------------------
    End Sub
    '-------------------------------------------------------
End Module
