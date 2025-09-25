#Region "Imports"
Imports System.Windows.Forms
#End Region
'-------------------------------------------------------
Module Modul_Fehlerbehandlung
    '-------------------------------------------------------
#Region "zentrale Fehlerbehandlung"
    '-------------------------------------------------------
    ''' <summary>
    ''' In dieser Anweisung wird die zentrale Fehlerbehandlung definiert und danach die Anwendung gestartet.
    ''' </summary>
    Public Sub Main()
        AddHandler Application.ThreadException, AddressOf GeneralErrorHandler
        '-------------------------------------------------------
        Application.Run(New Fenster_Start)
        '-------------------------------------------------------
    End Sub
    '-------------------------------------------------------
    ''' <summary>
    ''' In dieser Anweisung wird ein Fehlerfenster mit wichtigen Informationen ausgegeben, auch wenn in der Funktion keine Fehlerbehandlung enthalten ist.
    ''' </summary>
    Sub GeneralErrorHandler(ByVal sender As Object, ByVal e As System.Threading.ThreadExceptionEventArgs)
        '-------------------------------------------------------
        Dim sError As String = "Unerwarteter Anwendungsfehler: " & vbCrLf & e.Exception.Message.ToString & vbCrLf & vbCrLf & "Source: " & e.Exception.StackTrace.ToString & vbCrLf & vbCrLf & "Bitte kontaktieren Sie den Entwickler ihrer Software."
        '-------------------------------------------------------
        MsgBox(sError, MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly, "Unerwarteter Anwendungsfehler:")
        '-------------------------------------------------------
        'Application.Exit()
        '-------------------------------------------------------
    End Sub
    '-------------------------------------------------------
    ''' <summary>
    ''' In dieser Anweisung wird ein Fehlerfenster mit wichtigen Informationen ausgegeben.
    ''' </summary>
    Sub Fehlerfenster(ByVal ex As Exception)
        '-------------------------------------------------------
        Dim sError As String = "Unerwarteter Anwendungsfehler: " & vbCrLf & ex.Message.ToString & vbCrLf & vbCrLf & "Source: " & ex.StackTrace.ToString & vbCrLf & vbCrLf & "Bitte kontaktieren Sie den Entwickler ihrer Software."
        '-------------------------------------------------------
        MsgBox(sError, MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly, "Unerwarteter Anwendungsfehler:")
        '-------------------------------------------------------
    End Sub
    '-------------------------------------------------------
#End Region
    '-------------------------------------------------------
#Region "Speicher freigeben"
    '-------------------------------------------------------
    ''' <summary>
    ''' In dieser Anweisung wird der unbenutzte Speicher wieder freigegeben.
    ''' </summary>
    Sub Speicher_freigeben()
        '-------------------------------------------------------
        GC.Collect()
        GC.WaitForPendingFinalizers()
        '-------------------------------------------------------
    End Sub
    '-------------------------------------------------------
#End Region
    '-------------------------------------------------------
End Module
