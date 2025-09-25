Module Modul_HTML
    '-------------------------------------------------------------------
#Region "Ergebnisse eintragen"
    '-------------------------------------------------------------------
    ''' <summary>
    ''' HTML Baustein für die Ergebnisdarstellung
    ''' </summary>
    Function HTML_Start() As String
        '-------------------------------------------------------------------
        HTML_Start = "<html><head><title></title></head><body>"
        '-------------------------------------------------------------------
    End Function
    '-------------------------------------------------------------------
    ''' <summary>
    ''' HTML Baustein für die Ergebnisdarstellung
    ''' </summary>
    Function HTML_Ende() As String
        '-------------------------------------------------------------------
        HTML_Ende = "</body></html>"
        '-------------------------------------------------------------------
    End Function
    '-------------------------------------------------------------------
    ''' <summary>
    ''' HTML Baustein für die Ergebnisdarstellung
    ''' </summary>
    Function HTML_Tabelle_Start() As String
        '-------------------------------------------------------------------
        HTML_Tabelle_Start = "<table align=""left"" border=""0"" cellpadding=""1"" cellspacing=""1"" width=""100%""><tbody>"
        '-------------------------------------------------------------------
    End Function
    '-------------------------------------------------------------------
    ''' <summary>
    ''' HTML Baustein für die Ergebnisdarstellung
    ''' </summary>
    Function HTML_Tabelle_Abschnitt(ByVal Bezeichnung As String) As String
        '-------------------------------------------------------------------
        Dim Zeichengroesse As String = "16"
        '-------------------------------------------------------------------
        HTML_Tabelle_Abschnitt = "<tr bgcolor=""silver""><td colspan=""4""><span style=""font-size: " & Zeichengroesse & "px;""><span style=""font-family: arial, helvetica, sans - serif; color:black;""><strong>" & Bezeichnung & "</strong></span></span></td></tr>"
        '-------------------------------------------------------------------
    End Function
    '-------------------------------------------------------------------
    ''' <summary>
    ''' HTML Baustein für die Ergebnisdarstellung
    ''' </summary>
    Function HTML_Tabelle_Zeile_Abschnitt(ByVal Bezeichnung As String) As String
        '-------------------------------------------------------------------
        Dim Zeichengroesse As String = "14"
        '-------------------------------------------------------------------
        HTML_Tabelle_Zeile_Abschnitt = "<tr><td colspan=""4""><span style=""font-size: " & Zeichengroesse & "px;""><span style=""font-family: arial, helvetica, sans - serif;""><strong><u>" & Bezeichnung & "</u></strong></span></span></td></tr>"
        '-------------------------------------------------------------------
    End Function
    '-------------------------------------------------------------------
    ''' <summary>
    ''' HTML Baustein für die Ergebnisdarstellung
    ''' </summary>
    Function HTML_Tabelle_Ueberschrift(ByVal Bezeichnung As String) As String
        '-------------------------------------------------------------------
        Dim Zeichengroesse As String = "14"
        '-------------------------------------------------------------------
        HTML_Tabelle_Ueberschrift = "<tr bgcolor=""silver""><td colspan=""4""><span style=""font-size: " & Zeichengroesse & "px;""><span style=""font-family: arial, helvetica, sans - serif; color:black;"">" & Bezeichnung & "</span></span></td></tr>"
        '-------------------------------------------------------------------
    End Function
    '-------------------------------------------------------------------
    ''' <summary>
    ''' HTML Baustein für die Ergebnisdarstellung
    ''' </summary>
    Function HTML_Tabelle_Ueberschrift_Rot(ByVal Bezeichnung As String) As String
        '-------------------------------------------------------------------
        Dim Zeichengroesse As String = "14"
        '-------------------------------------------------------------------
        HTML_Tabelle_Ueberschrift_Rot = "<tr bgcolor=""red""><td colspan=""4""><span style=""font-size: " & Zeichengroesse & "px;""><span style=""font-family: arial, helvetica, sans - serif; color:white;"">" & Bezeichnung & "</span></span></td></tr>"
        '-------------------------------------------------------------------
    End Function
    '-------------------------------------------------------------------
    ''' <summary>
    ''' HTML Baustein für die Ergebnisdarstellung
    ''' </summary>
    Function HTML_Tabelle_Ueberschrift_Gruen(ByVal Bezeichnung As String) As String
        '-------------------------------------------------------------------
        Dim Zeichengroesse As String = "14"
        '-------------------------------------------------------------------
        HTML_Tabelle_Ueberschrift_Gruen = "<tr bgcolor=""green""><td colspan=""4""><span style=""font-size: " & Zeichengroesse & "px;""><span style=""font-family: arial, helvetica, sans - serif; color:white;"">" & Bezeichnung & "</span></span></td></tr>"
        '-------------------------------------------------------------------
    End Function
    '-------------------------------------------------------------------
    ''' <summary>
    ''' HTML Baustein für die Ergebnisdarstellung
    ''' </summary>
    Function HTML_Tabelle_Zeile_Leerzeile() As String
        '-------------------------------------------------------------------
        Dim Zeichengroesse As String = "14"
        '-------------------------------------------------------------------
        HTML_Tabelle_Zeile_Leerzeile = "<tr><td colspan=""4""><span style=""font-size: " & Zeichengroesse & "px;""><span style=""font-family: arial, helvetica, sans - serif;""><strong> </strong></span></span></td></tr>"
        '-------------------------------------------------------------------
    End Function
    '-------------------------------------------------------------------
    ''' <summary>
    ''' HTML Baustein für die Ergebnisdarstellung
    ''' </summary>
    Function HTML_Tabelle_Zeile(ByVal Bezeichnung As String, ByVal Ergebnis As String) As String
        '-------------------------------------------------------------------
        Dim Zeichengroesse As String = "14"
        '-------------------------------------------------------------------
        HTML_Tabelle_Zeile = "<tr><td width=""20%""><span style=""font-size: " & Zeichengroesse & "px;""><span style=""font-family: arial, helvetica, sans - serif;"">" & Bezeichnung & "</span></span></td>
			<td width=""80%"" style=""text-align: Left;""><span style=""font-size: " & Zeichengroesse & "px;""><span style=""font-family: arial, helvetica, sans - serif;"">" & Ergebnis & "</span></span></td></tr>"
        '-------------------------------------------------------------------
    End Function
    '-------------------------------------------------------------------
    ''' <summary>
    ''' HTML Baustein für die Ergebnisdarstellung
    ''' </summary>
    Function HTML_Tabelle_Zeile_ID(ByVal Bezeichnung As String, ByVal Ergebnis As String) As String
        '-------------------------------------------------------------------
        Dim Zeichengroesse As String = "14"
        '-------------------------------------------------------------------
        HTML_Tabelle_Zeile_ID = "<tr><td width=""20%""><span style=""font-size: " & Zeichengroesse & "px;""><span style=""font-family: arial, helvetica, sans - serif;"">" & Bezeichnung & "</span></span></td>
			<td width=""80%"" style=""text-align: Left;""><span style=""font-size: " & Zeichengroesse & "px;""><span style=""font-family: arial, helvetica, sans - serif;"">" & Ergebnis & "</span></span></td></tr>"
        '-------------------------------------------------------------------
    End Function
    '-------------------------------------------------------------------
    ''' <summary>
    ''' HTML Baustein für die Ergebnisdarstellung
    ''' </summary>
    Function HTML_Tabelle_Zeile_ID(ByVal Bezeichnung As String, ByVal Ergebnis As String, ByVal Fontgroesse As Integer) As String
        '-------------------------------------------------------------------
        Dim Zeichengroesse As String = Fontgroesse.ToString
        '-------------------------------------------------------------------
        HTML_Tabelle_Zeile_ID = "<tr><td width=""20%""><span style=""font-size: " & Zeichengroesse & "px;""><span style=""font-family: arial, helvetica, sans - serif;"">" & Bezeichnung & "</span></span></td>
			<td width=""80%"" style=""text-align: Left;""><span style=""font-size: " & Zeichengroesse & "px;""><span style=""font-family: arial, helvetica, sans - serif;"">" & Ergebnis & "</span></span></td></tr>"
        '-------------------------------------------------------------------
    End Function
    '-------------------------------------------------------------------
    ''' <summary>
    ''' HTML Baustein für die Ergebnisdarstellung
    ''' </summary>
    Function HTML_Tabelle_Ende() As String
        '-------------------------------------------------------------------
        HTML_Tabelle_Ende = "</tbody></table>"
        '-------------------------------------------------------------------
    End Function
    '-------------------------------------------------------------------
#End Region
    '-------------------------------------------------------------------
End Module
