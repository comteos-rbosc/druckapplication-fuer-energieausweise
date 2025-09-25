#Region "Imports"
Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms
Imports BBSR_Energieausweis.DynaPDF
#End Region
'--------------------------------------------------
Module Modul_Image
    '-------------------------------------------------------
#Region "Image erstellen - String"
    '--------------------------------------------------
    ''' <summary>
    ''' Mit dieser Anweisung wird ein Text auf ein Image geschrieben.
    ''' </summary>
    Sub Image_Text_schreiben(ByVal ImageBox As Graphics, ByVal Ausrichtung As Ausrichtung, ByVal Text As String, ByVal X1 As Single, ByVal X2 As Single, ByVal Y1 As Single, ByVal Y2 As Single, ByVal Font As Font, ByVal Farbe As SolidBrush)
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            Dim PointText As PointF
            '-------------------------------------------------------
            Dim Abstand_Oben As Integer = 20
            '-------------------------------------------------------
            Dim Format_Text As New StringFormat
            '-------------------------------------------------------
            If Text <> "" Then
                '-------------------------------------------------------
                Select Case Ausrichtung
                    Case Ausrichtung.EinzeiligLinks
                        '-------------------------------------------------------
                        Format_Text.Alignment = StringAlignment.Near
                        Format_Text.LineAlignment = StringAlignment.Center
                        '-------------------------------------------------------
                        Dim Rechteck As New RectangleF(X1, Y1, (X2 - X1), (Y2 - Y1))
                        '-------------------------------------------------------
                        Text = TextAbschneiden(Text, (X2 - X1), Font)
                        '-------------------------------------------------------
                        ImageBox.DrawString(Text, Font, Farbe, Rechteck, Format_Text)
                    '-------------------------------------------------------
                    Case Ausrichtung.EinzeiligMitte
                        '-------------------------------------------------------
                        Format_Text.Alignment = StringAlignment.Center
                        Format_Text.LineAlignment = StringAlignment.Center
                        '-------------------------------------------------------
                        Dim Rechteck As New RectangleF(X1, Y1, (X2 - X1), (Y2 - Y1))
                        '-------------------------------------------------------
                        Text = TextAbschneiden(Text, (X2 - X1), Font)
                        '-------------------------------------------------------
                        ImageBox.DrawString(Text, Font, Farbe, Rechteck, Format_Text)
                    '-------------------------------------------------------
                    Case Ausrichtung.EinzeiligRechts
                        '-------------------------------------------------------
                        Format_Text.Alignment = StringAlignment.Far
                        Format_Text.LineAlignment = StringAlignment.Center
                        '-------------------------------------------------------
                        Dim Rechteck As New RectangleF(X1, Y1, (X2 - X1), (Y2 - Y1))
                        '-------------------------------------------------------
                        Text = TextAbschneiden(Text, (X2 - X1), Font)
                        '-------------------------------------------------------
                        ImageBox.DrawString(Text, Font, Farbe, Rechteck, Format_Text)
                    '-------------------------------------------------------
                    '-------------------------------------------------------
                    Case Ausrichtung.MehrzeiligLinksMitte
                        '-------------------------------------------------------
                        Format_Text.Alignment = StringAlignment.Near
                        Format_Text.LineAlignment = StringAlignment.Center
                        '-------------------------------------------------------
                        Dim Rechteck As New RectangleF(X1, Y1, (X2 - X1), (Y2 - Y1))
                        '-------------------------------------------------------
                        ImageBox.DrawString(Text, Font, Farbe, Rechteck, Format_Text)
                    '-------------------------------------------------------
                    Case Ausrichtung.MehrzeiligMitteMitte
                        '-------------------------------------------------------
                        Format_Text.Alignment = StringAlignment.Center
                        Format_Text.LineAlignment = StringAlignment.Center
                        '-------------------------------------------------------
                        Dim Rechteck As New RectangleF(X1, Y1, (X2 - X1), (Y2 - Y1))
                        '-------------------------------------------------------
                        ImageBox.DrawString(Text, Font, Farbe, Rechteck, Format_Text)
                    '-------------------------------------------------------
                    Case Ausrichtung.MehrzeiligRechtsMitte
                        '-------------------------------------------------------
                        Format_Text.Alignment = StringAlignment.Far
                        Format_Text.LineAlignment = StringAlignment.Center
                        '-------------------------------------------------------
                        Dim Rechteck As New RectangleF(X1, Y1, (X2 - X1), (Y2 - Y1))
                        '-------------------------------------------------------
                        ImageBox.DrawString(Text, Font, Farbe, Rechteck, Format_Text)
                    '-------------------------------------------------------
                    '-------------------------------------------------------
                    Case Ausrichtung.MehrzeiligLinksOben
                        '-------------------------------------------------------
                        Format_Text.Alignment = StringAlignment.Near
                        Format_Text.LineAlignment = StringAlignment.Near
                        '-------------------------------------------------------
                        Dim Rechteck As New RectangleF(X1, Y1 + Abstand_Oben, (X2 - X1), (Y2 - Y1 - Abstand_Oben))
                        '-------------------------------------------------------
                        ImageBox.DrawString(Text, Font, Farbe, Rechteck, Format_Text)
                    '-------------------------------------------------------
                    Case Ausrichtung.MehrzeiligMitteOben
                        '-------------------------------------------------------
                        Format_Text.Alignment = StringAlignment.Center
                        Format_Text.LineAlignment = StringAlignment.Near
                        '-------------------------------------------------------
                        Dim Rechteck As New RectangleF(X1, Y1 + Abstand_Oben, (X2 - X1), (Y2 - Y1 - Abstand_Oben))
                        '-------------------------------------------------------
                        ImageBox.DrawString(Text, Font, Farbe, Rechteck, Format_Text)
                    '-------------------------------------------------------
                    Case Ausrichtung.MehrzeiligRechtsOben
                        '-------------------------------------------------------
                        Format_Text.Alignment = StringAlignment.Far
                        Format_Text.LineAlignment = StringAlignment.Near
                        '-------------------------------------------------------
                        Dim Rechteck As New RectangleF(X1, Y1 + Abstand_Oben, (X2 - X1), (Y2 - Y1 - Abstand_Oben))
                        '-------------------------------------------------------
                        ImageBox.DrawString(Text, Font, Farbe, Rechteck, Format_Text)
                    '-------------------------------------------------------
                    Case Ausrichtung.Links_kein_Abschneiden
                        '-------------------------------------------------------
                        PointText.X = X1
                        PointText.Y = ((Y1 + ((Y2 - Y1) / 2))) - (ImageBox.MeasureString(Text, Font).Height * 0.5)
                        ImageBox.DrawString(Text, Font, Farbe, PointText, Format_Text)
                        '-------------------------------------------------------
                    Case Ausrichtung.Rechts_kein_Abschneiden
                        '-------------------------------------------------------
                        PointText.X = (X2) - (ImageBox.MeasureString(Text, Font).Width)
                        PointText.Y = ((Y1 + ((Y2 - Y1) / 2))) - (ImageBox.MeasureString(Text, Font).Height * 0.5)
                        ImageBox.DrawString(Text, Font, Farbe, PointText, Format_Text)
                    '-------------------------------------------------------
                    Case Ausrichtung.Mitte_kein_Abschneiden
                        '-------------------------------------------------------
                        PointText.X = ((X1 + ((X2 - X1) / 2))) - (ImageBox.MeasureString(Text, Font).Width * 0.5)
                        PointText.Y = ((Y1 + ((Y2 - Y1) / 2))) - (ImageBox.MeasureString(Text, Font).Height * 0.5)
                        ImageBox.DrawString(Text, Font, Farbe, PointText, Format_Text)
                        '-------------------------------------------------------
                        'Case Ausrichtung.Mehrzeilig
                        '    '-------------------------------------------------------
                        '    Dim Rechteck As New RectangleF(X1, Y1, (X2 - X1), (Y2 - Y1))
                        '    '-------------------------------------------------------
                        '    ImageBox.DrawString(Text, Font, Farbe, Rechteck, Format_Text)
                        '    '-------------------------------------------------------
                End Select
                '-------------------------------------------------------
            End If
            '--------------------------------------------------
        Catch ex As Exception
            Fehlerfenster(ex)
        End Try
        '-------------------------------------------------------
    End Sub
    '--------------------------------------------------
    ''' <summary>
    ''' Mit dieser Anweisung wird ein Text in einem Winkel auf ein Image geschrieben.
    ''' </summary>
    Sub Image_Text_im_Winkel_schreiben(ByVal ImageBox As Graphics, ByVal Winkel As Integer, ByVal Text As String, ByVal X1 As Single, ByVal X2 As Single, ByVal Y1 As Single, ByVal Y2 As Single, ByVal Font As Font, ByVal Farbe As SolidBrush)
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            If Text <> "" Then
                '-------------------------------------------------------
                Dim StringFormat As StringFormat = StringFormat.GenericTypographic
                Dim Alte_Position As Drawing2D.Matrix = ImageBox.Transform
                StringFormat.LineAlignment = StringAlignment.Center
                StringFormat.Alignment = StringAlignment.Center
                ImageBox.ResetTransform()
                ImageBox.TranslateTransform(X2 - X1, Y2 - Y1)
                ImageBox.RotateTransform(Winkel)
                ImageBox.DrawString(Text, Font, Farbe, 0, 0, StringFormat)
                ImageBox.Transform = Alte_Position
                '-------------------------------------------------------
            End If
            '--------------------------------------------------
        Catch ex As Exception
            Fehlerfenster(ex)
        End Try
        '-------------------------------------------------------
    End Sub
    '--------------------------------------------------
    ''' <summary>
    ''' Mit dieser Anweisung wird ein Text auf ein Image geschrieben, wenn der Text zu lang ist, wird er abgeschnitten.
    ''' </summary>
    Private Function TextAbschneiden(ByVal text As String, ByVal maxWidth As Integer, ByVal font As Font) As String
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            Dim ErgebnisText As String = text
            Dim TextBreite As Integer = TextRenderer.MeasureText(text, font).Width
            '--------------------------------------------------
            If TextBreite > maxWidth Then
                '--------------------------------------------------
                Dim ellipsisWidth As Integer = TextRenderer.MeasureText("...", font).Width
                Dim availableWidth As Integer = maxWidth - ellipsisWidth
                Dim currentIndex As Integer = text.Length - 1
                '--------------------------------------------------
                While TextBreite > availableWidth AndAlso currentIndex >= 0
                    '--------------------------------------------------
                    ErgebnisText = ErgebnisText.Substring(0, currentIndex)
                    TextBreite = TextRenderer.MeasureText(ErgebnisText, font).Width
                    currentIndex -= 1
                    '--------------------------------------------------
                End While
                '--------------------------------------------------
                ErgebnisText &= "..."
                '--------------------------------------------------
            End If
            '--------------------------------------------------
            Return ErgebnisText
            '--------------------------------------------------
        Catch ex As Exception
            Fehlerfenster(ex)
        End Try
        '--------------------------------------------------
    End Function
    '--------------------------------------------------
#End Region
    '--------------------------------------------------
#Region "Image erstellen - Zahl"
    '--------------------------------------------------
    ''' <summary>
    ''' Mit dieser Anweisung wird eine Zahl auf ein Image geschrieben.
    ''' </summary>
    Sub Image_Zahl_schreiben(ByVal ImageBox As Graphics, ByVal Ausrichtung As Ausrichtung, ByVal Zahl As Single, ByVal Kommastelle As Integer, ByVal Einheit As String, ByVal X1 As Single, ByVal X2 As Single, ByVal Y1 As Single, ByVal Y2 As Single, ByVal Font As Font, ByVal Farbe As SolidBrush)
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            Dim PointText As PointF
            '-------------------------------------------------------
            If Einheit <> "" Then
                Einheit = " " & Einheit
            End If
            '-------------------------------------------------------
            Dim Format_Text As New StringFormat
            '-------------------------------------------------------
            Select Case Ausrichtung
                Case Ausrichtung.EinzeiligLinks
                    '-------------------------------------------------------
                    PointText.X = X1
                    PointText.Y = ((Y1 + ((Y2 - Y1) / 2))) - (ImageBox.MeasureString(FormatNumber(Zahl, Kommastelle), Font).Height * 0.5)
                    ImageBox.DrawString(FormatNumber(Zahl, Kommastelle) & Einheit, Font, Farbe, PointText)
                '-------------------------------------------------------
                Case Ausrichtung.EinzeiligMitte
                    '-------------------------------------------------------
                    PointText.X = ((X1 + ((X2 - X1) / 2))) - (ImageBox.MeasureString(FormatNumber(Zahl, Kommastelle) & Einheit, Font).Width * 0.5)
                    PointText.Y = ((Y1 + ((Y2 - Y1) / 2))) - (ImageBox.MeasureString(FormatNumber(Zahl, Kommastelle), Font).Height * 0.5)
                    ImageBox.DrawString(FormatNumber(Zahl, Kommastelle) & Einheit, Font, Farbe, PointText)
                '-------------------------------------------------------
                Case Ausrichtung.EinzeiligRechts
                    '-------------------------------------------------------
                    PointText.X = (X2) - (ImageBox.MeasureString(FormatNumber(Zahl, Kommastelle) & Einheit, Font).Width)
                    PointText.Y = ((Y1 + ((Y2 - Y1) / 2))) - (ImageBox.MeasureString(FormatNumber(Zahl, Kommastelle), Font).Height * 0.5)
                    ImageBox.DrawString(FormatNumber(Zahl, Kommastelle) & Einheit, Font, Farbe, PointText)
                    '-------------------------------------------------------
                Case Ausrichtung.Links_kein_Abschneiden
                    '-------------------------------------------------------
                    PointText.X = X1
                    PointText.Y = ((Y1 + ((Y2 - Y1) / 2))) - (ImageBox.MeasureString(FormatNumber(Zahl, Kommastelle) & Einheit, Font).Height * 0.5)
                    ImageBox.DrawString(FormatNumber(Zahl, Kommastelle) & Einheit, Font, Farbe, PointText, Format_Text)
                        '-------------------------------------------------------
                Case Ausrichtung.Rechts_kein_Abschneiden
                    '-------------------------------------------------------
                    PointText.X = (X2) - (ImageBox.MeasureString(FormatNumber(Zahl, Kommastelle) & Einheit, Font).Width)
                    PointText.Y = ((Y1 + ((Y2 - Y1) / 2))) - (ImageBox.MeasureString(FormatNumber(Zahl, Kommastelle) & Einheit, Font).Height * 0.5)
                    ImageBox.DrawString(FormatNumber(Zahl, Kommastelle) & Einheit, Font, Farbe, PointText, Format_Text)
                    '-------------------------------------------------------
                Case Ausrichtung.Mitte_kein_Abschneiden
                    '-------------------------------------------------------
                    PointText.X = ((X1 + ((X2 - X1) / 2))) - (ImageBox.MeasureString(FormatNumber(Zahl, Kommastelle) & Einheit, Font).Width * 0.5)
                    PointText.Y = ((Y1 + ((Y2 - Y1) / 2))) - (ImageBox.MeasureString(FormatNumber(Zahl, Kommastelle) & Einheit, Font).Height * 0.5)
                    ImageBox.DrawString(FormatNumber(Zahl, Kommastelle) & Einheit, Font, Farbe, PointText, Format_Text)
            End Select
            '-------------------------------------------------------
        Catch ex As Exception
            Fehlerfenster(ex)
        End Try
        '-------------------------------------------------------
    End Sub
    '--------------------------------------------------
    ''' <summary>
    ''' Mit dieser Anweisung wird eine Zahl inkl. Attribut auf ein Image geschrieben.
    ''' </summary>
    Sub Image_Zahl_schreiben(ByVal ImageBox As Graphics, ByVal Ausrichtung As Ausrichtung, ByVal Zahl As Single, ByVal Kommastelle As Integer, ByVal Attribut As String, ByVal Einheit As String, ByVal X1 As Single, ByVal X2 As Single, ByVal Y1 As Single, ByVal Y2 As Single, ByVal Font As Font, ByVal Farbe As SolidBrush)
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            Dim PointText As PointF
            '-------------------------------------------------------
            If Einheit <> "" Then
                Einheit = " " & Einheit
            End If
            '-------------------------------------------------------
            Select Case Ausrichtung
                Case Ausrichtung.EinzeiligLinks
                    '-------------------------------------------------------
                    PointText.X = X1
                    PointText.Y = ((Y1 + ((Y2 - Y1) / 2))) - (ImageBox.MeasureString(FormatNumber(Zahl, Kommastelle), Font).Height * 0.5)
                    ImageBox.DrawString(Attribut & FormatNumber(Zahl, Kommastelle) & Einheit, Font, Farbe, PointText)
                '-------------------------------------------------------
                Case Ausrichtung.EinzeiligMitte
                    '-------------------------------------------------------
                    PointText.X = ((X1 + ((X2 - X1) / 2))) - (ImageBox.MeasureString(Attribut & FormatNumber(Zahl, Kommastelle) & Einheit, Font).Width * 0.5)
                    PointText.Y = ((Y1 + ((Y2 - Y1) / 2))) - (ImageBox.MeasureString(FormatNumber(Zahl, Kommastelle), Font).Height * 0.5)
                    ImageBox.DrawString(Attribut & FormatNumber(Zahl, Kommastelle) & Einheit, Font, Farbe, PointText)
                '-------------------------------------------------------
                Case Ausrichtung.EinzeiligRechts
                    '-------------------------------------------------------
                    PointText.X = (X2) - (ImageBox.MeasureString(Attribut & FormatNumber(Zahl, Kommastelle) & Einheit, Font).Width)
                    PointText.Y = ((Y1 + ((Y2 - Y1) / 2))) - (ImageBox.MeasureString(FormatNumber(Zahl, Kommastelle), Font).Height * 0.5)
                    ImageBox.DrawString(Attribut & FormatNumber(Zahl, Kommastelle) & Einheit, Font, Farbe, PointText)
                    '-------------------------------------------------------
            End Select
            '-------------------------------------------------------
        Catch ex As Exception
            Fehlerfenster(ex)
        End Try
        '-------------------------------------------------------
    End Sub
    '--------------------------------------------------
#End Region
    '--------------------------------------------------
#Region "Image erstellen - Datum"
    '--------------------------------------------------
    ''' <summary>
    ''' Mit dieser Anweisung wird eine Datum auf ein Image geschrieben.
    ''' </summary>
    Sub Image_Datum_schreiben(ByVal ImageBox As Graphics, ByVal Ausrichtung As Ausrichtung, ByVal Datum As Date, ByVal X1 As Single, ByVal X2 As Single, ByVal Y1 As Single, ByVal Y2 As Single, ByVal Font As Font, ByVal Farbe As SolidBrush)
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            Dim PointText As PointF
            '-------------------------------------------------------
            Select Case Ausrichtung
                Case Ausrichtung.EinzeiligLinks
                    '-------------------------------------------------------
                    PointText.X = X1
                    PointText.Y = ((Y1 + ((Y2 - Y1) / 2))) - (ImageBox.MeasureString(FormatDateTime(Datum, DateFormat.ShortDate).ToString, Font).Height * 0.5)
                    ImageBox.DrawString(FormatDateTime(Datum, DateFormat.ShortDate).ToString, Font, Farbe, PointText)
                '-------------------------------------------------------
                Case Ausrichtung.EinzeiligMitte
                    '-------------------------------------------------------
                    PointText.X = ((X1 + ((X2 - X1) / 2))) - (ImageBox.MeasureString(FormatDateTime(Datum, DateFormat.ShortDate).ToString, Font).Width * 0.5)
                    PointText.Y = ((Y1 + ((Y2 - Y1) / 2))) - (ImageBox.MeasureString(FormatDateTime(Datum, DateFormat.ShortDate).ToString, Font).Height * 0.5)
                    ImageBox.DrawString(FormatDateTime(Datum, DateFormat.ShortDate).ToString, Font, Farbe, PointText)
                '-------------------------------------------------------
                Case Ausrichtung.EinzeiligRechts
                    '-------------------------------------------------------
                    PointText.X = (X2) - (ImageBox.MeasureString(FormatDateTime(Datum, DateFormat.ShortDate).ToString, Font).Width)
                    PointText.Y = ((Y1 + ((Y2 - Y1) / 2))) - (ImageBox.MeasureString(FormatDateTime(Datum, DateFormat.ShortDate).ToString, Font).Height * 0.5)
                    ImageBox.DrawString(FormatDateTime(Datum, DateFormat.ShortDate).ToString, Font, Farbe, PointText)
                    '-------------------------------------------------------
            End Select
            '-------------------------------------------------------
        Catch ex As Exception
            Fehlerfenster(ex)
        End Try
        '-------------------------------------------------------
    End Sub
    '--------------------------------------------------
#End Region
    '--------------------------------------------------
#Region "Image erstellen - Auswahlfeld"
    '--------------------------------------------------
    ''' <summary>
    ''' Mit dieser Anweisung wird ein Auswahlfeld auf ein Image geschrieben.
    ''' </summary>
    Sub Image_Auswahl_schreiben(ByVal ImageBox As Graphics, ByVal Auswahl As Boolean, ByVal X1 As Single, ByVal X2 As Single, ByVal Y1 As Single, ByVal Y2 As Single)
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            Dim Stift_Auswahl As New Pen(Color.Black, 5)
            Dim Ueberhang As Integer = 10
            Dim Auswahl_Variante As Integer = 1
            '-------------------------------------------------------
            If Auswahl = True Then
                '-------------------------------------------------------
                Select Case Auswahl_Variante
                    Case 0
                        '-------------------------------------------------------
                        ImageBox.DrawLine(Stift_Auswahl, X1, Y1, X2, Y2)
                        ImageBox.DrawLine(Stift_Auswahl, X2, Y1, X1, Y2)
                    '-------------------------------------------------------
                    Case 1
                        '-------------------------------------------------------
                        ImageBox.DrawLine(Stift_Auswahl, X1 - Ueberhang, Y1 - Ueberhang, X2 + Ueberhang, Y2 + Ueberhang)
                        ImageBox.DrawLine(Stift_Auswahl, X2 + Ueberhang, Y1 - Ueberhang, X1 - Ueberhang, Y2 + Ueberhang)
                        '-------------------------------------------------------
                End Select
                '-------------------------------------------------------
            End If
            '-------------------------------------------------------
        Catch ex As Exception
            Fehlerfenster(ex)
        End Try
        '-------------------------------------------------------
    End Sub
    '--------------------------------------------------
#End Region
    '--------------------------------------------------
#Region "Image erstellen - Grafiken"
    '--------------------------------------------------
    ''' <summary>
    ''' Mit dieser Anweisung wird eine Projektgrafik auf ein Image geschrieben.
    ''' </summary>
    Sub Image_Grafik_Projekt_schreiben(ByVal ImageBox As Graphics, ByVal Bilddatei As String, ByVal Zeichenflaeche_X As Integer, ByVal Zeichenflaeche_Y As Integer, ByVal Zeichenflaeche_Widht As Integer, ByVal Zeichenflaeche_Height As Integer)
        '--------------------------------------------------
        Try
            '-------------------------------------------------------
            If Bilddatei <> "" Then
                If File.Exists(Bilddatei) Then
                    '-------------------------------------------------------
                    Dim Bild As Image = Image.FromFile(Bilddatei)
                    '-------------------------------------------------------
                    If Zeichenflaeche_Widht / Zeichenflaeche_Height < Bild.Width / Bild.Height Then
                        ImageBox.DrawImage(Bild, Zeichenflaeche_X + 0, Zeichenflaeche_Y + CInt((Zeichenflaeche_Height - (Zeichenflaeche_Widht / Bild.Width * Bild.Height)) / 2), Zeichenflaeche_Widht, CInt(Zeichenflaeche_Widht / Bild.Width * Bild.Height))
                    Else
                        ImageBox.DrawImage(Bild, Zeichenflaeche_X + CInt((Zeichenflaeche_Widht - (Zeichenflaeche_Height / Bild.Height * Bild.Width)) / 2), Zeichenflaeche_Y + 0, CInt(Zeichenflaeche_Height / Bild.Height * Bild.Width), Zeichenflaeche_Height)
                    End If
                    '-------------------------------------------------------
                Else
                    '-------------------------------------------------------
                    MsgBox("Die angegebene Bilddatei (" & Bilddatei & ") konnte nicht gefunden werden.")
                    '-------------------------------------------------------
                End If
            End If
            '--------------------------------------------------
        Catch ex As Exception
            Fehlerfenster(ex)
        End Try
        '-------------------------------------------------------
    End Sub
    '--------------------------------------------------
    ''' <summary>
    ''' Mit dieser Anweisung wird eine Unterschriftgrafik auf ein Image geschrieben.
    ''' </summary>
    Sub Image_Grafik_Unterschrift_schreiben(ByVal ImageBox As Graphics, ByVal Bilddatei As String, ByVal Zeichenflaeche_X As Integer, ByVal Zeichenflaeche_Y As Integer, ByVal Zeichenflaeche_Widht As Integer, ByVal Zeichenflaeche_Height As Integer)
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            If Bilddatei <> "" Then
                If File.Exists(Bilddatei) Then
                    '-------------------------------------------------------
                    Dim Bild As Image = Image.FromFile(Bilddatei)
                    '-------------------------------------------------------
                    If Zeichenflaeche_Widht / Zeichenflaeche_Height < Bild.Width / Bild.Height Then
                        ImageBox.DrawImage(Bild, Zeichenflaeche_X + 0, Zeichenflaeche_Y + CInt((Zeichenflaeche_Height - (Zeichenflaeche_Widht / Bild.Width * Bild.Height)) / 2), Zeichenflaeche_Widht, CInt(Zeichenflaeche_Widht / Bild.Width * Bild.Height))
                    Else
                        ImageBox.DrawImage(Bild, Zeichenflaeche_X + CInt((Zeichenflaeche_Widht - (Zeichenflaeche_Height / Bild.Height * Bild.Width)) / 2), Zeichenflaeche_Y + 0, CInt(Zeichenflaeche_Height / Bild.Height * Bild.Width), Zeichenflaeche_Height)
                    End If
                    '-------------------------------------------------------
                Else
                    '-------------------------------------------------------
                    MsgBox("Die angegebene Bilddatei (" & Bilddatei & ") konnte nicht gefunden werden.")
                    '-------------------------------------------------------
                End If
            End If
            '--------------------------------------------------
        Catch ex As Exception
            Fehlerfenster(ex)
        End Try
        '-------------------------------------------------------
    End Sub
    '--------------------------------------------------
    ''' <summary>
    ''' Mit dieser Anweisung wird eine Ausstellergrafik auf ein Image geschrieben.
    ''' </summary>
    Sub Image_Grafik_Aussteller_schreiben(ByVal ImageBox As Graphics, ByVal Bilddatei As String, ByVal Aussteller_Bezeichnung As String, ByVal Ausstellervorname As String, ByVal Ausstellername As String, ByVal Aussteller_Strasse_Nr As String, ByVal Aussteller_PLZ As String, ByVal Aussteller_Ort As String, ByVal Zeichenflaeche_X As Integer, ByVal Zeichenflaeche_Y As Integer, ByVal Zeichenflaeche_Widht As Integer, ByVal Zeichenflaeche_Height As Integer)
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            Dim Ausstellertext As String = ""
            If Aussteller_Bezeichnung <> "" Then Ausstellertext &= Aussteller_Bezeichnung & vbCrLf
            If Ausstellervorname <> "" Then Ausstellertext &= Ausstellervorname & " "
            Ausstellertext &= Ausstellername & vbCrLf
            Ausstellertext &= Aussteller_Strasse_Nr & vbCrLf
            Ausstellertext &= Aussteller_PLZ & " " & Aussteller_Ort
            '--------------------------------------------------
            Image_Text_schreiben(ImageBox, Ausrichtung.EinzeiligLinks, Ausstellertext, Zeichenflaeche_X, Zeichenflaeche_X + Zeichenflaeche_Widht, Zeichenflaeche_Y, Zeichenflaeche_Y + Zeichenflaeche_Height, Font_Schriftgroesse_40, Font_Druckfarbe)
            '--------------------------------------------------
            Zeichenflaeche_Widht = Zeichenflaeche_Widht / 4
            Zeichenflaeche_X = Zeichenflaeche_X + (Zeichenflaeche_Widht * 3)
            '--------------------------------------------------
            If Bilddatei <> "" Then
                If File.Exists(Bilddatei) Then
                    '-------------------------------------------------------
                    Dim Bild As Image = Image.FromFile(Bilddatei)
                    '-------------------------------------------------------
                    If Zeichenflaeche_Widht / Zeichenflaeche_Height < Bild.Width / Bild.Height Then
                        ImageBox.DrawImage(Bild, Zeichenflaeche_X + 0, Zeichenflaeche_Y + CInt((Zeichenflaeche_Height - (Zeichenflaeche_Widht / Bild.Width * Bild.Height)) / 2), Zeichenflaeche_Widht, CInt(Zeichenflaeche_Widht / Bild.Width * Bild.Height))
                    Else
                        ImageBox.DrawImage(Bild, Zeichenflaeche_X + CInt((Zeichenflaeche_Widht - (Zeichenflaeche_Height / Bild.Height * Bild.Width)) / 2), Zeichenflaeche_Y + 0, CInt(Zeichenflaeche_Height / Bild.Height * Bild.Width), Zeichenflaeche_Height)
                    End If
                    '-------------------------------------------------------
                Else
                    '-------------------------------------------------------
                    MsgBox("Die angegebene Bilddatei (" & Bilddatei & ") konnte nicht gefunden werden.")
                    '-------------------------------------------------------
                End If
            End If
            '--------------------------------------------------
        Catch ex As Exception
            Fehlerfenster(ex)
        End Try
        '-------------------------------------------------------
    End Sub
    '--------------------------------------------------
    ''' <summary>
    ''' Mit dieser Anweisung wird eine Grafik auf ein Image geschrieben.
    ''' </summary>
    Sub Image_Grafik_schreiben(ByVal ImageBox As Graphics, ByVal Bild As Image, ByVal Rechteck As Rectangle)
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            ImageBox.DrawImage(Bild, Rechteck)
            '-------------------------------------------------------
        Catch ex As Exception
            Fehlerfenster(ex)
        End Try
        '-------------------------------------------------------
    End Sub
    '--------------------------------------------------
    ''' <summary>
    ''' Mit dieser Anweisung wird eine Linie auf ein Image geschrieben.
    ''' </summary>
    Sub Image_Linie(ByVal ImageBox As Graphics, ByVal X1 As Single, ByVal X2 As Single, ByVal Y1 As Single, ByVal Y2 As Single)
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            Dim Stift As New Pen(Color.Black, 5)
            '-------------------------------------------------------
            ImageBox.DrawLine(Stift, X1, Y1, X2, Y2)
            '--------------------------------------------------
        Catch ex As Exception
            Fehlerfenster(ex)
        End Try
        '-------------------------------------------------------
    End Sub
    '--------------------------------------------------
    ''' <summary>
    ''' Mit dieser Anweisung wird ein Rechteck auf ein Image geschrieben.
    ''' </summary>
    Sub Image_Rechteck(ByVal ImageBox As Graphics, ByVal X1 As Single, ByVal X2 As Single, ByVal Y1 As Single, ByVal Y2 As Single, ByVal Farbe As Color)
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            Dim Fuellung As New SolidBrush(Farbe)
            '-------------------------------------------------------
            ImageBox.FillRectangle(Fuellung, X1, Y1, X2 - X1, Y2 - Y1)
            '--------------------------------------------------
        Catch ex As Exception
            Fehlerfenster(ex)
        End Try
        '-------------------------------------------------------
    End Sub
    '--------------------------------------------------
    ''' <summary>
    ''' Mit dieser Anweisung wird der Energiebalken für den Wohnbau inkl. der Pfeile auf ein Image geschrieben.
    ''' </summary>
    Sub Image_Pfeil_WB(ByVal ImageBox As Graphics, ByVal Lage As Lage, ByVal Wert As Decimal, ByVal Kommastelle As Integer, ByVal Beschriftung As String, ByVal X1 As Single, ByVal X2 As Single, ByVal Y1 As Single, ByVal Y2 As Single)
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            Dim X_Wert As Single = Berechnung_XWert_WB(Wert, X1, X2)
            Dim Y_Wert As Single
            Dim Energieklasse As String = Berechnung_Energieklasse_WB(Wert)
            '--------------------------------------------------
            Select Case Lage
                Case Lage.Oben
                    '--------------------------------------------------
                    ImageBox.DrawImage(Image.FromFile(Variable_Parameter.Arbeitsverzeichnis & "/Image/pfeil_runter.png"), X_Wert - 65, Y1 - 200, 130, 200)
                    '--------------------------------------------------
                    Select Case Wert
                        Case 0 To 130
                            Image_Text_schreiben(ImageBox, Ausrichtung.Links_kein_Abschneiden, Beschriftung, X_Wert + 100, X_Wert + 100, Y1 - 200, Y1 - 200, Font_Bold_Schriftgroesse_60, Font_Druckfarbe)
                            Image_Text_schreiben(ImageBox, Ausrichtung.Links_kein_Abschneiden, FormatNumber(Wert, Kommastelle) & " kWh/(m²·a)", X_Wert + 100, X_Wert + 100, Y1 - 100, Y1 - 100, Font_Bold_Schriftgroesse_60, Font_Druckfarbe)
                        Case > 130
                            Image_Text_schreiben(ImageBox, Ausrichtung.Rechts_kein_Abschneiden, Beschriftung, X_Wert - 100, X_Wert - 100, Y1 - 200, Y1 - 200, Font_Bold_Schriftgroesse_60, Font_Druckfarbe)
                            Image_Text_schreiben(ImageBox, Ausrichtung.Rechts_kein_Abschneiden, FormatNumber(Wert, Kommastelle) & " kWh/(m²·a)", X_Wert - 100, X_Wert - 100, Y1 - 100, Y1 - 100, Font_Bold_Schriftgroesse_60, Font_Druckfarbe)
                    End Select
                    '--------------------------------------------------
                    'Energieklassen
                    '-------------------------------------------------------
                    Y_Wert = Y1 + 65
                    '-------------------------------------------------------
                    X_Wert = Berechnung_XWert_WB(15 + 2, X1, X2)
                    If Energieklasse = "A+" Then
                        Image_Text_schreiben(ImageBox, Ausrichtung.EinzeiligMitte, "A+", X_Wert - 100, X_Wert + 100, Y_Wert, Y_Wert, Font_Bold_Schriftgroesse_70, Font_Druckfarbe)
                    Else
                        Image_Text_schreiben(ImageBox, Ausrichtung.EinzeiligMitte, "A+", X_Wert - 100, X_Wert + 100, Y_Wert, Y_Wert, Font_Schriftgroesse_60, Font_Druckfarbe)
                    End If
                    '-------------------------------------------------------
                    X_Wert = Berechnung_XWert_WB(40 - 0.5, X1, X2)
                    If Energieklasse = "A" Then
                        Image_Text_schreiben(ImageBox, Ausrichtung.EinzeiligMitte, "A", X_Wert - 100, X_Wert + 100, Y_Wert, Y_Wert, Font_Bold_Schriftgroesse_70, Font_Druckfarbe)
                    Else
                        Image_Text_schreiben(ImageBox, Ausrichtung.EinzeiligMitte, "A", X_Wert - 100, X_Wert + 100, Y_Wert, Y_Wert, Font_Schriftgroesse_60, Font_Druckfarbe)
                    End If
                    '-------------------------------------------------------
                    X_Wert = Berechnung_XWert_WB(62.5, X1, X2)
                    If Energieklasse = "B" Then
                        Image_Text_schreiben(ImageBox, Ausrichtung.EinzeiligMitte, "B", X_Wert - 100, X_Wert + 100, Y_Wert, Y_Wert, Font_Bold_Schriftgroesse_70, Font_Druckfarbe)
                    Else
                        Image_Text_schreiben(ImageBox, Ausrichtung.EinzeiligMitte, "B", X_Wert - 100, X_Wert + 100, Y_Wert, Y_Wert, Font_Schriftgroesse_60, Font_Druckfarbe)
                    End If
                    '-------------------------------------------------------
                    X_Wert = Berechnung_XWert_WB(87.5 - 1, X1, X2)
                    If Energieklasse = "C" Then
                        Image_Text_schreiben(ImageBox, Ausrichtung.EinzeiligMitte, "C", X_Wert - 100, X_Wert + 100, Y_Wert, Y_Wert, Font_Bold_Schriftgroesse_70, Font_Druckfarbe)
                    Else
                        Image_Text_schreiben(ImageBox, Ausrichtung.EinzeiligMitte, "C", X_Wert - 100, X_Wert + 100, Y_Wert, Y_Wert, Font_Schriftgroesse_60, Font_Druckfarbe)
                    End If
                    '-------------------------------------------------------
                    X_Wert = Berechnung_XWert_WB(115 - 1, X1, X2)
                    If Energieklasse = "D" Then
                        Image_Text_schreiben(ImageBox, Ausrichtung.EinzeiligMitte, "D", X_Wert - 100, X_Wert + 100, Y_Wert, Y_Wert, Font_Bold_Schriftgroesse_70, Font_Druckfarbe)
                    Else
                        Image_Text_schreiben(ImageBox, Ausrichtung.EinzeiligMitte, "D", X_Wert - 100, X_Wert + 100, Y_Wert, Y_Wert, Font_Schriftgroesse_60, Font_Druckfarbe)
                    End If
                    '-------------------------------------------------------
                    X_Wert = Berechnung_XWert_WB(145 - 0.5, X1, X2)
                    If Energieklasse = "E" Then
                        Image_Text_schreiben(ImageBox, Ausrichtung.EinzeiligMitte, "E", X_Wert - 100, X_Wert + 100, Y_Wert, Y_Wert, Font_Bold_Schriftgroesse_70, Font_Druckfarbe)
                    Else
                        Image_Text_schreiben(ImageBox, Ausrichtung.EinzeiligMitte, "E", X_Wert - 100, X_Wert + 100, Y_Wert, Y_Wert, Font_Schriftgroesse_60, Font_Druckfarbe)
                    End If
                    '-------------------------------------------------------
                    X_Wert = Berechnung_XWert_WB(180 + 0.5, X1, X2)
                    If Energieklasse = "F" Then
                        Image_Text_schreiben(ImageBox, Ausrichtung.EinzeiligMitte, "F", X_Wert - 100, X_Wert + 100, Y_Wert, Y_Wert, Font_Bold_Schriftgroesse_70, Font_Druckfarbe)
                    Else
                        Image_Text_schreiben(ImageBox, Ausrichtung.EinzeiligMitte, "F", X_Wert - 100, X_Wert + 100, Y_Wert, Y_Wert, Font_Schriftgroesse_60, Font_Druckfarbe)
                    End If
                    '-------------------------------------------------------
                    X_Wert = Berechnung_XWert_WB(225 + 0, X1, X2)
                    If Energieklasse = "G" Then
                        Image_Text_schreiben(ImageBox, Ausrichtung.EinzeiligMitte, "G", X_Wert - 100, X_Wert + 100, Y_Wert, Y_Wert, Font_Bold_Schriftgroesse_70, Font_Druckfarbe)
                    Else
                        Image_Text_schreiben(ImageBox, Ausrichtung.EinzeiligMitte, "G", X_Wert - 100, X_Wert + 100, Y_Wert, Y_Wert, Font_Schriftgroesse_60, Font_Druckfarbe)
                    End If
                    '-------------------------------------------------------
                    X_Wert = Berechnung_XWert_WB(262.5 + 0, X1, X2)
                    If Energieklasse = "H" Then
                        Image_Text_schreiben(ImageBox, Ausrichtung.EinzeiligMitte, "H", X_Wert - 100, X_Wert + 100, Y_Wert, Y_Wert, Font_Bold_Schriftgroesse_70, Font_Druckfarbe)
                    Else
                        Image_Text_schreiben(ImageBox, Ausrichtung.EinzeiligMitte, "H", X_Wert - 100, X_Wert + 100, Y_Wert, Y_Wert, Font_Schriftgroesse_60, Font_Druckfarbe)
                    End If
                '-------------------------------------------------------
                Case Lage.Unten
                    '--------------------------------------------------
                    ImageBox.DrawImage(Image.FromFile(Variable_Parameter.Arbeitsverzeichnis & "/Image/pfeil_hoch.png"), X_Wert - 65, Y2, 130, 200)
                    '--------------------------------------------------
                    Select Case Wert
                        Case 0 To 130
                            Image_Text_schreiben(ImageBox, Ausrichtung.Links_kein_Abschneiden, Beschriftung, X_Wert + 100, X_Wert + 100, Y2 + 100, Y2 + 100, Font_Bold_Schriftgroesse_60, Font_Druckfarbe)
                            Image_Text_schreiben(ImageBox, Ausrichtung.Links_kein_Abschneiden, FormatNumber(Wert, Kommastelle) & " kWh/(m²·a)", X_Wert + 100, X_Wert + 100, Y2 + 200, Y2 + 200, Font_Bold_Schriftgroesse_60, Font_Druckfarbe)
                        Case > 130
                            Image_Text_schreiben(ImageBox, Ausrichtung.Rechts_kein_Abschneiden, Beschriftung, X_Wert - 100, X_Wert - 100, Y2 + 100, Y2 + 100, Font_Bold_Schriftgroesse_60, Font_Druckfarbe)
                            Image_Text_schreiben(ImageBox, Ausrichtung.Rechts_kein_Abschneiden, FormatNumber(Wert, Kommastelle) & " kWh/(m²·a)", X_Wert - 100, X_Wert - 100, Y2 + 200, Y2 + 200, Font_Bold_Schriftgroesse_60, Font_Druckfarbe)
                    End Select
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
    ''' Mit dieser Anweisung wird der Energiebalken für den Nichtwohnbau inkl. der Pfeile auf ein Image geschrieben.
    ''' </summary>
    Sub Image_Pfeil_NWB(ByVal ImageBox As Graphics, ByVal Lage As Lage, ByVal Wert As Decimal, ByVal Kommastelle As Integer, ByVal Vergleichswert As Decimal, ByVal Faktor As Decimal, ByVal Beschriftung As String, ByVal Beschriftung_Zeile_2 As String, ByVal X1 As Single, ByVal X2 As Single, ByVal Y1 As Single, ByVal Y2 As Single)
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            Dim X_Wert As Single = Berechnung_Skala_XWert_NWB(ImageBox, Wert, Vergleichswert, Faktor, X1, X2, Y1, Y2)
            '--------------------------------------------------
            Select Case Lage
                Case Lage.Oben
                    '--------------------------------------------------
                    ImageBox.DrawImage(Image.FromFile(Variable_Parameter.Arbeitsverzeichnis & "/Image/pfeil_runter.png"), X_Wert - 65, Y1 - 200, 130, 200)
                    '--------------------------------------------------
                    Select Case (X_Wert - X1)
                        Case <= ((X2 - X1) / 2)
                            '--------------------------------------------------
                            If Beschriftung_Zeile_2 = "" Then
                                '--------------------------------------------------
                                Image_Text_schreiben(ImageBox, Ausrichtung.Links_kein_Abschneiden, Beschriftung, X_Wert + 100, X_Wert + 100, Y1 - 200, Y1 - 200, Font_Bold_Schriftgroesse_60, Font_Druckfarbe)
                                Image_Text_schreiben(ImageBox, Ausrichtung.Links_kein_Abschneiden, FormatNumber(Wert, Kommastelle) & " kWh/(m²·a)", X_Wert + 100, X_Wert + 100, Y1 - 100, Y1 - 100, Font_Bold_Schriftgroesse_60, Font_Druckfarbe)
                                '--------------------------------------------------
                            Else
                                '--------------------------------------------------
                                Image_Text_schreiben(ImageBox, Ausrichtung.Links_kein_Abschneiden, Beschriftung, X_Wert + 100, X_Wert + 100, Y1 - 270, Y1 - 270, Font_Bold_Schriftgroesse_60, Font_Druckfarbe)
                                Image_Text_schreiben(ImageBox, Ausrichtung.Links_kein_Abschneiden, Beschriftung_Zeile_2, X_Wert + 100, X_Wert + 100, Y1 - 170, Y1 - 170, Font_Schriftgroesse_60, Font_Druckfarbe)
                                Image_Text_schreiben(ImageBox, Ausrichtung.Links_kein_Abschneiden, FormatNumber(Wert, Kommastelle) & " kWh/(m²·a)", X_Wert + 100, X_Wert + 100, Y1 - 70, Y1 - 70, Font_Bold_Schriftgroesse_60, Font_Druckfarbe)
                                '--------------------------------------------------
                            End If
                        '--------------------------------------------------
                        Case > ((X2 - X1) / 2)
                            '--------------------------------------------------
                            If Beschriftung_Zeile_2 = "" Then
                                '--------------------------------------------------
                                Image_Text_schreiben(ImageBox, Ausrichtung.Rechts_kein_Abschneiden, Beschriftung, X_Wert - 100, X_Wert - 100, Y1 - 200, Y1 - 200, Font_Bold_Schriftgroesse_60, Font_Druckfarbe)
                                Image_Text_schreiben(ImageBox, Ausrichtung.Rechts_kein_Abschneiden, FormatNumber(Wert, Kommastelle) & " kWh/(m²·a)", X_Wert - 100, X_Wert - 100, Y1 - 100, Y1 - 100, Font_Bold_Schriftgroesse_60, Font_Druckfarbe)
                                '--------------------------------------------------
                            Else
                                '--------------------------------------------------
                                Image_Text_schreiben(ImageBox, Ausrichtung.Rechts_kein_Abschneiden, Beschriftung, X_Wert - 100, X_Wert - 100, Y1 - 270, Y1 - 270, Font_Bold_Schriftgroesse_60, Font_Druckfarbe)
                                Image_Text_schreiben(ImageBox, Ausrichtung.Rechts_kein_Abschneiden, Beschriftung_Zeile_2, X_Wert - 100, X_Wert - 100, Y1 - 170, Y1 - 170, Font_Schriftgroesse_60, Font_Druckfarbe)
                                Image_Text_schreiben(ImageBox, Ausrichtung.Rechts_kein_Abschneiden, FormatNumber(Wert, Kommastelle) & " kWh/(m²·a)", X_Wert - 100, X_Wert - 100, Y1 - 70, Y1 - 70, Font_Bold_Schriftgroesse_60, Font_Druckfarbe)
                                '--------------------------------------------------
                            End If
                            '--------------------------------------------------
                    End Select
                '--------------------------------------------------
                Case Lage.Unten
                    '--------------------------------------------------
                    ImageBox.DrawImage(Image.FromFile(Variable_Parameter.Arbeitsverzeichnis & "/Image/pfeil_hoch.png"), X_Wert - 65, Y2, 130, 200)
                    '--------------------------------------------------
                    Select Case Wert
                        Case 0 To 500
                            Image_Text_schreiben(ImageBox, Ausrichtung.Links_kein_Abschneiden, Beschriftung, X_Wert + 100, X_Wert + 100, Y2 + 100, Y2 + 100, Font_Bold_Schriftgroesse_60, Font_Druckfarbe)
                            Image_Text_schreiben(ImageBox, Ausrichtung.Links_kein_Abschneiden, FormatNumber(Wert, Kommastelle) & " kWh/(m²·a)", X_Wert + 100, X_Wert + 100, Y2 + 200, Y2 + 200, Font_Bold_Schriftgroesse_60, Font_Druckfarbe)
                        Case > 500
                            Image_Text_schreiben(ImageBox, Ausrichtung.Rechts_kein_Abschneiden, Beschriftung, X_Wert - 100, X_Wert - 100, Y2 + 100, Y2 + 100, Font_Bold_Schriftgroesse_60, Font_Druckfarbe)
                            Image_Text_schreiben(ImageBox, Ausrichtung.Rechts_kein_Abschneiden, FormatNumber(Wert, Kommastelle) & " kWh/(m²·a)", X_Wert - 100, X_Wert - 100, Y2 + 200, Y2 + 200, Font_Bold_Schriftgroesse_60, Font_Druckfarbe)
                    End Select
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
    ''' Mit dieser Anweisung wird der X-Wert für den Energiebalken berechnet.
    ''' </summary>
    Function Berechnung_XWert_WB(ByVal Wert As Decimal, ByVal X1 As Single, ByVal X2 As Single)
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            Dim WertX As Single = 0
            Dim Breite As Single = X2 - X1
            '--------------------------------------------------
            Dim X_Werte(9) As Integer
            '--------------------------------------------------
            X_Werte(0) = 0
            X_Werte(1) = 320
            X_Werte(2) = 571
            X_Werte(3) = 851
            X_Werte(4) = 1154
            X_Werte(5) = 1537
            X_Werte(6) = 1952
            X_Werte(7) = 2396
            X_Werte(8) = 2887
            X_Werte(9) = 3084
            '--------------------------------------------------
            Select Case Wert
                Case 0 To 30
                    WertX = X_Werte(0) + ((X_Werte(1) - X_Werte(0)) / 30 * (Wert - 0))
                Case 30 To 50
                    WertX = X_Werte(1) + ((X_Werte(2) - X_Werte(1)) / 20 * (Wert - 30))
                Case 50 To 75
                    WertX = X_Werte(2) + ((X_Werte(3) - X_Werte(2)) / 25 * (Wert - 50))
                Case 75 To 100
                    WertX = X_Werte(3) + ((X_Werte(4) - X_Werte(3)) / 25 * (Wert - 75))
                Case 100 To 130
                    WertX = X_Werte(4) + ((X_Werte(5) - X_Werte(4)) / 30 * (Wert - 100))
                Case 130 To 160
                    WertX = X_Werte(5) + ((X_Werte(6) - X_Werte(5)) / 30 * (Wert - 130))
                Case 160 To 200
                    WertX = X_Werte(6) + ((X_Werte(7) - X_Werte(6)) / 40 * (Wert - 160))
                Case 200 To 250
                    WertX = X_Werte(7) + ((X_Werte(8) - X_Werte(7)) / 50 * (Wert - 200))
                Case > 250
                    WertX = X_Werte(8) + ((X_Werte(9) - X_Werte(8)) / 25 * (Wert - 250))
                    If WertX > X_Werte(9) Then WertX = X_Werte(9)
            End Select
            '--------------------------------------------------
            Berechnung_XWert_WB = X1 + WertX
            '--------------------------------------------------
        Catch ex As Exception
            Fehlerfenster(ex)
        End Try
        '--------------------------------------------------
    End Function
    '--------------------------------------------------
    ''' <summary>
    ''' Mit dieser Anweisung wird die Energieklasse berechnet.
    ''' </summary>
    Function Berechnung_Energieklasse_WB(ByVal Wert As Decimal) As String
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            Select Case Wert
                Case 0 To 30
                    Berechnung_Energieklasse_WB = "A+"
                Case 30 To 50
                    Berechnung_Energieklasse_WB = "A"
                Case 50 To 75
                    Berechnung_Energieklasse_WB = "B"
                Case 75 To 100
                    Berechnung_Energieklasse_WB = "C"
                Case 100 To 130
                    Berechnung_Energieklasse_WB = "D"
                Case 130 To 160
                    Berechnung_Energieklasse_WB = "E"
                Case 160 To 200
                    Berechnung_Energieklasse_WB = "F"
                Case 200 To 250
                    Berechnung_Energieklasse_WB = "G"
                Case > 250
                    Berechnung_Energieklasse_WB = "H"
            End Select
            '--------------------------------------------------
        Catch ex As Exception
            Fehlerfenster(ex)
        End Try
        '--------------------------------------------------
    End Function
    '--------------------------------------------------
    ''' <summary>
    ''' Mit dieser Anweisung wird die Skalenteilung für den Nichtwohnbau berechnet.
    ''' </summary>
    Function Berechnung_Skala_XWert_NWB(ByVal ImageBox As Graphics, ByVal Wert As Decimal, ByVal Vergleichswert As Decimal, ByVal Faktor As Decimal, ByVal X1 As Single, ByVal X2 As Single, ByVal Y1 As Single, ByVal Y2 As Single) As Single
        '--------------------------------------------------
        Berechnung_Skala_XWert_NWB = 0
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            Dim WertX_Offset_Nullpunkt As Single = 70
            Dim WertX_Offset_Maxpunkt As Single = 120
            Dim WertX As Single = 0
            Dim Schrittweite As Single = 5
            Dim Gesamtbreite As Single = X2 - X1
            Dim Wert_Breite As Single = (X2 - WertX_Offset_Maxpunkt) - (X1 + WertX_Offset_Nullpunkt)
            Dim Wert_Maxpunkt As Integer
            '--------------------------------------------------
            Wert_Maxpunkt = CInt(Math.Round(Vergleichswert * Faktor / 10, 0) * 10)
            '--------------------------------------------------
            Select Case Wert_Maxpunkt
                Case 10 To 70
                    Schrittweite = 5
                Case 80 To 100
                    Schrittweite = 10
                Case 110 To 230
                    Schrittweite = 20
                Case 240 To 350
                    Schrittweite = 30
                Case 360 To 590
                    Schrittweite = 50
                Case 600 To 1050
                    Schrittweite = 100
                Case 1060 To 1900
                    Schrittweite = 200
                Case 1910 To 4760
                    Schrittweite = 500
                Case 4770 To 9520
                    Schrittweite = 1000
                Case 9530 To 10000
                    Schrittweite = 2000
                Case > 10000
                    Schrittweite = 3000
            End Select
            '--------------------------------------------------
            Image_Zahl_schreiben(ImageBox, Ausrichtung.EinzeiligMitte, 0, 0, "", X1 + WertX_Offset_Nullpunkt - 50, X1 + WertX_Offset_Nullpunkt + 50, Y1, Y2, Font_Skala, Font_Druckfarbe)
            '--------------------------------------------------
            Image_Zahl_schreiben(ImageBox, Ausrichtung.EinzeiligMitte, Wert_Maxpunkt, 0, "≥", "", X2 - WertX_Offset_Maxpunkt - 50, X2 - WertX_Offset_Maxpunkt + 50, Y1, Y2, Font_Skala, Font_Druckfarbe)
            '--------------------------------------------------
            Dim Gesamtwert As Single = 0
            Dim Anzahl_Schritte As Integer = 1
            Dim Anzahl_Durchlauf As Integer
            '--------------------------------------------------
            For Anzahl_Durchlauf = 0 To 20
                '--------------------------------------------------
                Gesamtwert = Anzahl_Schritte * Schrittweite
                Anzahl_Schritte += 1
                '--------------------------------------------------
                WertX = X1 + WertX_Offset_Nullpunkt + (Wert_Breite / Wert_Maxpunkt * Gesamtwert)
                '--------------------------------------------------
                If Gesamtwert > (Wert_Maxpunkt * 0.9) Then
                    Exit For
                End If
                '--------------------------------------------------
                Image_Zahl_schreiben(ImageBox, Ausrichtung.EinzeiligMitte, Gesamtwert, 0, "", WertX - 50, WertX + 50, Y1, Y2, Font_Skala, Font_Druckfarbe)
                '--------------------------------------------------
            Next
            '--------------------------------------------------
            WertX = Wert_Breite / Wert_Maxpunkt * Wert
            '--------------------------------------------------
            Berechnung_Skala_XWert_NWB = X1 + WertX_Offset_Nullpunkt + WertX
            '--------------------------------------------------
            If Berechnung_Skala_XWert_NWB > (X2 - WertX_Offset_Maxpunkt) Then
                Berechnung_Skala_XWert_NWB = X2 - WertX_Offset_Maxpunkt
            End If
            '--------------------------------------------------
        Catch ex As Exception
            Fehlerfenster(ex)
        End Try
        '--------------------------------------------------
    End Function
    '--------------------------------------------------
    ''' <summary>
    ''' Mit dieser Anweisung wird die Energieaufteilung des Balkendiagramms für den Nichtwohnbau berechnet.
    ''' </summary>
    Sub Image_Diagram_NWB_Energieaufteilung(ByVal ImageBox As Graphics, ByVal X1 As Single, ByVal X2 As Single, ByVal Y1 As Single, ByVal Y2 As Single)
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            Dim Wert_X_Breite_Gesamt As Decimal
            Dim Wert_X_Nutzenergie As Decimal
            Dim Wert_X_Endenergie As Decimal
            Dim Wert_X_Primaerenergie As Decimal
            '--------------------------------------------------
            Wert_X_Breite_Gesamt = X2 - X1
            Wert_X_Nutzenergie = X1 + Wert_X_Breite_Gesamt * (464 / 1916)
            Wert_X_Endenergie = X1 + Wert_X_Breite_Gesamt * (1057 / 1916)
            Wert_X_Primaerenergie = X1 + Wert_X_Breite_Gesamt * (1642 / 1916)
            '--------------------------------------------------
            Dim Breite_Balken As Decimal = 200
            '--------------------------------------------------
            Dim Maximalwert As Integer = Berechnung_des_maximal_Energiebedarf()
            Dim Skalenteilung As Integer = Berechnung_der_Skalenteilung_Energiebedarf(Maximalwert)
            '--------------------------------------------------
            Dim Wert_Y_Max As Decimal = Y1 + 50
            Dim Wert_Y_Null As Decimal = Y2
            Dim Wert_Y_Anzahl_Teilung As Decimal = 5
            Dim Wert_Y_Hoehe_Gesamt As Decimal = Wert_Y_Null - Wert_Y_Max
            Dim Wert_Y_Hoehe_Teilung As Decimal = Wert_Y_Hoehe_Gesamt / Wert_Y_Anzahl_Teilung
            '--------------------------------------------------
            Dim Wert_Y_oben As Decimal = 0
            Dim Wert_Y_unten As Decimal = 0
            '--------------------------------------------------
            Dim Anzahl_Zeilen As Integer
            '--------------------------------------------------
            For Anzahl_Zeilen = 1 To 5
                '-------------------------------------------------------
                Image_Linie(ImageBox, X1, X2, Wert_Y_Null - (Wert_Y_Hoehe_Teilung * Anzahl_Zeilen), Wert_Y_Null - (Wert_Y_Hoehe_Teilung * Anzahl_Zeilen))
                '-------------------------------------------------------
                Image_Zahl_schreiben(ImageBox, Ausrichtung.EinzeiligRechts, Anzahl_Zeilen * Skalenteilung, 0, "", X1 - 100, X1 - 20, Wert_Y_Null - (Wert_Y_Hoehe_Teilung * Anzahl_Zeilen), Wert_Y_Null - (Wert_Y_Hoehe_Teilung * Anzahl_Zeilen), Font_Schriftgroesse_30, Font_Druckfarbe)
                '-------------------------------------------------------
            Next
            '--------------------------------------------------
            Dim Y_Skalierung As Decimal = Wert_Y_Hoehe_Gesamt / Maximalwert
            '--------------------------------------------------
            With Variable_XML_Import.Gebaeudebezogene_Daten.NWG_Aushang_Daten
                '--------------------------------------------------
                Wert_Y_oben = Wert_Y_Null
                '--------------------------------------------------
                If .Nutzenergiebedarf_Heizung_Diagramm.Wert > 0 Then
                    '--------------------------------------------------
                    Wert_Y_unten = Wert_Y_oben
                    Wert_Y_oben = Wert_Y_oben - (.Nutzenergiebedarf_Heizung_Diagramm.Wert * Y_Skalierung)
                    Image_Rechteck(ImageBox, Wert_X_Nutzenergie - (Breite_Balken / 2), Wert_X_Nutzenergie + (Breite_Balken / 2), Wert_Y_oben, Wert_Y_unten, Color.FromArgb(255, 238, 57, 35))
                    '--------------------------------------------------
                End If
                '--------------------------------------------------
                If .Nutzenergiebedarf_Trinkwarmwasser_Diagramm.Wert > 0 Then
                    '--------------------------------------------------
                    Wert_Y_unten = Wert_Y_oben
                    Wert_Y_oben = Wert_Y_oben - (.Nutzenergiebedarf_Trinkwarmwasser_Diagramm.Wert * Y_Skalierung)
                    Image_Rechteck(ImageBox, Wert_X_Nutzenergie - (Breite_Balken / 2), Wert_X_Nutzenergie + (Breite_Balken / 2), Wert_Y_oben, Wert_Y_unten, Color.FromArgb(255, 191, 192, 225))
                    '--------------------------------------------------
                End If
                '--------------------------------------------------
                If .Nutzenergiebedarf_Beleuchtung_Diagramm.Wert > 0 Then
                    '--------------------------------------------------
                    Wert_Y_unten = Wert_Y_oben
                    Wert_Y_oben = Wert_Y_oben - (.Nutzenergiebedarf_Beleuchtung_Diagramm.Wert * Y_Skalierung)
                    Image_Rechteck(ImageBox, Wert_X_Nutzenergie - (Breite_Balken / 2), Wert_X_Nutzenergie + (Breite_Balken / 2), Wert_Y_oben, Wert_Y_unten, Color.FromArgb(255, 243, 236, 25))
                    '--------------------------------------------------
                End If
                '--------------------------------------------------
                If .Nutzenergiebedarf_Lueftung_Diagramm.Wert > 0 Then
                    '--------------------------------------------------
                    Wert_Y_unten = Wert_Y_oben
                    Wert_Y_oben = Wert_Y_oben - (.Nutzenergiebedarf_Lueftung_Diagramm.Wert * Y_Skalierung)
                    Image_Rechteck(ImageBox, Wert_X_Nutzenergie - (Breite_Balken / 2), Wert_X_Nutzenergie + (Breite_Balken / 2), Wert_Y_oben, Wert_Y_unten, Color.FromArgb(255, 182, 182, 182))
                    '--------------------------------------------------
                End If
                '--------------------------------------------------
                If .Nutzenergiebedarf_Kuehlung_Befeuchtung_Diagramm.Wert > 0 Then
                    '--------------------------------------------------
                    Wert_Y_unten = Wert_Y_oben
                    Wert_Y_oben = Wert_Y_oben - (.Nutzenergiebedarf_Kuehlung_Befeuchtung_Diagramm.Wert * Y_Skalierung)
                    Image_Rechteck(ImageBox, Wert_X_Nutzenergie - (Breite_Balken / 2), Wert_X_Nutzenergie + (Breite_Balken / 2), Wert_Y_oben, Wert_Y_unten, Color.FromArgb(255, 52, 198, 244))
                    '--------------------------------------------------
                End If
                '--------------------------------------------------
                Wert_Y_oben = Wert_Y_Null
                '--------------------------------------------------
                If .Endenergiebedarf_Heizung_Diagramm.Wert > 0 Then
                    '--------------------------------------------------
                    Wert_Y_unten = Wert_Y_oben
                    Wert_Y_oben = Wert_Y_oben - (.Endenergiebedarf_Heizung_Diagramm.Wert * Y_Skalierung)
                    Image_Rechteck(ImageBox, Wert_X_Endenergie - (Breite_Balken / 2), Wert_X_Endenergie + (Breite_Balken / 2), Wert_Y_oben, Wert_Y_unten, Color.FromArgb(255, 238, 57, 35))
                    '--------------------------------------------------
                End If
                '--------------------------------------------------
                If .Endenergiebedarf_Trinkwarmwasser_Diagramm.Wert > 0 Then
                    '--------------------------------------------------
                    Wert_Y_unten = Wert_Y_oben
                    Wert_Y_oben = Wert_Y_oben - (.Endenergiebedarf_Trinkwarmwasser_Diagramm.Wert * Y_Skalierung)
                    Image_Rechteck(ImageBox, Wert_X_Endenergie - (Breite_Balken / 2), Wert_X_Endenergie + (Breite_Balken / 2), Wert_Y_oben, Wert_Y_unten, Color.FromArgb(255, 191, 192, 225))
                    '--------------------------------------------------
                End If
                '--------------------------------------------------
                If .Endenergiebedarf_Beleuchtung_Diagramm.Wert > 0 Then
                    '--------------------------------------------------
                    Wert_Y_unten = Wert_Y_oben
                    Wert_Y_oben = Wert_Y_oben - (.Endenergiebedarf_Beleuchtung_Diagramm.Wert * Y_Skalierung)
                    Image_Rechteck(ImageBox, Wert_X_Endenergie - (Breite_Balken / 2), Wert_X_Endenergie + (Breite_Balken / 2), Wert_Y_oben, Wert_Y_unten, Color.FromArgb(255, 243, 236, 25))
                    '--------------------------------------------------
                End If
                '--------------------------------------------------
                If .Endenergiebedarf_Lueftung_Diagramm.Wert > 0 Then
                    '--------------------------------------------------
                    Wert_Y_unten = Wert_Y_oben
                    Wert_Y_oben = Wert_Y_oben - (.Endenergiebedarf_Lueftung_Diagramm.Wert * Y_Skalierung)
                    Image_Rechteck(ImageBox, Wert_X_Endenergie - (Breite_Balken / 2), Wert_X_Endenergie + (Breite_Balken / 2), Wert_Y_oben, Wert_Y_unten, Color.FromArgb(255, 182, 182, 182))
                    '--------------------------------------------------
                End If
                '--------------------------------------------------
                If .Endenergiebedarf_Kuehlung_Befeuchtung_Diagramm.Wert > 0 Then
                    '--------------------------------------------------
                    Wert_Y_unten = Wert_Y_oben
                    Wert_Y_oben = Wert_Y_oben - (.Endenergiebedarf_Kuehlung_Befeuchtung_Diagramm.Wert * Y_Skalierung)
                    Image_Rechteck(ImageBox, Wert_X_Endenergie - (Breite_Balken / 2), Wert_X_Endenergie + (Breite_Balken / 2), Wert_Y_oben, Wert_Y_unten, Color.FromArgb(255, 52, 198, 244))
                    '--------------------------------------------------
                End If
                '--------------------------------------------------
                Wert_Y_oben = Wert_Y_Null
                '--------------------------------------------------
                If .Primaerenergiebedarf_Heizung_Diagramm.Wert > 0 Then
                    '--------------------------------------------------
                    Wert_Y_unten = Wert_Y_oben
                    Wert_Y_oben = Wert_Y_oben - (.Primaerenergiebedarf_Heizung_Diagramm.Wert * Y_Skalierung)
                    Image_Rechteck(ImageBox, Wert_X_Primaerenergie - (Breite_Balken / 2), Wert_X_Primaerenergie + (Breite_Balken / 2), Wert_Y_oben, Wert_Y_unten, Color.FromArgb(255, 238, 57, 35))
                    '--------------------------------------------------
                End If
                '--------------------------------------------------
                If .Primaerenergiebedarf_Trinkwarmwasser_Diagramm.Wert > 0 Then
                    '--------------------------------------------------
                    Wert_Y_unten = Wert_Y_oben
                    Wert_Y_oben = Wert_Y_oben - (.Primaerenergiebedarf_Trinkwarmwasser_Diagramm.Wert * Y_Skalierung)
                    Image_Rechteck(ImageBox, Wert_X_Primaerenergie - (Breite_Balken / 2), Wert_X_Primaerenergie + (Breite_Balken / 2), Wert_Y_oben, Wert_Y_unten, Color.FromArgb(255, 191, 192, 225))
                    '--------------------------------------------------
                End If
                '--------------------------------------------------
                If .Primaerenergiebedarf_Beleuchtung_Diagramm.Wert > 0 Then
                    '--------------------------------------------------
                    Wert_Y_unten = Wert_Y_oben
                    Wert_Y_oben = Wert_Y_oben - (.Primaerenergiebedarf_Beleuchtung_Diagramm.Wert * Y_Skalierung)
                    Image_Rechteck(ImageBox, Wert_X_Primaerenergie - (Breite_Balken / 2), Wert_X_Primaerenergie + (Breite_Balken / 2), Wert_Y_oben, Wert_Y_unten, Color.FromArgb(255, 243, 236, 25))
                    '--------------------------------------------------
                End If
                '--------------------------------------------------
                If .Primaerenergiebedarf_Lueftung_Diagramm.Wert > 0 Then
                    '--------------------------------------------------
                    Wert_Y_unten = Wert_Y_oben
                    Wert_Y_oben = Wert_Y_oben - (.Primaerenergiebedarf_Lueftung_Diagramm.Wert * Y_Skalierung)
                    Image_Rechteck(ImageBox, Wert_X_Primaerenergie - (Breite_Balken / 2), Wert_X_Primaerenergie + (Breite_Balken / 2), Wert_Y_oben, Wert_Y_unten, Color.FromArgb(255, 182, 182, 182))
                    '--------------------------------------------------
                End If
                '--------------------------------------------------
                If .Primaerenergiebedarf_Kuehlung_Befeuchtung_Diagramm.Wert > 0 Then
                    '--------------------------------------------------
                    Wert_Y_unten = Wert_Y_oben
                    Wert_Y_oben = Wert_Y_oben - ((.Primaerenergiebedarf_Kuehlung_Befeuchtung_Diagramm.Wert) * Y_Skalierung)
                    Image_Rechteck(ImageBox, Wert_X_Primaerenergie - (Breite_Balken / 2), Wert_X_Primaerenergie + (Breite_Balken / 2), Wert_Y_oben, Wert_Y_unten, Color.FromArgb(255, 52, 198, 244))
                    '--------------------------------------------------
                End If
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
#Region "PDF erstellen"
    '--------------------------------------------------
    ''' <summary>
    ''' Mit dieser Anweisung wird das PDF erzeugt.
    ''' </summary>
    Sub PDF_erzeugen(ByVal PDF_offnen As Boolean)
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            Dim Seitenzahl As Integer = 0
            '--------------------------------------------------
            Dim Wert_X As Integer
            '--------------------------------------------------
            For Wert_X = 1 To 40
                If Variable_XML_Import.Bilddateien(Wert_X) <> "" Then
                    Seitenzahl += 1
                End If
            Next
            '--------------------------------------------------
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            '--------------------------------------------------
            Dim pdf As Modul_PDF = New Modul_PDF()
            '--------------------------------------------------
            pdf.SetLicenseKey(My.Settings.PDF_Key)
            '--------------------------------------------------
            pdf.CreateNewPDF(Nothing)
            '--------------------------------------------------
            pdf.SetOnErrorProc(AddressOf PDFError)
            '--------------------------------------------------
            pdf.SetDocInfoA(TDocumentInfo.diCreator, My.Settings.Creator)
            pdf.SetDocInfoA(TDocumentInfo.diSubject, My.Settings.Subject)
            pdf.SetDocInfoA(TDocumentInfo.diTitle, My.Settings.Title)
            pdf.SetDocInfoA(TDocumentInfo.diAuthor, My.Settings.Author)
            pdf.SetDocInfoA(TDocumentInfo.diProducer, My.Settings.Producer)
            pdf.SetDocInfoA(TDocumentInfo.diKeywords, My.Settings.Keywords)
            pdf.SetViewerPreferences(TViewerPreference.vpDisplayDocTitle, TViewPrefAddVal.avNone)
            '--------------------------------------------------
            Select Case Variable_Steuerung.Ausgabequalitaet
                Case 0
                    PDF_PDFVersion = TPDFVersion.pvPDF_1_4
                    PDF_JPEGQuality = 70
                    PDF_Resolution = 150
                Case 1
                    PDF_PDFVersion = TPDFVersion.pvPDF_1_4
                    PDF_JPEGQuality = 70
                    PDF_Resolution = 300
                Case 2
                    PDF_PDFVersion = TPDFVersion.pvPDF_1_4
                    PDF_JPEGQuality = 70
                    PDF_Resolution = 600
                Case Else
                    PDF_PDFVersion = TPDFVersion.pvPDF_1_4
                    PDF_JPEGQuality = 70
                    PDF_Resolution = 300
            End Select
            '--------------------------------------------------
            pdf.SetPDFVersion(PDF_PDFVersion)
            pdf.SetJPEGQuality(PDF_JPEGQuality)
            pdf.SetResolution(PDF_Resolution)
            pdf.SetCompressionFilter(PDF_CompressionFilter)
            pdf.SetCompressionLevel(PDF_CompressionLevel)
            '--------------------------------------------------
            If Seitenzahl > 0 Then
                '--------------------------------------------------
                Dim WertXX As Integer
                '--------------------------------------------------
                For WertXX = 1 To Seitenzahl
                    '--------------------------------------------------
                    pdf.Append()
                    pdf.InsertImageFromBuffer(0.0, 0.0, 595.0, 842.0, GetArrayFromImage(String_to_Image(Variable_XML_Import.Bilddateien(WertXX))), WertXX - 1)
                    pdf.EndPage()
                    '--------------------------------------------------
                    Speicher_freigeben()
                    '--------------------------------------------------
                Next
                '--------------------------------------------------
            End If
            '--------------------------------------------------
            If Variable_Steuerung.PDF_Registriernummer = True Then
                '--------------------------------------------------
                Dim myFile As New FileInfo(Variable_Steuerung.Ausgabedatei)
                Dim Extension As String = myFile.Extension
                Dim Name_ohne_Extension As String = myFile.Name.Remove(myFile.Name.Length - Extension.Length, Extension.Length)
                Variable_Steuerung.Ausgabedatei = Name_ohne_Extension & "_Energieausweis_" & Variable_Steuerung.Registriernummer & Extension
                '--------------------------------------------------
            End If
            '--------------------------------------------------
            If pdf.HaveOpenDoc() Then
                '--------------------------------------------------
                Dim filePath As String = Variable_Steuerung.Ausgabepfad & "\" & Variable_Steuerung.Ausgabedatei
                '--------------------------------------------------
                If pdf.OpenOutputFile(filePath) Then
                    '--------------------------------------------------
                    If pdf.CloseFileEx("", PDF_Passwort, 3, TRestrictions.rsModify) Then
                        '--------------------------------------------------
                        If PDF_offnen = True Then
                            '--------------------------------------------------
                            Dim p As System.Diagnostics.Process = New System.Diagnostics.Process()
                            p.StartInfo.FileName = filePath
                            p.Start()
                            '--------------------------------------------------
                        End If
                        '--------------------------------------------------
                    Else
                        '--------------------------------------------------
                        MsgBox("Die PDF Datei (" & Variable_Steuerung.Ausgabepfad & "\" & Variable_Steuerung.Ausgabedatei & ") des Energieausweises konnte nicht gespeichert werden." & vbCrLf & vbCrLf & "Haben Sie eine PDF Datei mit dem gleichen Dateinamen geöffnet?", MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly, "Unerwarteter Anwendungsfehler:")
                        '--------------------------------------------------
                    End If
                    '--------------------------------------------------
                Else
                    '--------------------------------------------------
                    MsgBox("Die PDF Datei (" & Variable_Steuerung.Ausgabepfad & "\" & Variable_Steuerung.Ausgabedatei & ") des Energieausweises konnte nicht gespeichert werden." & vbCrLf & vbCrLf & "Haben Sie eine PDF Datei mit dem gleichen Dateinamen geöffnet?", MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly, "Unerwarteter Anwendungsfehler:")
                    '--------------------------------------------------
                End If
                '--------------------------------------------------
            End If
            '--------------------------------------------------
            Windows.Forms.Cursor.Current = Cursors.Default
            '--------------------------------------------------
        Catch ex As Exception
            Fehlerfenster(ex)
        End Try
        '--------------------------------------------------
    End Sub
    '--------------------------------------------------
    ''' <summary>
    ''' Mit dieser Anweisung dient als Hilfsanweisung für das erzeugen eines PDF.
    ''' </summary>
    Private Function PDFError(ByVal Data As IntPtr, ByVal ErrCode As Integer, ByVal ErrMessage As IntPtr, ByVal ErrType As Integer) As Integer
        '--------------------------------------------------
        Console.Write("{0}" + Chr(10), System.Runtime.InteropServices.Marshal.PtrToStringAnsi(ErrMessage))
        Return 0
        '--------------------------------------------------
    End Function
    '--------------------------------------------------
    ''' <summary>
    ''' Mit dieser Anweisung dient als Hilfsanweisung für das erzeugen eines PDF.
    ''' </summary>
    Public Function GetArrayFromImage(image As Image) As Byte()
        '--------------------------------------------------
        Try
            '--------------------------------------------------
            If image IsNot Nothing Then
                Dim ic As New ImageConverter()
                Dim buffer As Byte() = DirectCast(ic.ConvertTo(image, GetType(Byte())), Byte())
                Return buffer
            Else
                Return Nothing
            End If
            '--------------------------------------------------
        Catch ex As Exception
            Fehlerfenster(ex)
        End Try
        '--------------------------------------------------
    End Function
    '-----------------------------------------------------------------------------
#End Region
    '-----------------------------------------------------------------------------
#Region "Image to String"
    '---------------------------------------------------------------------------------
    ''' <summary>
    ''' Mit dieser Anweisung wird ein Image als Base64 String umgewandelt.
    ''' </summary>
    Public Function Image_to_String(ByVal image As Image) As String
        '---------------------------------------------------------------------------------
        Try
            '---------------------------------------------------------------------------------
            If Not image Is Nothing Then
                '---------------------------------------------------------------------------------
                Dim ic As New ImageConverter()
                Dim buffer As Byte() = DirectCast(ic.ConvertTo(image, GetType(Byte())), Byte())
                Return Convert.ToBase64String(buffer, Base64FormattingOptions.InsertLineBreaks)
                '---------------------------------------------------------------------------------
            Else
                '---------------------------------------------------------------------------------
                Return Nothing
                '---------------------------------------------------------------------------------
            End If
            '---------------------------------------------------------------------------------
        Catch ex As Exception
            '---------------------------------------------------------------------------------
            Fehlerfenster(ex)
            Return Nothing
            '---------------------------------------------------------------------------------
        End Try
        '---------------------------------------------------------------------------------
    End Function
    '---------------------------------------------------------------------
    ''' <summary>
    ''' Mit dieser Anweisung wird ein Base64 String zum Image umgewandelt.
    ''' </summary>
    Public Function String_to_Image(ByVal base64String As String) As Image
        '---------------------------------------------------------------------------------
        Try
            '---------------------------------------------------------------------------------
            Dim buffer As Byte() = Convert.FromBase64String(base64String)
            '---------------------------------------------------------------------------------
            If Not buffer Is Nothing Then
                Dim ic As New ImageConverter()
                Return TryCast(ic.ConvertFrom(buffer), Image)
            Else
                Return Nothing
            End If
            '---------------------------------------------------------------------------------
        Catch ex As Exception
            '---------------------------------------------------------------------------------
            Fehlerfenster(ex)
            Return Nothing
            '---------------------------------------------------------------------------------
        End Try
        '---------------------------------------------------------------------------------
    End Function
    '---------------------------------------------------------------------------------
#End Region
    '---------------------------------------------------------------------------------
#Region "Bitmap skalieren"
    '--------------------------------------------------
    ''' <summary>
    ''' Mit dieser Anweisung wird ein Bitmap skaliert.
    ''' </summary>
    Function Bitmap_skalieren(ByVal pic As Bitmap, ByVal Scale As Single) As Bitmap
        '--------------------------------------------------
        Dim w, h As Integer
        '--------------------------------------------------
        w = pic.Width * Scale
        h = pic.Height * Scale
        '--------------------------------------------------
        Dim bmp As Bitmap = New Bitmap(w, h)
        '--------------------------------------------------
        Using g As Graphics = Graphics.FromImage(bmp)
            g.DrawImage(pic, 0, 0, bmp.Width, bmp.Height)
        End Using
        '--------------------------------------------------
        Bitmap_skalieren = bmp
        '--------------------------------------------------
    End Function
    '--------------------------------------------------
#End Region
    '-----------------------------------------------------------------------------
End Module
