Public Module Formattings

    Public Enum DmsFormat As Integer
        SimpleDMS = 0
        SimpleDM = 1
        SimpleD = 2
        VerboseDMS = 3
        VerboseDM = 4
        VerboseD = 5
        SimpleR = 6
        EsriDMS = 7
        EsriDM = 8
        EsriD = 9
        EsriPackedDMS = 10
        UkooaDMS = 11
        NMEA = 12
        SpacedDMS = 13
        SpacedDM = 14
    End Enum

    Public Enum DmsSign As Integer
        PlusMinus = 0
        Prefix = 1
        Suffix = 2
        Generic = 3
    End Enum

    ''' <summary>
    ''' Create an array of angle formatting strings
    ''' </summary>
    Public Function ListDmsFormats() As List(Of String)
        Dim FormatArray As New List(Of String)

        FormatArray.Add("DDD:MM:SS.000")
        FormatArray.Add("DDD:MM.000")
        FormatArray.Add("DDD.000")
        FormatArray.Add("DDD°MM" & Convert.ToChar(39) & "SS.000" & Convert.ToChar(34))
        FormatArray.Add("DDD°MM.000" & Convert.ToChar(39))
        FormatArray.Add("DDD.000°")
        FormatArray.Add("R.000000r")
        FormatArray.Add("DDD° MM" & Convert.ToChar(39) & " SS.000" & Convert.ToChar(34))
        FormatArray.Add("DDD° MM.000" & Convert.ToChar(39))
        FormatArray.Add("DDD.000°")
        FormatArray.Add("DDD.MMSS000000")
        FormatArray.Add("DDMMSS.00")
        FormatArray.Add("DDMM.000")
        FormatArray.Add("DD MM SS.000")
        FormatArray.Add("DD MM.000")

        Return FormatArray
    End Function

    ''' <summary>
    ''' Create an array of angle formatting strings
    ''' </summary>
    Public Function ListDmsSigns() As List(Of String)
        Dim FormatArray As New List(Of String)

        FormatArray.Add("+/-")
        FormatArray.Add("Prefix (E/W & N/S)")
        FormatArray.Add("Suffix (E/W & N/S)")
        FormatArray.Add("Generic Angle")

        Return FormatArray
    End Function

    Public Enum MetricSign As Integer
        Number = 0
        Unit = 1
        Prefix = 2
        Suffix = 3
        UnitPrefix = 4
        UnitSuffix = 5
    End Enum

    ''' <summary>
    ''' Create an array of metric coordinates formatting strings
    ''' </summary>
    Public Function ListMetricSigns() As List(Of String)
        Dim SignArray As New List(Of String)

        SignArray.Add("Simple Number")
        SignArray.Add("Unit (m)")
        SignArray.Add("Prefix (E/W & N/S)")
        SignArray.Add("Suffix (E/W & N/S)")
        SignArray.Add("Prefix (E/W & N/S) & Unit (m)")
        SignArray.Add("Unit (m) & Suffix (E/W & N/S)")

        Return SignArray
    End Function

    ''' <summary>
    ''' Convert a number in to string representation with given decimals
    ''' </summary>
    ''' <param name="value">The value to convert</param>
    ''' <param name="decimals">Quantity of numbers after the decimal point</param>
    ''' <param name="leaderPaddings">Quantity of numbers before the decimal point</param>
    ''' <returns>String representation of the given value</returns>
    ''' <remarks></remarks>
    Public Function FormatNumber(value As Double, decimals As Integer, leaderPaddings As Integer) As String
        Dim TempString As String
        Dim IntegerPart, DecimalPart, DecimalMultiplier As Integer
        Try
            IntegerPart = CInt(Math.Truncate(value))
            DecimalMultiplier = CInt(Math.Pow(10.0, decimals))
            DecimalPart = Math.Abs(CInt(Math.Round((value - Math.Truncate(value)) * DecimalMultiplier)))
            If DecimalPart >= DecimalMultiplier Then
                IntegerPart = IntegerPart + 1
                DecimalPart = DecimalPart - 10
            End If
            TempString = IntegerPart.ToString(Globalization.CultureInfo.InvariantCulture)
            If leaderPaddings > TempString.Length Then
                TempString = TempString.PadLeft(leaderPaddings, CChar("0"))
            End If
            TempString = TempString & "." & DecimalPart.ToString(Globalization.CultureInfo.InvariantCulture).PadLeft(decimals, CChar("0"))
        Catch ex As Exception
            TempString = value.ToString(Globalization.CultureInfo.InvariantCulture)
        End Try
        Return TempString
    End Function

    ''' <summary>
    ''' Convert a number in to string representation with given decimals
    ''' </summary>
    ''' <param name="value">The value to convert</param>
    ''' <param name="decimals">Quantity of numbers after the decimal point</param>
    ''' <returns>String representation of the given value</returns>
    Public Function FormatNumber(value As Double, decimals As Integer) As String
        Return FormatNumber(value, decimals, 0)
    End Function

    ''' <summary>
    ''' Retrieve the angular value from a formatted string
    ''' </summary>
    ''' <param name="angleString">Formatted string to parse</param>
    ''' <param name="angleFormat">Expected angle format</param>
    ''' <param name="signFormat">Expected sign format</param>
    ''' <returns>Parsed angle</returns>
    ''' <remarks></remarks>
    Public Function DmsParse(ByVal angleString As String, ByVal angleFormat As DmsFormat, ByVal signFormat As DmsSign) As Double
        Dim SignMult As Integer
        Dim Degs, Mins, Secs, DecSecs, Signs, TmpString, TmpSubString() As String
        Dim ParsedAngleD, ParsedAngleR As Double
        'Define the number format
        Dim InternalCulture As New Globalization.CultureInfo("en-US")
        InternalCulture.NumberFormat.NumberDecimalSeparator = "."
        InternalCulture.NumberFormat.NumberGroupSeparator = "'"
        If angleString Is Nothing Then
            Return Double.NaN
        End If
        'Clean the Angle string from extra space characters
        angleString = angleString.Trim
        'Retrieve the sign of the angle
        Select Case signFormat
            Case DmsSign.Suffix 'Suffix
                Signs = angleString.Substring(angleString.Length - 1)
            Case Else 'Prefix and Sign
                Signs = angleString.Substring(0, 1)
        End Select
        'Handle the sign of the angle
        Select Case Signs.ToUpper
            Case "E", "N"
                SignMult = 1
            Case "W", "S", "-"
                SignMult = -1
            Case Else
                SignMult = 1
        End Select
        'Clean the Angle string from extra characters
        Do
            Select Case angleString.Substring(0, 1)
                Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
                    Exit Do
                Case Else
                    angleString = angleString.Substring(1)
            End Select
        Loop
        Do
            Select Case angleString.Substring(angleString.Length - 1)
                Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", Convert.ToChar(34).ToString, Convert.ToChar(39).ToString, Convert.ToChar(176).ToString
                    Exit Do
                Case Else
                    angleString = angleString.Substring(0, angleString.Length - 1)
            End Select
        Loop
        'Parse the Angle value
        Select Case angleFormat
            Case DmsFormat.SimpleDMS 'ddd:mm:ss.000
                TmpSubString = angleString.Split(CChar(":"))
                If TmpSubString.Count = 3 Then
                    Try
                        ParsedAngleD = Integer.Parse(TmpSubString(0), Globalization.CultureInfo.InvariantCulture)
                        ParsedAngleD = ParsedAngleD + (Integer.Parse(TmpSubString(1), Globalization.CultureInfo.InvariantCulture) / 60)
                        ParsedAngleD = ParsedAngleD + (Double.Parse(TmpSubString(2), Globalization.CultureInfo.InvariantCulture) / 3600)
                    Catch ex As Exception
                        Return Double.NaN
                    End Try
                Else
                    Return Double.NaN
                End If

            Case DmsFormat.SimpleDM 'ddd:mm.000
                TmpSubString = angleString.Split(CChar(":"))
                If TmpSubString.Count = 2 Then
                    Try
                        ParsedAngleD = Integer.Parse(TmpSubString(0), Globalization.CultureInfo.InvariantCulture)
                        ParsedAngleD = ParsedAngleD + (Double.Parse(TmpSubString(1), Globalization.CultureInfo.InvariantCulture) / 60)
                    Catch ex As Exception
                        Return Double.NaN
                    End Try
                Else
                    Return Double.NaN
                End If

            Case DmsFormat.SimpleD 'ddd.000
                ParsedAngleD = (Double.Parse(angleString, Globalization.CultureInfo.InvariantCulture))

            Case DmsFormat.VerboseDMS, DmsFormat.EsriDMS  'dd°mm'ss"
                TmpSubString = angleString.Split(CChar("°"))
                If TmpSubString.Count = 2 Then
                    Degs = TmpSubString(0)
                    TmpString = TmpSubString(1)
                    TmpSubString = TmpString.Split(CChar("'"))
                    If TmpSubString.Count = 2 Then
                        Mins = TmpSubString(0)
                        Secs = TmpSubString(1)
                        If Secs.EndsWith(Convert.ToChar(39), StringComparison.CurrentCulture) Then
                            Secs = Secs.Substring(0, Secs.Length - 1)
                        End If
                        Try
                            ParsedAngleD = Integer.Parse(Degs, Globalization.CultureInfo.InvariantCulture)
                            ParsedAngleD = ParsedAngleD + (Integer.Parse(Mins, Globalization.CultureInfo.InvariantCulture) / 60)
                            ParsedAngleD = ParsedAngleD + (Double.Parse(Secs, Globalization.CultureInfo.InvariantCulture) / 3600)
                        Catch ex As Exception
                            Return Double.NaN
                        End Try
                    Else
                        Return Double.NaN
                    End If
                Else
                    Return Double.NaN
                End If

            Case DmsFormat.VerboseDM, DmsFormat.EsriDM  'dd°mm'
                TmpSubString = angleString.Split(CChar("°"))
                If TmpSubString.Count = 2 Then
                    Try
                        ParsedAngleD = Integer.Parse(TmpSubString(0), Globalization.CultureInfo.InvariantCulture)
                        ParsedAngleD = ParsedAngleD + (Double.Parse(TmpSubString(1), Globalization.CultureInfo.InvariantCulture) / 60)
                    Catch ex As Exception
                        Return Double.NaN
                    End Try
                Else
                    Return Double.NaN
                End If

            Case DmsFormat.VerboseD, DmsFormat.EsriD  'dd°
                If angleString.EndsWith(Convert.ToChar(39), StringComparison.CurrentCulture) Then
                    Degs = angleString.Substring(0, angleString.Length - 1)
                Else
                    Degs = angleString
                End If
                Try
                    ParsedAngleD = Double.Parse(Degs, Globalization.CultureInfo.InvariantCulture)
                Catch ex As Exception
                    Return Double.NaN
                End Try

            Case DmsFormat.SimpleR  'RRR.000000r
                If angleString.EndsWith("r", StringComparison.CurrentCulture) Then
                    Degs = angleString.Substring(0, angleString.Length - 1)
                Else
                    Degs = angleString
                End If
                Try
                    ParsedAngleD = (Double.Parse(Degs, Globalization.CultureInfo.InvariantCulture) * (180 / Math.PI))
                Catch ex As Exception
                    Return Double.NaN
                End Try

            Case DmsFormat.EsriPackedDMS  'DDD.MMSSsss
                TmpSubString = angleString.Split(CChar("."))
                If TmpSubString.Count = 2 Then
                    Degs = TmpSubString(0)
                    Mins = TmpSubString(1).Substring(0, 2)
                    Secs = TmpSubString(1).Substring(2, 2)
                    DecSecs = TmpSubString(1).Substring(4)
                    Try
                        ParsedAngleD = Integer.Parse(Degs, Globalization.CultureInfo.InvariantCulture)
                        ParsedAngleD = ParsedAngleD + (Integer.Parse(Mins, Globalization.CultureInfo.InvariantCulture) / 60)
                        ParsedAngleD = ParsedAngleD + (Double.Parse(Secs & "." & DecSecs, Globalization.CultureInfo.InvariantCulture) / 3600)
                    Catch ex As Exception
                        Return Double.NaN
                    End Try
                Else
                    Return Double.NaN
                End If

            Case DmsFormat.UkooaDMS  'DDMMSS.ss
                TmpSubString = angleString.Split(CChar("."))
                If TmpSubString.Count = 2 Then
                    Degs = TmpSubString(0).Substring(0, TmpSubString.Count - 4)
                    Mins = TmpSubString(0).Substring(TmpSubString.Count - 4, 2)
                    Secs = TmpSubString(0).Substring(TmpSubString.Count - 2, 2)
                    DecSecs = TmpSubString(1)
                    Try
                        ParsedAngleD = Integer.Parse(Degs, Globalization.CultureInfo.InvariantCulture)
                        ParsedAngleD = ParsedAngleD + (Integer.Parse(Mins, Globalization.CultureInfo.InvariantCulture) / 60)
                        ParsedAngleD = ParsedAngleD + (Double.Parse(Secs & "." & DecSecs, Globalization.CultureInfo.InvariantCulture) / 3600)
                    Catch ex As Exception
                        Return Double.NaN
                    End Try
                Else
                    Return Double.NaN
                End If

            Case DmsFormat.NMEA 'DDMM.mmm
                TmpSubString = angleString.Split(CChar("."))
                If TmpSubString.Count = 2 Then
                    Degs = TmpSubString(0).Substring(0, TmpSubString.Count - 2)
                    Mins = TmpSubString(0).Substring(TmpSubString.Count - 2, 2)
                    Mins = Mins & "." & TmpSubString(1)
                    Try
                        ParsedAngleD = Integer.Parse(Degs, Globalization.CultureInfo.InvariantCulture)
                        ParsedAngleD = ParsedAngleD + (Double.Parse(Mins, Globalization.CultureInfo.InvariantCulture) / 60)
                    Catch ex As Exception
                        Return Double.NaN
                    End Try
                Else
                    Return Double.NaN
                End If

            Case DmsFormat.SpacedDMS 'ddd mm ss.000
                TmpSubString = angleString.Split(CChar(" "))
                If TmpSubString.Count = 3 Then
                    Try
                        ParsedAngleD = Integer.Parse(TmpSubString(0), Globalization.CultureInfo.InvariantCulture)
                        ParsedAngleD = ParsedAngleD + (Integer.Parse(TmpSubString(1), Globalization.CultureInfo.InvariantCulture) / 60)
                        ParsedAngleD = ParsedAngleD + (Double.Parse(TmpSubString(2), Globalization.CultureInfo.InvariantCulture) / 3600)
                    Catch ex As Exception
                        Return Double.NaN
                    End Try
                Else
                    Return Double.NaN
                End If

            Case DmsFormat.SpacedDM  'ddd:mm.000
                TmpSubString = angleString.Split(CChar(" "))
                If TmpSubString.Count = 2 Then
                    Try
                        ParsedAngleD = Integer.Parse(TmpSubString(0), Globalization.CultureInfo.InvariantCulture)
                        ParsedAngleD = ParsedAngleD + (Double.Parse(TmpSubString(1), Globalization.CultureInfo.InvariantCulture) / 60)
                    Catch ex As Exception
                        Return Double.NaN
                    End Try
                Else
                    Return Double.NaN
                End If

            Case Else
                Return Double.NaN

        End Select
        ParsedAngleD = ParsedAngleD * SignMult
        ParsedAngleR = ParsedAngleD / (180 / Math.PI)
        Return ParsedAngleR

    End Function

    ''' <summary>
    ''' Convert the degree value in a formatted geographic coordinate string
    ''' </summary>
    ''' <param name="angleR">Latitude or Longitude value</param>
    ''' <param name="angleFormat">Format for the numbers</param>
    ''' <param name="signFormat">Format for the sign</param>
    ''' <param name="decimals">Decimal places</param>
    ''' <param name="isLat">True for Latitude, False for Longitude values</param>
    ''' <returns>Formatted String</returns>
    ''' <remarks></remarks>
    Public Function FormatDMS(ByVal angleR As Double, ByVal angleFormat As DmsFormat, ByVal signFormat As DmsSign,
                              ByVal decimals As Integer, ByVal isLat As Boolean) As String
        Dim Degs, Mins, Secs, AngleD As Double
        Dim Lbl, Result As String
        Dim DblQt, SngQt As Char
        DblQt = Convert.ToChar(34)
        SngQt = Convert.ToChar(39)

        AngleD = angleR * (180 / Math.PI)

        'Evaluate the quadrant of the angle
        If isLat Then
            If angleR < 0 Then
                Lbl = If(signFormat = DmsSign.PlusMinus, "-", "S")
            Else
                Lbl = If(signFormat = DmsSign.PlusMinus, "+", "N")
            End If
        Else
            If angleR < 0 Then
                Lbl = If(signFormat = DmsSign.PlusMinus, "-", "W")
            Else
                Lbl = If(signFormat = DmsSign.PlusMinus, "+", "E")
            End If
        End If

        'Take the absolute value of angle
        AngleD = Math.Abs(AngleD)
        'Compose the formatted string
        Select Case angleFormat
            Case DmsFormat.SimpleDMS  'dd:mm:ss
                Degs = Math.Truncate(AngleD)
                Mins = Math.Truncate((AngleD - Degs) * 60)
                Secs = (((AngleD - Degs) * 60) - Mins) * 60
                If Math.Round(Secs, decimals) = 60.0 Then
                    Secs = 0
                    Mins = Mins + 1
                    If Mins = 60 Then
                        Mins = 0
                        Degs = Degs + 1
                    End If
                End If
                Result = Degs.ToString("00", Globalization.CultureInfo.InvariantCulture) & ":"
                Result = Result & Mins.ToString("00", Globalization.CultureInfo.InvariantCulture) & ":"
                Result = Result & FormatNumber(Secs, decimals, 2)

            Case DmsFormat.SimpleDM  'dd:mm
                Degs = Math.Truncate(AngleD)
                Mins = (AngleD - Degs) * 60
                If Math.Round(Mins, decimals) = 60.0 Then
                    Mins = 0
                    Degs = Degs + 1
                End If
                Result = Degs.ToString("00", Globalization.CultureInfo.InvariantCulture) & ":" & FormatNumber(Mins, decimals, 2)

            Case DmsFormat.SimpleD  'dd
                Result = FormatNumber(AngleD, decimals)

            Case DmsFormat.VerboseDMS  'dd°mm'ss"
                Degs = Math.Truncate(AngleD)
                Mins = Math.Truncate((AngleD - Degs) * 60)
                Secs = (((AngleD - Degs) * 60) - Mins) * 60
                If Math.Round(Secs, decimals) = 60.0 Then
                    Secs = 0
                    Mins = Mins + 1
                    If Mins = 60 Then
                        Mins = 0
                        Degs = Degs + 1
                    End If
                End If
                Result = Degs.ToString("00", Globalization.CultureInfo.InvariantCulture) & "°"
                Result = Result & Mins.ToString("00", Globalization.CultureInfo.InvariantCulture) & SngQt
                Result = Result & FormatNumber(Secs, decimals, 2) & DblQt

            Case DmsFormat.VerboseDM  'dd°mm'
                Degs = Math.Truncate(AngleD)
                Mins = (AngleD - Degs) * 60
                If Math.Round(Mins, decimals) = 60.0 Then
                    Mins = 0
                    Degs = Degs + 1
                End If
                Result = Degs.ToString("00", Globalization.CultureInfo.InvariantCulture) & "°"
                Result = Result & FormatNumber(Mins, decimals, 2) & SngQt

            Case DmsFormat.VerboseD  'dd°
                Result = FormatNumber(AngleD, decimals) & "°"

            Case DmsFormat.SimpleR  'rad
                Result = FormatNumber(angleR, decimals) & "r"

            Case DmsFormat.EsriDMS  'dd° mm' ss"
                Degs = Math.Truncate(AngleD)
                Mins = Math.Truncate((AngleD - Degs) * 60)
                Secs = (((AngleD - Degs) * 60) - Mins) * 60
                If Math.Round(Secs, decimals) = 60.0 Then
                    Secs = 0
                    Mins = Mins + 1
                    If Mins = 60 Then
                        Mins = 0
                        Degs = Degs + 1
                    End If
                End If
                Result = Degs.ToString("00", Globalization.CultureInfo.InvariantCulture) & "° "
                Result = Result & Mins.ToString("00", Globalization.CultureInfo.InvariantCulture) & SngQt
                Result = Result & FormatNumber(Secs, decimals, 2) & DblQt

            Case DmsFormat.EsriDM  'dd° mm'
                Degs = Math.Truncate(AngleD)
                Mins = (AngleD - Degs) * 60
                If Math.Round(Mins, decimals) = 60.0 Then
                    Mins = 0
                    Degs = Degs + 1
                End If
                Result = Degs.ToString("00", Globalization.CultureInfo.InvariantCulture) & "° "
                Result = Result & FormatNumber(Mins, decimals, 2) & SngQt

            Case DmsFormat.EsriD  'dd°
                Result = FormatNumber(AngleD, decimals) & "°"

            Case DmsFormat.EsriPackedDMS  'DDD.MMSSsss
                Degs = Math.Truncate(AngleD)
                Mins = Math.Truncate((AngleD - Degs) * 60)
                Secs = ((((AngleD - Degs) * 60) - Mins) * 60)
                If Math.Round(Secs, decimals) = 60.0 Then
                    Secs = 0
                    Mins = Mins + 1
                    If Mins = 60 Then
                        Mins = 0
                        Degs = Degs + 1
                    End If
                End If
                Result = Degs.ToString("000", Globalization.CultureInfo.InvariantCulture) & "."
                Result = Result & Mins.ToString("00", Globalization.CultureInfo.InvariantCulture)
                Result = Result & FormatNumber(Secs, decimals, 2)

            Case DmsFormat.UkooaDMS  'DDMMSS.ss
                Degs = Math.Truncate(AngleD)
                Mins = Math.Truncate((AngleD - Degs) * 60)
                Secs = ((((AngleD - Degs) * 60) - Mins) * 60)
                If Math.Round(Secs, decimals) = 60.0 Then
                    Secs = 0
                    Mins = Mins + 1
                    If Mins = 60 Then
                        Mins = 0
                        Degs = Degs + 1
                    End If
                End If
                Result = Degs.ToString("0", Globalization.CultureInfo.InvariantCulture)
                Result = Result & Mins.ToString("00", Globalization.CultureInfo.InvariantCulture)
                Result = Result & FormatNumber(Secs, decimals, 2)

            Case DmsFormat.NMEA
                Degs = Math.Truncate(AngleD)
                Mins = (AngleD - Degs) * 60
                If Math.Round(Mins, decimals) = 60.0 Then
                    Mins = 0
                    Degs = Degs + 1
                End If
                Result = Degs.ToString("00", Globalization.CultureInfo.InvariantCulture)
                Result = Result & FormatNumber(Mins, decimals, 2)

            Case DmsFormat.SpacedDMS  'dd:mm:ss
                Degs = Math.Truncate(AngleD)
                Mins = Math.Truncate((AngleD - Degs) * 60)
                Secs = (((AngleD - Degs) * 60) - Mins) * 60
                If Math.Round(Secs, decimals) = 60.0 Then
                    Secs = 0
                    Mins = Mins + 1
                    If Mins = 60 Then
                        Mins = 0
                        Degs = Degs + 1
                    End If
                End If
                Result = Degs.ToString("00", Globalization.CultureInfo.InvariantCulture) & " "
                Result = Result & Mins.ToString("00", Globalization.CultureInfo.InvariantCulture) & " "
                Result = Result & FormatNumber(Secs, decimals, 2)

            Case DmsFormat.SpacedDM  'dd:mm
                Degs = Math.Truncate(AngleD)
                Mins = (AngleD - Degs) * 60
                If Math.Round(Mins, decimals) = 60.0 Then
                    Mins = 0
                    Degs = Degs + 1
                End If
                Result = Degs.ToString("00", Globalization.CultureInfo.InvariantCulture) & " "
                Result = Result & FormatNumber(Mins, decimals, 2)

            Case Else 'rad
                Result = FormatNumber(angleR, decimals) & "r"

        End Select
        'Add the correct sign
        Select Case signFormat
            Case DmsSign.Suffix
                Return Result & Lbl
            Case DmsSign.Generic
                Return If(angleR < 0, "-" & Result, Result)
            Case Else
                Return Lbl & Result
        End Select

    End Function

    ''' <summary>
    ''' Retrieve the metric value from a formatted string
    ''' </summary>
    ''' <param name="metricString">Text to be parsed</param>
    ''' <param name="metricFormat">Expected format</param>
    ''' <returns>Parsed value</returns>
    ''' <remarks></remarks>
    Public Function MetricParse(ByVal metricString As String, ByVal metricFormat As MetricSign) As Double
        Dim SignMult As Integer
        Dim Signs As String
        Dim ParsedMetric As Double

        'Clean the metric string from extra space characters
        If metricString Is Nothing Then Return Double.NaN
        metricString = metricString.Trim
        'Retrieve the sign of the coord
        Select Case metricFormat
            Case MetricSign.Suffix, MetricSign.UnitSuffix 'Suffix
                Signs = metricString.Substring(metricString.Length - 1, 1)
            Case Else 'Prefix and Sign
                Signs = metricString.Substring(0, 1)
        End Select

        'Handle the sign of the coord
        Select Case Signs.ToUpper
            Case "E", "N"
                SignMult = 1
            Case "W", "S", "-"
                SignMult = -1
            Case Else
                SignMult = 1
        End Select

        'Clean the coord string from extra characters
        Do
            Select Case metricString.Substring(0, 1)
                Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
                    'String ok on the left side
                    Exit Do
                Case Else
                    metricString = metricString.Substring(1)
            End Select
        Loop
        Do
            Select Case metricString.Substring(metricString.Length - 1, 1)
                Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
                    'String ok on the right side
                    Exit Do
                Case Else
                    metricString = metricString.Substring(0, metricString.Length - 1)
            End Select
        Loop
        'Convert the coord in a value
        Try
            ParsedMetric = CDbl(metricString)
        Catch ex As Exception
            ParsedMetric = 0
        End Try

        'Prepare the final result
        ParsedMetric = ParsedMetric * SignMult
        Return ParsedMetric

    End Function

    ''' <summary>
    ''' Convert the metric value in a formatted coordinate string
    ''' </summary>
    ''' <param name="metricCoord">East or North coordinate</param>
    ''' <param name="metricFormat">Format for the numbers</param>
    ''' <param name="decimals">Decimal places</param>
    ''' <param name="isNorth">True for North, False for East values</param>
    ''' <returns>Formatted string</returns>
    ''' <remarks></remarks>
    Public Function FormatMetric(ByVal metricCoord As Double, ByVal metricFormat As MetricSign, ByVal decimals As Integer, ByVal isNorth As Boolean) As String
        Dim Lbl, Result As String

        'Evaluate the quadrant of the angle
        If metricCoord < 0 Then
            Lbl = If(isNorth, "S", "W")
        Else
            Lbl = If(isNorth, "N", "E")
        End If

        'Prepare the coord string
        Select Case metricFormat
            Case MetricSign.Number
                Result = FormatNumber(metricCoord, decimals)
            Case MetricSign.Unit
                Result = FormatNumber(metricCoord, decimals) & "m"
            Case MetricSign.Prefix
                Result = Lbl & FormatNumber(Math.Abs(metricCoord), decimals)
            Case MetricSign.Suffix
                Result = FormatNumber(Math.Abs(metricCoord), decimals) & Lbl
            Case MetricSign.UnitPrefix
                Result = Lbl & FormatNumber(Math.Abs(metricCoord), decimals) & "m"
            Case MetricSign.UnitSuffix
                Result = FormatNumber(Math.Abs(metricCoord), decimals) & "m" & Lbl
            Case Else
                Result = FormatNumber(metricCoord, decimals)
        End Select
        Return Result
    End Function

    ''' <summary>
    ''' Convert the date value in to Posix value
    ''' </summary>
    ''' <param name="inDateTime">Input Date value</param>
    ''' <returns>Converted Posix value</returns>
    ''' <remarks></remarks>
    Public Function FormatPosixTime(inDateTime As DateTime) As Double
        Dim UnixTimeSpan As TimeSpan = (inDateTime.Subtract(New DateTime(1970, 1, 1, 0, 0, 0)))
        Return UnixTimeSpan.TotalSeconds
    End Function

    ''' <summary>
    ''' Convert Posix value in to a Date value
    ''' </summary>
    ''' <param name="inPosixTime">Input Posix value</param>
    ''' <returns>Converted Date value</returns>
    ''' <remarks></remarks>
    Public Function ParsePosixTime(inPosixTime As Double) As DateTime
        Return (New DateTime(1970, 1, 1, 0, 0, 0)).AddSeconds(inPosixTime)
    End Function

    ''' <summary>
    ''' Count the occurrences of a given char inside a string
    ''' </summary>
    ''' <param name="value">Input String</param>
    ''' <param name="ch">Char to count</param>
    ''' <returns>Number of char found</returns>
    ''' <remarks></remarks>
    Public Function CountCharacter(ByVal value As String, ByVal ch As Char) As Integer
        If value Is Nothing Then Return 0
        Dim cnt As Integer = 0
        For Each c As Char In value
            If c = ch Then cnt += 1
        Next
        Return cnt
    End Function

    ''' <summary>
    ''' Delete duplicated spaces inside the given string
    ''' </summary>
    ''' <param name="value">Input String</param>
    ''' <returns>Cleaned String</returns>
    ''' <remarks></remarks>
    Public Function DeleteExtraSpaceChar(value As String) As String
        If value Is Nothing Then Return ""
        Dim ResString, ActChar As String
        ResString = ""
        For c = 0 To value.Count - 1
            ActChar = value.Substring(c, 1)
            If ActChar = " " Then
                'Actual character is a space so check if the last saved character
                If ResString.Count > 0 Then
                    If ResString.Substring(ResString.Count - 1) = " " Then
                        'Skip space
                    Else
                        'Add the space
                        ResString = ResString & ActChar
                    End If
                End If
            Else
                'Actual character is not a space so add to the result string
                ResString = ResString & ActChar
            End If
        Next
        'Delete the last character if is a space
        If ResString.Substring(ResString.Count - 1, 1) = " " Then
            ResString = ResString.Substring(0, ResString.Count - 1)
        End If
        Return ResString
    End Function

    Public Function ParseJulianDate(dateString As String, formatString As String, ByRef resultDate As Date) As Boolean
        Dim year, jday, FebDays As Integer
        Dim fmt1, fmt2, date1, date2, newfmt, newdate As String

        If dateString Is Nothing Then
            resultDate = Date.MinValue
            Return False
        End If
        If formatString Is Nothing Then
            resultDate = Date.MinValue
            Return False
        End If

        Dim countY As Integer = CountCharacter(formatString, "y"c)
        If countY = 0 Then
            'Define actual time as dummy value
            resultDate = Date.MinValue
            Return False
        Else
            If countY <= 2 Then
                year = CInt(dateString.Substring(formatString.IndexOf("y"c), countY))
                year = If(year >= 70, 1900, 2000) + year
            ElseIf countY = 3 Then
                year = CInt(dateString.Substring(formatString.IndexOf("y"c), countY))
                year = If(year >= 970, 1000, 2000) + year
            Else '4 digits
                year = CInt(dateString.Substring(formatString.IndexOf("y"c), countY))
            End If
            FebDays = If(Date.IsLeapYear(year), 29, 28)
        End If

        Dim countJ As Integer = CountCharacter(formatString, "j"c)
        If countJ = 0 Then
            'Define actual time as dummy value
            resultDate = Date.MinValue
            Return False
        Else
            jday = CInt(dateString.Substring(formatString.IndexOf("j"c), countJ))

        End If

        Dim JulMonth As Integer

        If jday <= 31 Then 'Day of January
            JulMonth = 1
        Else
            jday = jday - 31
            If jday <= FebDays Then 'Day of February
                JulMonth = 2
            Else
                jday = jday - FebDays
                If jday <= 31 Then 'Day of March
                    JulMonth = 3
                Else
                    jday = jday - 31
                    If jday <= 30 Then 'Day of April
                        JulMonth = 4
                    Else
                        jday = jday - 30
                        If jday <= 31 Then 'Day of May
                            JulMonth = 5
                        Else
                            jday = jday - 31
                            If jday <= 30 Then 'Day of June
                                JulMonth = 6
                            Else
                                jday = jday - 30
                                If jday <= 31 Then 'Day of July
                                    JulMonth = 7
                                Else
                                    jday = jday - 31
                                    If jday <= 31 Then 'Day of August
                                        JulMonth = 8
                                    Else
                                        jday = jday - 31
                                        If jday <= 30 Then 'Day of September
                                            JulMonth = 9
                                        Else
                                            jday = jday - 30
                                            If jday <= 31 Then 'Day of October
                                                JulMonth = 10
                                            Else
                                                jday = jday - 31
                                                If jday <= 30 Then 'Day of November
                                                    JulMonth = 11
                                                Else 'Day of December
                                                    jday = jday - 30
                                                    JulMonth = 12
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If

        fmt1 = formatString.Substring(0, formatString.IndexOf("j"c))
        fmt2 = formatString.Substring(formatString.LastIndexOf("j"c))
        date1 = dateString.Substring(0, formatString.IndexOf("j"c))
        date2 = dateString.Substring(formatString.LastIndexOf("j"c))

        'Prepare the date string for parsing
        newdate = date1 & JulMonth.ToString("00", Globalization.CultureInfo.InvariantCulture) & "-" & jday.ToString("00", Globalization.CultureInfo.InvariantCulture) & date2
        newfmt = fmt1 & "MM-dd" & fmt2

        If Date.TryParseExact(newdate, newfmt, Globalization.CultureInfo.InvariantCulture, Globalization.DateTimeStyles.None, resultDate) Then
            'Time parsed
            Return True
        Else
            'Define actual time as dummy value
            resultDate = Date.MinValue
            Return False
        End If

    End Function

    ''' <summary>
    ''' Convert the Year and day of the year in a date object.
    ''' </summary>
    ''' <param name="year">Year</param>
    ''' <param name="dayOfTheYear">Day of the year</param>
    ''' <param name="hours">Hours</param>
    ''' <param name="minutes">Minutes</param>
    ''' <param name="seconds">Seconds</param>
    ''' <param name="milliSecs">MilliSeconds</param>
    ''' <param name="resultDate">Parsed Date</param>
    ''' <returns></returns>
    Public Function ParseJulianDate(year As Integer, dayOfTheYear As Integer, hours As Integer, minutes As Integer, seconds As Integer, milliSecs As Integer, ByRef resultDate As Date) As Boolean
        Dim LeapYear As Boolean = Date.IsLeapYear(year)
        Dim FebDays As Integer = If(LeapYear, 29, 28)
        Dim JulMonth As Integer
        Dim ShotTimeString As String

        If dayOfTheYear <= 31 Then 'Day of January
            JulMonth = 1
        Else
            dayOfTheYear = dayOfTheYear - 31
            If dayOfTheYear <= FebDays Then 'Day of February
                JulMonth = 2
            Else
                dayOfTheYear = dayOfTheYear - FebDays
                If dayOfTheYear <= 31 Then 'Day of March
                    JulMonth = 3
                Else
                    dayOfTheYear = dayOfTheYear - 31
                    If dayOfTheYear <= 30 Then 'Day of April
                        JulMonth = 4
                    Else
                        dayOfTheYear = dayOfTheYear - 30
                        If dayOfTheYear <= 31 Then 'Day of May
                            JulMonth = 5
                        Else
                            dayOfTheYear = dayOfTheYear - 31
                            If dayOfTheYear <= 30 Then 'Day of June
                                JulMonth = 6
                            Else
                                dayOfTheYear = dayOfTheYear - 30
                                If dayOfTheYear <= 31 Then 'Day of July
                                    JulMonth = 7
                                Else
                                    dayOfTheYear = dayOfTheYear - 31
                                    If dayOfTheYear <= 31 Then 'Day of August
                                        JulMonth = 8
                                    Else
                                        dayOfTheYear = dayOfTheYear - 31
                                        If dayOfTheYear <= 30 Then 'Day of September
                                            JulMonth = 9
                                        Else
                                            dayOfTheYear = dayOfTheYear - 30
                                            If dayOfTheYear <= 31 Then 'Day of October
                                                JulMonth = 10
                                            Else
                                                dayOfTheYear = dayOfTheYear - 31
                                                If dayOfTheYear <= 30 Then 'Day of November
                                                    JulMonth = 11
                                                Else 'Day of December
                                                    dayOfTheYear = dayOfTheYear - 30
                                                    JulMonth = 12
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        'Prepare the date string for parsing
        ShotTimeString = year.ToString("0000", Globalization.CultureInfo.InvariantCulture) & "-"
        ShotTimeString = ShotTimeString & JulMonth.ToString("00", Globalization.CultureInfo.InvariantCulture) & "-"
        ShotTimeString = ShotTimeString & dayOfTheYear.ToString("00", Globalization.CultureInfo.InvariantCulture) & " "
        ShotTimeString = ShotTimeString & hours.ToString("00", Globalization.CultureInfo.InvariantCulture) & ":"
        ShotTimeString = ShotTimeString & minutes.ToString("00", Globalization.CultureInfo.InvariantCulture) & ":"
        ShotTimeString = ShotTimeString & seconds.ToString("00", Globalization.CultureInfo.InvariantCulture) & "."
        ShotTimeString = ShotTimeString & milliSecs.ToString("000", Globalization.CultureInfo.InvariantCulture)

        If Date.TryParseExact(ShotTimeString, "yyyy-MM-dd HH:mm:ss.fff", Globalization.CultureInfo.InvariantCulture,
                              Globalization.DateTimeStyles.AssumeLocal, resultDate) Then
            'Time parsed
            Return True
        Else
            'Define actual time as dummy value
            resultDate = Date.MinValue
            Return False
        End If

    End Function

    ''' <summary>
    ''' Convert a date object in a formatted string with year, day of the year and time.
    ''' </summary>
    ''' <param name="dateValue">Date to convert</param>
    ''' <returns></returns>
    Public Function FormatJulianDate(dateValue As Date, formatString As String) As String
        If formatString Is Nothing Then Return ""
        Dim ResultString, fmt1, fmt2 As String
        Dim countJ As Integer = CountCharacter(formatString, "j"c)
        If countJ > 0 Then
            fmt1 = formatString.Substring(0, formatString.IndexOf("j"c))
            fmt2 = formatString.Substring(formatString.LastIndexOf("j"c) + 1)
            ResultString = If(fmt1.Length > 0, dateValue.ToString(fmt1, Globalization.CultureInfo.InvariantCulture), "")
            ResultString = ResultString & dateValue.DayOfYear.ToString("000", Globalization.CultureInfo.InvariantCulture)
            ResultString = ResultString & If(fmt2.Length > 0, dateValue.ToString(fmt2, Globalization.CultureInfo.InvariantCulture), "")
        Else
            ResultString = dateValue.ToString(formatString, Globalization.CultureInfo.InvariantCulture)
        End If
        Return ResultString
    End Function

End Module