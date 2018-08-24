Public Module Parsers

    Public Function DoubleParseOrDefault(value As String, culture As IFormatProvider, defaultValue As Double) As Double
        Dim AnyNumb As Globalization.NumberStyles = Globalization.NumberStyles.Any
        Dim tmpValue As Double

        If Double.TryParse(value, AnyNumb, culture, tmpValue) Then
            Return tmpValue
        Else
            Return defaultValue
        End If

    End Function

    Public Function IntegerParseOrDefault(value As String, culture As IFormatProvider, defaultValue As Integer) As Integer
        Dim AnyNumb As Globalization.NumberStyles = Globalization.NumberStyles.Any
        Dim tmpValue As Integer

        If Integer.TryParse(value, AnyNumb, culture, tmpValue) Then
            Return tmpValue
        Else
            Return defaultValue
        End If

    End Function

End Module
