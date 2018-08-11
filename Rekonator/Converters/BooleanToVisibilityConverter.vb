Imports System.Data
Imports System.Globalization

Public Class BooleanToVisibilityConverter
    Implements IValueConverter

    Private Function Convert(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.Convert
        Dim isBoolean As Boolean
        If Boolean.TryParse(value, isBoolean) Then
            If isBoolean Then
                Return Visibility.Visible
            Else
                Return Visibility.Collapsed
            End If
        End If
    End Function

    Private Function ConvertBack(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.ConvertBack
        Throw New NotSupportedException()
    End Function
End Class
