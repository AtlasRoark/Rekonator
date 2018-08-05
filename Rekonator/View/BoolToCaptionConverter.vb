Imports System.Globalization

Public Class BoolToCaptionConverter
    Implements IValueConverter

    Private Function Convert(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.Convert
        Dim isLoaded As Boolean = value
        If isLoaded Then
            Return "Reload"
        Else
            Return "Load"
        End If
    End Function

    Private Function ConvertBack(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.ConvertBack
        Throw New NotSupportedException()
    End Function
End Class
