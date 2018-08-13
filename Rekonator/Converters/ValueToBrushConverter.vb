Imports System.Globalization

Public Class ValueToBrushConverter
    Implements IValueConverter

    Private Function Convert(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.Convert
        Dim cell As String
        cell = TryCast(value, String)
        If IsNothing(cell) Then Exit Function
        If cell.Contains("<>") Then
            If parameter = "F" Then
                Return Brushes.White
            Else
                Return Application.Current.MainWindow.FindResource("RekonatorErrorBrush")
            End If
        End If
        Return DependencyProperty.UnsetValue
    End Function

    Private Function ConvertBack(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.ConvertBack
        Throw New NotSupportedException()
    End Function

End Class
