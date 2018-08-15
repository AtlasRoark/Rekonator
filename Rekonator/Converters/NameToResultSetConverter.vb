Imports System.Globalization
Imports Rekonator.ResultGroup

Public Class NameToResultSetConverter
    Implements IValueConverter

    Private Function Convert(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.Convert
        Dim resultSets As Dictionary(Of ResultSetType, ResultSet) = TryCast(value, Dictionary(Of ResultSetType, ResultSet))
        If resultSets Is Nothing Then
            Return Nothing
        ElseIf resultSets.ContainsKey(CInt(parameter)) Then
            Return resultSets(CInt(parameter))
        Else
            Return nothing
        End If

    End Function

    Private Function ConvertBack(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.ConvertBack
        Throw New NotSupportedException()
    End Function
End Class

