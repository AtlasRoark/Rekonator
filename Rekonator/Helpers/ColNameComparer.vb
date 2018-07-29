Imports System.Data
Imports System.Globalization

'Used to as .Distinct(New ColNameComparer)
Public Class ColNameComparer
    Implements IEqualityComparer(Of DataColumn)

    Public Function GetHashCode(obj As DataColumn) As Integer Implements IEqualityComparer(Of DataColumn).GetHashCode
        Return obj.GetHashCode()
    End Function

    Public Function Equals(x As DataColumn, y As DataColumn) As Boolean Implements IEqualityComparer(Of DataColumn).Equals
        If CultureInfo.CurrentCulture.CompareInfo.Compare(x.ColumnName, y.ColumnName, CompareOptions.IgnoreCase) = 0 Then
            Return True
        Else
            Return False

        End If
    End Function

End Class