<Serializable()>
Public Class DataSource
    Property DataSourceName As String = String.Empty
    Property ImageSource As String = String.Empty
    Property IsSlowLoading As Boolean = False

    Public Shared Property DataSources As New List(Of DataSource)
    Public Shared Sub Add(dataSourceName As String, imageSource As String, isSlowLoading As Boolean)
        DataSources.Add(New DataSource With {.DataSourceName = dataSourceName, .ImageSource = imageSource, .IsSlowLoading = isSlowLoading})
    End Sub
    Public Shared Function GetDataSource(dataSourceName As String) As DataSource
        Return DataSources.Where(Function(w) w.DataSourceName = dataSourceName).FirstOrDefault
    End Function

End Class


