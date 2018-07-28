<Serializable()>
Public Class DataSource
    Property DataSourceName As String
    Property ImageSource As String

    Public Shared DataSources As New List(Of DataSource)
    Public Shared Sub Add(dataSourceName As String, imageSource As String)
        DataSources.Add(New DataSource With {.DataSourceName = dataSourceName, .ImageSource = imageSource})
    End Sub
    Public Shared Function GetDataSource(dataSourceName As String) As DataSource
        Return DataSources.Where(Function(w) w.DataSourceName = dataSourceName).FirstOrDefault
    End Function

End Class


