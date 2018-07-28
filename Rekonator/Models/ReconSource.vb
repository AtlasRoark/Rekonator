Imports System.Data

<Serializable()>
Public Class ReconSource
    Property ReconDataSource As DataSource
    Property ReconTable As String
    Property IsLoaded As Boolean
    Property Parameters As Dictionary(Of String, String)
    Property LoadedSet As DataTable
    Property Filters As List(Of Filter)
    Property Aggregations As List(Of Aggregate)

End Class