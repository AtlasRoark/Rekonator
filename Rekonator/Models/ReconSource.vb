Imports System.Data

<Serializable()>
Public Class ReconSource
    Property ReconDataSource As New DataSource
    Property ReconTable As String = String.Empty
    Property IsLoaded As Boolean
    Property Parameters As New List(Of Parameter)
    Property LoadedSet As DataTable
    Property WhereClause As String = String.Empty
    Property Columns As New List(Of Column)
    Property Aggregations As New List(Of Aggregate)

    Public Shared Function GetSelect(reconSource As ReconSource) As String
        Dim selectCommand As String = $"SELECT * FROM {reconSource.ReconTable} x"
        If Not String.IsNullOrWhiteSpace(reconSource.WhereClause) Then
            selectCommand += $" WHERE {reconSource.WhereClause.Replace("x!.", "x.")}"
        End If
        Return selectCommand
    End Function
    Public Enum SideName
        Left
        Right
    End Enum
End Class