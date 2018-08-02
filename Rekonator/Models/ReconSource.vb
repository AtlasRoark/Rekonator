Imports System.Data

<Serializable()>
Public Class ReconSource
    Property ReconDataSource As DataSource
    Property ReconTable As String
    Property IsLoaded As Boolean
    Property Parameters As List(Of Parameter) ' Dictionary(Of String, String)
    Property LoadedSet As DataTable
    Property WhereClause As String
    'Property Fields As String()
    Property Types As String()
    Property Columns As List(Of Column)
    Property Aggregations As List(Of Aggregate)
    Property InstantiatedSide As Side

    Public Shared Function GetSelect(reconSource As ReconSource) As String
        Dim selectCommand As String = $"SELECT * FROM {reconSource.ReconTable} x"
        If Not String.IsNullOrWhiteSpace(reconSource.WhereClause) Then
            selectCommand += $" WHERE {reconSource.WhereClause.Replace("x!.", "x.")}"
        End If
        Return selectCommand
    End Function

    Public Enum Side
        Left
        Rigth
    End Enum
End Class