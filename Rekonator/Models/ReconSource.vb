﻿Imports System.Data

<Serializable()>
Public Class ReconSource
    Property ReconDataSource As DataSource
    Property ReconTable As String
    Property IsLoaded As Boolean
    Property Parameters As Dictionary(Of String, String)
    Property LoadedSet As DataTable
    Property Where As String
    Property Fields As String()
    Property Types As String()
    Property Aggregations As List(Of Aggregate)

    Public Shared Function GetSelect(reconSource As ReconSource) As String
        Dim selectCommand As String = $"SELECT * FROM {reconSource.ReconTable} x"
        If Not String.IsNullOrWhiteSpace(reconSource.Where) Then
            selectCommand += $" WHERE {reconSource.Where.Replace("x!.", "x.")}"
        End If
        Return selectCommand
    End Function
End Class