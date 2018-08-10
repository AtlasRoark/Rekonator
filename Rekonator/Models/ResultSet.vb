Imports System.Data

Public Class ResultSet
    Public Property ResultSet As ResultSetName
    Public Property ResultDataView As DataView = Nothing
    Public Property RecordCount As Integer = 0
    Public Property SQLCommand As String = String.Empty

    Public Sub New(resultSetName As ResultSetName)
        ResultSet = resultSetName
    End Sub
    Public Enum ResultSetName
        Left
        Right
        Differ
        Match
    End Enum
End Class

