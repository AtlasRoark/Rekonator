Public Class ResultGroup
    Public Property ResultGroupName As ResultGroupType
    Public Property ResultSets As New Dictionary(Of ResultSetType, ResultSet)

    Public Sub New(resultGroupName As ResultGroupType)
        Me.ResultGroupName = resultGroupName
    End Sub
    Public Enum ResultGroupType
        Left
        Right
        Differ
        Match
    End Enum

    Public Enum ResultSetType
        Loaded
        Result
        DrillDown
    End Enum
End Class

