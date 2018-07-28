<Serializable()>
Public Class AggregateOperation
    Property SourceColumn As String
    Property AggregateColumn As String
    Property Operation As AggregateFunction

    Public Enum AggregateFunction
        Count
        Sum
        Avg
        Min
        Max
        First
        Last
    End Enum

    Public Sub New()

    End Sub

End Class
