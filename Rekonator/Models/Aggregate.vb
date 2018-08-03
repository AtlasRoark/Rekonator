<Serializable()>
Public Class Aggregate
    Property GroupByColumns As String()
    Property AggregateOperations As New List(Of AggregateOperation)
End Class

