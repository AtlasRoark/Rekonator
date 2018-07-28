<Serializable()>
Public Class Aggregate
    Property DataSourceName As String
    Property GroupByColumns As String()
    Property AggregateOperations As List(Of AggregateOperation)
End Class

