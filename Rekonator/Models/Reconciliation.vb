<Serializable()>
Public Class Reconciliation
    Property ReconciliationName As String
    Property LeftReconSource As ReconSource
    Property RightReconSource As ReconSource
    Property CompletenessComparisions As List(Of Comparision)
    Property MatchingComparisions As List(Of Comparision)

    Public Shared Reconciliations As New List(Of Reconciliation)
    Public Shared Sub Add(reconciliationName As String,
                          leftDataSource As ReconSource,
                          rightDataSource As ReconSource,
                          completenessComparision As List(Of Comparision),
                          matchingComparision As List(Of Comparision))
        Reconciliations.Add(New Reconciliation With {
                            .ReconciliationName = reconciliationName,
                            .LeftReconSource = leftDataSource,
                            .RightReconSource = rightDataSource,
                            .CompletenessComparisions = completenessComparision,
                            .MatchingComparisions = matchingComparision}
                            )
    End Sub
    Public Shared Function GetReconciliation(reconciliationName As String) As Reconciliation
        Return Reconciliations.Where(Function(w) w.ReconciliationName = reconciliationName).FirstOrDefault
    End Function
End Class
