Partial Public Class SolutionUserControl
    Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Sub CBReconciliation_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        Dim mainWindow As MainWindow = Utility.FindAncestor(Me, GetType(MahApps.Metro.Controls.MetroWindow))
        If mainWindow.IsLoaded Then
            Dim cb As ComboBox = TryCast(sender, ComboBox)
            If cb IsNot Nothing Then
                Dim rc As Reconciliation = TryCast(cb.SelectedItem, Reconciliation)
                If rc IsNot Nothing Then
                    If MainWindow IsNot Nothing Then MainWindow.ChangeReconciliation(rc)
                End If
            End If
        End If

    End Sub

End Class
