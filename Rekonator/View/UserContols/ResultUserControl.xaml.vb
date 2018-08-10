Public Class ResultUserControl
    Inherits UserControl

#Region "Dependency Properties"

    Public Property ResultSet As ResultSet
        Get
            Return GetValue(ResultSetProperty)
        End Get

        Set(ByVal value As ResultSet)
            SetValue(ResultSetProperty, value)
        End Set
    End Property

    Public Shared ReadOnly ResultSetProperty As DependencyProperty =
                           DependencyProperty.Register("ResultSet",
                           GetType(ResultSet), GetType(ResultUserControl),
                           New PropertyMetadata(Nothing))
#End Region

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
    End Sub

    Private Sub DataGridCell_PreviewKeyDown(sender As Object, e As KeyEventArgs)

    End Sub
End Class
