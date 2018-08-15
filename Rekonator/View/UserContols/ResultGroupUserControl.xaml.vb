Public Class ResultGroupUserControl
    Inherits UserControl

#Region "Dependency Properties"

    Public Property ResultGroup As ResultGroup
        Get
            Return GetValue(ResultGroupProperty)
        End Get

        Set(ByVal value As ResultGroup)
            SetValue(ResultGroupProperty, value)
        End Set
    End Property

    Public Shared ReadOnly ResultGroupProperty As DependencyProperty =
                           DependencyProperty.Register("ResultGroup",
                           GetType(ResultGroup), GetType(ResultGroupUserControl),
                           New PropertyMetadata(Nothing))


    Public Property HasLoaded As Boolean
        Get
            Return GetValue(HasLoadedProperty)
        End Get

        Set(ByVal value As Boolean)
            SetValue(HasLoadedProperty, value)
        End Set
    End Property

    Public Shared ReadOnly HasLoadedProperty As DependencyProperty =
                           DependencyProperty.Register("HasLoaded",
                           GetType(Boolean), GetType(ResultGroupUserControl),
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
