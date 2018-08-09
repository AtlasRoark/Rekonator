Partial Class DataSourceUserContol
    Inherits UserControl

#Region "Dependency Properties"
    Public Property DataSources As List(Of DataSource)
        Get
            Return GetValue(DataSourcesProperty)
        End Get

        Set(ByVal value As List(Of DataSource))
            SetValue(DataSourcesProperty, value)
        End Set
    End Property

    Public Shared ReadOnly DataSourcesProperty As DependencyProperty =
                           DependencyProperty.Register("DataSources",
                           GetType(List(Of DataSource)), GetType(DataSourceUserContol),
                           New PropertyMetadata(Nothing))


    Public Property Side As ReconSource.SideName
        Get
            Return GetValue(SideProperty)
        End Get

        Set(ByVal value As ReconSource.SideName)
            SetValue(SideProperty, value)
        End Set
    End Property

    Public Shared ReadOnly SideProperty As DependencyProperty =
                           DependencyProperty.Register("Side",
                           GetType(ReconSource.SideName), GetType(DataSourceUserContol),
                           New PropertyMetadata(Nothing))


    Public Property ReconSource As ReconSource
        Get
            Return GetValue(ReconSourceProperty)
        End Get

        Set(ByVal value As ReconSource)
            SetValue(ReconSourceProperty, value)
        End Set
    End Property


    Public Shared ReadOnly ReconSourceProperty As DependencyProperty =
                           DependencyProperty.Register("ReconSource",
                           GetType(ReconSource), GetType(DataSourceUserContol),
                           New PropertyMetadata(Nothing))
#End Region

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        'GridReconSource.DataContext = Me
        ' Add any initialization after the InitializeComponent() call.
    End Sub
    Private Sub CBDataSources_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        'If Not IsLoaded Then Exit Sub
        'Dim cBox As ComboBox = TryCast(sender, ComboBox)
        'Dim appVM As AppViewModel = TryCast(cBox.DataContext, AppViewModel)
        'Dim reconSource As ReconSource = TryCast(Me.DataContext, ReconSource)
        'If reconSource IsNot Nothing AndAlso e.AddedItems.Count > 0 Then reconSource.ReconDataSource = e.AddedItems(0)
        'Dim userControl As UserControl = Utility.FindAncestor(cBox, GetType(UserControl))
        'If reconSource.InstantiatedSide = ReconSource.Side.Left Then
        'Window.Reconciliation.LeftReconSource.ReconDataSource = e.AddedItems(0)
        'appVM.MainWindow.Solution = appVM.MainWindow.Solution
        'Else
        'appVM.MainWindow.Reconciliation.RightReconSource.ReconDataSource = e.AddedItems(0)

        'End If

    End Sub





    Private Sub DataSourcePage_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        'Dim userControl As UserControl = TryCast(sender, UserControl)
        'If userControl IsNot Nothing Then
        '    Dim mainWindow As MainWindow = Utility.FindAncestor(Me, GetType(MahApps.Metro.Controls.MetroWindow))

        'End If
    End Sub

    Private Sub ButtonLoad_Click(sender As Object, e As RoutedEventArgs)
        Dim mainWindow As MainWindow = Utility.FindAncestor(Me, GetType(MahApps.Metro.Controls.MetroWindow))
        mainWindow.LoadReconSource(Side)
    End Sub
End Class

