Partial Class DataSourceUserContol

    Private _isLoaded As Boolean = False

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

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
        _isLoaded = True
    End Sub

    Private Sub BTN_Load_Click(sender As Object, e As RoutedEventArgs)
        Dim reconSource As ReconSource = Me.DataContext
        Dim mainWindow As MainWindow = Utility.FindAncestor(Me, GetType(MahApps.Metro.Controls.MetroWindow))

        mainWindow.btnLeft_Click(sender, e)

        'vm.LoadReconSources()
    End Sub
End Class

