Partial Class DataSourceUserContol
    Inherits UserControl
    'https://blog.scottlogic.com/2012/02/06/a-simple-pattern-for-creating-re-useable-usercontrols-in-wpf-silverlight.html

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

    Private Sub ButtonLoad_Click(sender As Object, e As RoutedEventArgs)
        Dim mainWindow As MainWindow = Utility.FindAncestor(Me, GetType(MahApps.Metro.Controls.MetroWindow))
        mainWindow.LoadReconSource(Side)
    End Sub

    Private Sub DataGridParameters_PreviewKeyDown(sender As Object, e As KeyEventArgs)
        If (e.Key = Key.Enter) Or
           (e.Key = Key.V AndAlso (Keyboard.Modifiers And ModifierKeys.Control) = ModifierKeys.Control) Then

            Dim datagrid As DataGrid = TryCast(sender, DataGrid)
            If datagrid IsNot Nothing Then
                Dim textbox As TextBox = TryCast(e.OriginalSource, TextBox)
                If textbox IsNot Nothing Then
                    If e.Key = Key.E Then
                        textbox.Text = textbox.Text + vbCrLf
                    Else
                        Dim clipBoardData As String = Clipboard.GetData(DataFormats.UnicodeText)
                        If textbox.SelectedText Is Nothing Then
                            textbox.Text.Insert(textbox.CaretIndex, clipBoardData)
                        Else
                            textbox.Text = textbox.Text.Replace(textbox.SelectedText, clipBoardData)

                        End If
                    End If
                    e.Handled = True
                End If
            End If
        End If
    End Sub

    Private Sub CanPaste(ByVal sender As Object, ByVal e As CanExecuteRoutedEventArgs)
        e.CanExecute = (Clipboard.GetData(DataFormats.UnicodeText) IsNot Nothing)
        e.Handled = True
    End Sub

    Private Sub Paste(ByVal sender As Object, ByVal e As ExecutedRoutedEventArgs)
        'Pasting in multiline text just adds the first line.  Fix with custom paste.
        Dim clipBoardData As String = Clipboard.GetData(DataFormats.UnicodeText)
        If Not String.IsNullOrWhiteSpace(clipBoardData) Then
            Dim datagrid As DataGrid = TryCast(sender, DataGrid)
            If datagrid IsNot Nothing Then
                Dim textbox As TextBox = TryCast(e.OriginalSource, TextBox)
                If textbox IsNot Nothing Then

                    e.Handled = True
                End If
            End If
        End If
    End Sub
End Class

