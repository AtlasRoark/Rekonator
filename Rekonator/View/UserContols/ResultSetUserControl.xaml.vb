Imports System.Data

Partial Public Class ResultSetUserControl
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
                           GetType(ResultSet), GetType(ResultSetUserControl),
                           New PropertyMetadata(Nothing))

#End Region
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
    End Sub

    Private Sub DataGridRow_MouseRightButtonDown(sender As Object, e As MouseButtonEventArgs)
        Dim dgc As DataGridCell = TryCast(sender, DataGridCell)
        If dgc Is Nothing Then Exit Sub

        Dim dgr As DataGridRow = Utility.FindAncestor(dgc, GetType(DataGridRow))
        If dgr Is Nothing Then Exit Sub

        Dim dr As DataRow = TryCast(dgr.Item.Row, DataRow)
        If dr Is Nothing Then Exit Sub

        Dim dg As DataGrid = Utility.FindAncestor(dgc, GetType(DataGrid))
        If dg Is Nothing Then Exit Sub

        Dim rsuc As ResultSetUserControl = TryCast(dg.DataContext, ResultSetUserControl)
        If rsuc Is Nothing Then Exit Sub

        Dim rguc As ResultGroupUserControl = TryCast(rsuc.DataContext, ResultGroupUserControl)
        If rguc Is Nothing Then Exit Sub

        'Dim rgt As ResultGroup.ResultGroupType = TryCast([Enum].Parse(GetType(ResultGroup.ResultGroupType), rguc.ResultGroup), ResultGroup.ResultGroupType)

        Dim mainWindow As MainWindow = Utility.FindAncestor(Me, GetType(MahApps.Metro.Controls.MetroWindow))
        If mainWindow Is Nothing Then Exit Sub

        Dim columns As List(Of String) = dg.Columns.Select(Function(s) s.Header.ToString).ToList
        mainWindow.DrillDownRekonate(rguc.ResultGroup.ResultGroupName, dr, columns)
    End Sub

    Private Sub DataGridCell_PreviewKeyDown(sender As Object, e As KeyEventArgs)
        Beep()
    End Sub
End Class
