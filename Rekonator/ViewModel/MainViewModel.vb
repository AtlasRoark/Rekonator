Imports System.Collections.ObjectModel
Imports System.Data

Public Class MainViewModel
    Inherits ViewModelBase
    'Inherits ObservableCollection(Of Solution)
    'Implements INotifyPropertyChanged

#Region "-- App Properties --"
    Public Property MainWindow As MainWindow

    Public Property DataSources As List(Of DataSource)
        Get
            DataSources = DataSource.DataSources
        End Get
        Set(value As List(Of DataSource))
            DataSource.DataSources = value
            OnPropertyChanged("DataSources")
        End Set
    End Property

    'Private _dataSources As List(Of DataSource)
    Public Property Translations As List(Of Translation)
        Get
            Translations = _translations
        End Get
        Set(value As List(Of Translation))
            _translations = value
            'OnPropertyChanged("Translations")
        End Set
    End Property
    Private _translations As List(Of Translation)

    Public Property CompareMethods As List(Of CompareMethod)
        Get
            CompareMethods = _compareMethods
        End Get
        Set(value As List(Of CompareMethod))
            _compareMethods = value
        End Set
    End Property
    Private _compareMethods As List(Of CompareMethod)

    Public Property MessageLog As List(Of MessageEntry)
        Get
            MessageLog = _messages
        End Get
        Set(value As List(Of MessageEntry))
            OnPropertyChanged("MessageLog")
        End Set
    End Property
    Private _messages As New List(Of MessageEntry)
#End Region
#Region "-- Solution Model Properties --"
    'Surface parent properties here so two way binding works
    Public Property Solution As Solution 'active solution
        Get
            Solution = _solution
        End Get
        Set(value As Solution)
            _solution = value
            If _solution IsNot Nothing Then
                Reconciliations = New ObservableCollection(Of Reconciliation)(_solution.Reconciliations)
                Reconciliation = _solution.Reconciliations.FirstOrDefault
            End If
            OnPropertyChanged("Solution")
        End Set
    End Property
    Private _solution As Solution

    Public Property Reconciliations As ObservableCollection(Of Reconciliation)  ' active reconciliation from _solution
        Get
            Reconciliations = _reconciliations
        End Get
        Set(value As ObservableCollection(Of Reconciliation))
            _reconciliations = value
            'If _reconciliation IsNot Nothing Then
            '    LeftReconSource = _reconciliation.LeftReconSource
            '    RightReconSource = _reconciliation.RightReconSource
            'End If
            OnPropertyChanged("Reconciliations")
        End Set
    End Property
    Private _reconciliations As ObservableCollection(Of Reconciliation)


    Public Property Reconciliation As Reconciliation  ' active reconciliation from _solution
        Get
            Reconciliation = _reconciliation
        End Get
        Set(value As Reconciliation)
            _reconciliation = value
            If _reconciliation IsNot Nothing Then
                'LeftReconSource = _reconciliation.LeftReconSource
                'RightReconSource = _reconciliation.RightReconSource
            End If
            OnPropertyChanged("Reconciliation")
        End Set
    End Property
    Private _reconciliation As Reconciliation

    'Public Property NewReconciliation As String
    '    Get
    '        NewReconciliation = _newReconciliation
    '    End Get
    '    Set(value As String)
    '        _newReconciliation = value
    '        If _reconciliation Is Nothing Then
    '        End If
    '    End Set
    'End Property
    'Private _newReconciliation As String



#End Region
#Region "-- Result Set Properties --"

    Public Property LeftSet As DataView
        Get
            LeftSet = _leftSet
        End Get
        Set(value As DataView)
            _leftSet = value
            'OnPropertyChanged("LeftSet")
        End Set
    End Property
    Private _leftSet As New DataView

    Public Property RightSet As DataView
        Get
            RightSet = _rightSet
        End Get
        Set(value As DataView)
            _rightSet = value
            'OnPropertyChanged("RightSet")
        End Set
    End Property
    Private _rightSet As New DataView

    Public Property DifferSet As DataView
        Get
            DifferSet = _differSet
        End Get
        Set(value As DataView)
            _differSet = value
            'OnPropertyChanged("DifferSet")
        End Set
    End Property
    Private _differSet As New DataView

    Public Property MatchSet As DataView
        Get
            MatchSet = _matchSet
        End Get
        Set(value As DataView)
            _matchSet = value
            'OnPropertyChanged("MatchSet")
        End Set
    End Property
    Private _matchSet As New DataView
#End Region

#Region "-- Notify Property Change --"
    'Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    'Public Sub OnPropertyChanged(propertyName As String)
    '    Me.CheckPropertyName(propertyName)
    '    RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propertyName))
    'End Sub

    '<Conditional("DEBUG")>
    '<DebuggerStepThrough>
    'Public Sub CheckPropertyName(propertyName As String)
    '    If TypeDescriptor.GetProperties(Me)(propertyName) Is Nothing Then
    '        Throw New Exception($"Could not find property: {propertyName}")
    '    End If
    'End Sub
#End Region

#Region "Commands"
    Public Sub LoadReconSources()
        'If Not LeftReconSource.IsLoaded Then LoadReconSource(LeftReconSource)
        'If Not RightReconSource.IsLoaded Then LoadReconSource(RightReconSource)

        Using sql As New SQL
            'Dim rs As ReconSource = _vmLeft.ReconSource ' LeftReconSource
            '_left = sql.GetDataTable(ReconSource.GetSelect(rs)) ', _reconciliation.FromDate, _reconciliation.ToDate)
            'LeftSet = _left.AsDataView
            'rs = _rightReconSource
            '_right = sql.GetDataTable(ReconSource.GetSelect(rs)) ', _reconciliation.FromDate, _reconciliation.ToDate)
            'RightSet = _right.AsDataView
        End Using
    End Sub


#End Region

End Class


'Public Class MockMainViewModel

'    Public ReadOnly Property Solution As Solution 'active solution
'        Get
'            'Dim _solution As Solution
'            'Using m As New Mock
'            '    DataSources = m.MockLoadDataSources
'            '    _solution = Task.Run(Function() m.MockLoadSolutionAsync(1)).GetAwaiter().GetResult() 'Model for Solution
'            '    Reconciliation = _solution.Reconciliations(0)
'            '    LeftReconSource = _solution.Reconciliations(0).LeftReconSource
'            'End Using
'            'Solution = _solution
'        End Get
'    End Property
'    Public Property Reconciliation As Reconciliation
'    Public Property LeftReconSource As ReconSource
'    Public Shared ReadOnly Property DataSources As List(Of DataSource)
'        Get
'            DataSources.Add(New DataSource With {.DataSourceName = "Excel"})
'            DataSources.Add(New DataSource With {.DataSourceName = "Intact"})
'            DataSources.Add(New DataSource With {.DataSourceName = "QuickBooks"})
'            DataSources.Add(New DataSource With {.DataSourceName = "ServiceTitan"})
'            DataSources.Add(New DataSource With {.DataSourceName = "SQL"})
'        End Get
'    End Property

'End Class

