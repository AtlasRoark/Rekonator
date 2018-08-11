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

    Public Property MessageLog As ObservableCollection(Of MessageEntry)
        Get
            MessageLog = _messages
        End Get
        Set(value As ObservableCollection(Of MessageEntry))
            OnPropertyChanged("MessageLog")
        End Set
    End Property
    Private _messages As New ObservableCollection(Of MessageEntry)
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
                'LoadReconSources(_reconciliation)
            End If
            OnPropertyChanged("Reconciliation")
        End Set
    End Property
    Private _reconciliation As Reconciliation
#End Region

#Region "-- Result Set Properties --"

    Public Property LeftResultSet As ResultSet
        Get
            LeftResultSet = _LeftResultSet
        End Get
        Set(value As ResultSet)
            _LeftResultSet = value
            OnPropertyChanged("LeftResultSet")
        End Set
    End Property
    Private _leftResultSet As ResultSet

    Public Property RightResultSet As ResultSet
        Get
            RightResultSet = _rightResultSet
        End Get
        Set(value As ResultSet)
            _rightResultSet = value
            OnPropertyChanged("RightResultSet")
        End Set
    End Property
    Private _rightResultSet As ResultSet

    Public Property DifferResultSet As ResultSet
        Get
            DifferResultSet = _differResultSet
        End Get
        Set(value As ResultSet)
            _differResultSet = value
            OnPropertyChanged("DifferResultSet")
        End Set
    End Property
    Private _differResultSet As ResultSet

    Public Property MatchResultSet As ResultSet
        Get
            MatchResultSet = _matchResultSet
        End Get
        Set(value As ResultSet)
            _matchResultSet = value
            OnPropertyChanged("MatchResultSet")
        End Set
    End Property

    Public Sub ClearMessageLog()
        _messages.Clear()
        MessageLog = _messages
    End Sub

    Private _matchResultSet As ResultSet

    'Public Property ResultSets As ObservableCollection(Of ResultSet)
    '    Get
    '        ResultSets = _resultSets
    '    End Get
    '    Set(value As ObservableCollection(Of ResultSet))
    '        _resultSets = value
    '        OnPropertyChanged("ResultSets")
    '    End Set
    'End Property
    'Private _resultSets As ObservableCollection(Of ResultSet)

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

