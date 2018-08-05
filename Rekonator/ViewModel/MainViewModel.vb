Imports System.Collections.ObjectModel
Imports System.ComponentModel

Public Class MainViewModel
    Inherits ObservableCollection(Of Solution)
    Implements INotifyPropertyChanged

#Region "-- App Properties --"
    Public Property DataSources As List(Of DataSource)
        Get
            DataSources = DataSource.DataSources
        End Get
        Set(value As List(Of DataSource))
            DataSource.DataSources = value
            'OnPropertyChanged("DataSources")
        End Set
    End Property

    Public Property SelectedDataSource As DataSource
    Public Property MainWindow As MainWindow

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
            'OnPropertyChanged("MessageLog")
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
            Reconciliation = _solution.Reconciliations(0)
            OnPropertyChanged("Solution")
        End Set
    End Property
    Private _solution As Solution

    Public Property Reconciliation As Reconciliation  ' active reconciliation from _solution
        Get
            Reconciliation = _reconciliation
        End Get
        Set(value As Reconciliation)
            _reconciliation = value
            LeftReconSource = _reconciliation.LeftReconSource
            RightReconSource = _reconciliation.RightReconSource
            OnPropertyChanged("Reconciliation")
        End Set
    End Property
    Private _reconciliation As Reconciliation

    Public Property LeftReconSource As ReconSource  ' active reconciliation from _solution
        Get
            LeftReconSource = _leftReconSource
        End Get
        Set(value As ReconSource)
            _leftReconSource = value
            OnPropertyChanged("LeftReconSource")
        End Set
    End Property
    Private _leftReconSource As ReconSource

    Public Property RightReconSource As ReconSource  ' active reconciliation from _solution
        Get
            RightReconSource = _rightReconSource
        End Get
        Set(value As ReconSource)
            _rightReconSource = value
            OnPropertyChanged("RightReconSource")
        End Set
    End Property
    Private _rightReconSource As ReconSource
#End Region

#Region "-- Notify Property Change --"
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    Public Sub OnPropertyChanged(propertyName As String)
        Me.CheckPropertyName(propertyName)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propertyName))
    End Sub

    <Conditional("DEBUG")>
    <DebuggerStepThrough>
    Public Sub CheckPropertyName(propertyName As String)
        If TypeDescriptor.GetProperties(Me)(propertyName) Is Nothing Then
            Throw New Exception($"Could not find property: {propertyName}")
        End If
    End Sub
#End Region

#Region "Commands"
    Public Sub LoadReconSources()
        If Not LeftReconSource.IsLoaded Then LoadReconSource(LeftReconSource)
        If Not RightReconSource.IsLoaded Then LoadReconSource(RightReconSource)

        Using sql As New SQL
            Dim rs As ReconSource = LeftReconSource
            '_left = sql.GetDataTable(ReconSource.GetSelect(rs)) ', _reconciliation.FromDate, _reconciliation.ToDate)
            'LeftSet = _left.AsDataView
            rs = _rightReconSource
            '_right = sql.GetDataTable(ReconSource.GetSelect(rs)) ', _reconciliation.FromDate, _reconciliation.ToDate)
            'RightSet = _right.AsDataView
        End Using
    End Sub

    Private Sub LoadReconSource(reconSource As ReconSource)
        Select Case reconSource.ReconDataSource.DataSourceName
            Case "Excel"
                Using excel As New GetExcel
                    reconSource.IsLoaded = excel.Load(reconSource)
                End Using
            Case "SQL"
                Using sql As New GetSQL
                    reconSource.IsLoaded = sql.Load(reconSource, Reconciliation.FromDate, Reconciliation.ToDate)
                End Using
            Case "QuickBooks"
                If reconSource.IsLoaded = True And reconSource.ReconDataSource.IsSlowLoading Then
                    If reconSource.ReconDataSource.IsSlowLoading Then
                    End If
                End If
                Using qbd As New GetQBD
                    reconSource.IsLoaded = qbd.LoadReport(reconSource, Reconciliation.FromDate, Reconciliation.ToDate)
                End Using
        End Select
    End Sub

#End Region

End Class

Public Class MockMainViewModel

    Public ReadOnly Property Solution As Solution 'active solution
        Get
            'Dim _solution As Solution
            'Using m As New Mock
            '    DataSources = m.MockLoadDataSources
            '    _solution = Task.Run(Function() m.MockLoadSolutionAsync(1)).GetAwaiter().GetResult() 'Model for Solution
            '    Reconciliation = _solution.Reconciliations(0)
            '    LeftReconSource = _solution.Reconciliations(0).LeftReconSource
            'End Using
            'Solution = _solution
        End Get
    End Property
    Public Property Reconciliation As Reconciliation
    Public Property LeftReconSource As ReconSource
    Public Shared ReadOnly Property DataSources As List(Of DataSource)
        Get
            DataSources.Add(New DataSource With {.DataSourceName = "Excel"})
            DataSources.Add(New DataSource With {.DataSourceName = "Intact"})
            DataSources.Add(New DataSource With {.DataSourceName = "QuickBooks"})
            DataSources.Add(New DataSource With {.DataSourceName = "ServiceTitan"})
            DataSources.Add(New DataSource With {.DataSourceName = "SQL"})
        End Get
    End Property

End Class