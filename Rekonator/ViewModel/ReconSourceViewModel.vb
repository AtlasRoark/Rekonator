Public Class ReconSourceViewModel
    Inherits ViewModelBase

    Public Property DataSources As List(Of DataSource)
        Get
            DataSources = DataSource.DataSources
        End Get
        Set(value As List(Of DataSource))
            DataSource.DataSources = value
            'OnPropertyChanged("DataSources")
        End Set
    End Property

    'Public Property SelectedDataSource As DataSource

    Public Property MainViewModel As MainViewModel
    Public Property ReconSource As ReconSource  ' active reconciliation from _solution
        Get
            ReconSource = _reconSource
        End Get
        Set(value As ReconSource)
            _reconSource = value
            OnPropertyChanged("ReconSource")
        End Set
    End Property
    Private _reconSource As New ReconSource

End Class

