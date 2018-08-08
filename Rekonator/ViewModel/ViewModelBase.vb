Imports System.ComponentModel

Public MustInherit Class ViewModelBase
    Implements INotifyPropertyChanged

    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    Public Sub OnPropertyChanged(propertyName As String)
        Me.CheckPropertyName(propertyName)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propertyName))
    End Sub

    <Conditional("DEBUG")>
    <DebuggerStepThrough>
    Public Sub CheckPropertyName(propertyName As String)
        If TypeDescriptor.GetProperties(Me)(propertyName) Is Nothing Then
            Throw New Exception(String.Format("Count not find property: {0}", propertyName))
        End If
    End Sub

End Class

