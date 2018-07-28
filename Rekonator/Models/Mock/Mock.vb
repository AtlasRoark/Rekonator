Imports Rekonator

Public Class SettingsModel
    Property Datasources As List(Of DataSource)
    'Property CompareMethods As List(Of CompareMethod)
End Class
Public Class Mock
    Implements IDisposable

    Public Function MockLoadSettings() As SettingsModel
        Dim sm As New SettingsModel
        'CompareMethod.Add("Integer Equals", AddressOf ValueComparer.CompareIntegerValues)
        'CompareMethod.Add("Single Equals", AddressOf ValueComparer.CompareSingleValues)
        'CompareMethod.Add("String Equals", AddressOf ValueComparer.CompareStringValues)
        'CompareMethod.Add("Date Equals", AddressOf ValueComparer.CompareDateValues)
        'sm.CompareMethods = CompareMethod.CompareMethods

        Dim ds As New List(Of DataSource)
        ds.Add(New DataSource With {.DataSourceName = "Excel"})
        ds.Add(New DataSource With {.DataSourceName = "Intact"})
        ds.Add(New DataSource With {.DataSourceName = "QuickBooks"})
        ds.Add(New DataSource With {.DataSourceName = "ServiceTitan"})
        ds.Add(New DataSource With {.DataSourceName = "SQL"})
        sm.Datasources = ds

        Return sm
    End Function

    Public Function MockLoadSolutionAsync(mainWindow As MainWindow) As Task(Of Solution)
        Try
        Catch ex As Exception
        End Try
    End Function


#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
        ' TODO: uncomment the following line if Finalize() is overridden above.
        ' GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class
