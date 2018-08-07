Imports Microsoft.Win32

Public Class SystemDialog
    Inherits Window
    Implements IDisposable

    Public Function OpenFile() As String
        Dim openFileDialog As New OpenFileDialog With {
            .Multiselect = False,
            .Filter = "Rekonator files (*.rek)|*.rek",
            .InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            .Title = "Open Rekonator Solution"
        }

        If openFileDialog.ShowDialog() = True Then
            For Each filename As String In openFileDialog.FileNames
                Return filename
            Next
        End If
        Return String.Empty
    End Function

    Public Function SaveFile() As String
        Dim saveFileDialog As New SaveFileDialog With {
            .Filter = "Rekonator files (*.rek)|*.rek",
            .InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            .AddExtension = True,
            .DefaultExt = "rek",
            .Title = "Save Rekonator Solution"
        }

        If saveFileDialog.ShowDialog() = True Then
            Return saveFileDialog.FileName
        End If
        Return String.Empty
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


