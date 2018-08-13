Imports System.Data.SqlClient

Public Class GetSQL
    Implements IDisposable

    Public Function Load(reconSource As ReconSource, fromDate As DateTime, toDate As DateTime)
        Using sourceConnection = New SqlConnection(reconSource.Parameters.GetParameter("connectionstring"))
            Using rekonConnection As New SQL()
                rekonConnection.DropTable(reconSource.ReconTable)
            End Using
            Using rekonConnection As New SQL()
                Dim commandText As String = String.Empty

                sourceConnection.Open()

                commandText = reconSource.Parameters.GetParameter("Create")
                commandText = commandText.Insert(commandText.Length - 1, ", rekonid int IDENTITY(1,1)")
                rekonConnection.ExecuteNonQuery(commandText)

                If reconSource.Parameters.IsExist("commandtext") Then
                    commandText = reconSource.Parameters.GetParameter("commandtext")
                Else
                    commandText = rekonConnection.GetFromCommandPath(reconSource.Parameters.GetParameter("commandpath"))
                End If
                Using selectCommand As New SqlCommand(commandText, sourceConnection)

                    If Not fromDate.Equals(DateTime.MinValue) Then
                        selectCommand.Parameters.Add(New SqlParameter With {.ParameterName = "@From", .SqlDbType = Data.SqlDbType.DateTime, .Value = fromDate})
                    End If
                    If Not toDate.Equals(DateTime.MinValue) Then
                        selectCommand.Parameters.Add(New SqlParameter With {.ParameterName = "@To", .SqlDbType = Data.SqlDbType.DateTime, .Value = toDate})
                    End If
                    Application.Message($"Loading SQL Datasource {selectCommand.CommandText}")
                    Using selectReader As SqlDataReader = selectCommand.ExecuteReader()
                        Using destBulkInsert = New SqlBulkCopy(rekonConnection.Connection)
                            destBulkInsert.BulkCopyTimeout = 300
                            destBulkInsert.DestinationTableName = reconSource.ReconTable
                            destBulkInsert.WriteToServer(selectReader)
                        End Using
                    End Using

                End Using
            End Using
        End Using
    End Function


    '    var mergeQuery = "INSERT INTO table2(id, name, adresse) SELECT * FROM #t WHERE #t.id NOT IN(SELECT id FROM table2)";
    '    Using (var mergeCommand = New SqlCommand(mergeQuery, destinationConnection))
    '    {
    '        mergeCommand.ExecuteNonQuery();
    '    }
    '}

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
