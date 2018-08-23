Imports System.Data
Imports System.Data.SqlClient
Imports System.Text
Imports Rekonator

Public Class GetSQL
    Implements IDisposable

    Dim dtResultToTranspose As DataTable = Nothing
    Public Function Load(reconSource As ReconSource, fromDate As DateTime, toDate As DateTime) As Boolean
        Try
            If Not reconSource.Parameters.IsExist("Create") Then
                Application.ErrorMessage($"SQL Datasource requires a Create parameter for table {reconSource.ReconTable}")
                Return False
            End If

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

                        Dim sqlParams As SqlParameter() = MakeParameters(reconSource, fromDate, toDate)
                        If sqlParams IsNot Nothing Then
                            selectCommand.Parameters.AddRange(sqlParams)
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

                    If reconSource.Parameters.IsExist("Transpose") Then
                        Dim paramValue As String = reconSource.Parameters.GetParameter("Transpose")
                        If paramValue.ToLower.StartsWith("y") Or paramValue.ToLower.StartsWith("t") Then
                            commandText = ReconSource.GetLoaded(reconSource)
                            dtResultToTranspose = rekonConnection.GetDataTable(commandText)
                            If dtResultToTranspose.Rows.Count <> 1 Then
                                Application.ErrorMessage($"Unable to Transpose one row of {reconSource.ReconTable} because it has {dtResultToTranspose.Rows.Count} rows.")
                            End If
                        End If
                    End If
                End Using
            End Using

            If dtResultToTranspose IsNot Nothing Then Return TransposeResult(reconSource)
            Return True
        Catch ex As Exception
            Application.ErrorMessage($"Error loading table {reconSource.ReconTable}: {ex.Message}")
        End Try
        Return False
    End Function

    Private Function MakeParameters(reconSource As ReconSource, fromDate As DateTime, toDate As DateTime) As SqlParameter()
        Dim params As New List(Of SqlParameter)
        If reconSource.Parameters.IsExist("Arguments") Then
            Dim args As String() = reconSource.Parameters.GetParameter("Arguments").Split(";")
            For Each arg As String In args
                Dim parts As String() = arg.Split("=")
                If parts.Count = 3 Then
                    Dim sqlDbType As SqlDbType = Nothing
                    Select Case parts(1).ToString.ToLower
                        Case "nvarchar"
                            sqlDbType = Data.SqlDbType.NVarChar
                        Case "datetime"
                            sqlDbType = Data.SqlDbType.DateTime
                    End Select
                    If parts(2).ToString.ToLower = "{from}" Then parts(2) = fromDate
                    If parts(2).ToString.ToLower = "{to}" Then parts(2) = toDate
                    If parts(2).ToString.ToLower = "null" Then parts(2) = String.Empty

                    Dim param As New SqlParameter With {.ParameterName = parts(0), .SqlDbType = sqlDbType, .Value = parts(2)}
                    params.Add(param)
                End If
            Next
        End If
        Return params.ToArray
    End Function

    Private Function TransposeResult(reconSource As ReconSource) As Boolean
        Using rekonConnection As New SQL()
            rekonConnection.DropTable(reconSource.ReconTable)
            'Allow end using and connection to close to drop table
        End Using

        Dim sb As New StringBuilder
        sb.Append($"CREATE TABLE {reconSource.ReconTable} ( ")
        For Each column As Column In reconSource.Columns
            sb.Append("[")
            sb.Append($"{reconSource.ColumnPrefix}{column.ColumnName}")
            sb.Append("] ")
            sb.Append(column.ColumnType)
            sb.Append(", ")
        Next
        sb.AppendLine("rekonid int IDENTITY(1,1) )")
        sb.AppendLine($"INSERT INTO {reconSource.ReconTable}")
        sb.AppendLine("VALUES")

        For Each resultColumn As DataColumn In dtResultToTranspose.Columns
            If resultColumn.ColumnName = "rekonid" Then Continue For
            sb.AppendLine($"( N'{resultColumn.ColumnName}', {dtResultToTranspose.Rows(0)(resultColumn)} ),")
        Next
        sb.Length = sb.Length - 3
        sb.AppendLine()
        Using rekonConnection As New SQL()
            rekonConnection.ExecuteNonQuery(sb.ToString)
        End Using
        Return True
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
