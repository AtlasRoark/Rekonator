Imports System.Data
Imports System.Data.SqlClient
Imports System.Text

Public Class SQL
    Implements IDisposable

    Private _connection As SqlConnection
    Private _reconTable As String
    Private _fieldCount As Integer
    Private _headerList As List(Of String)
    Private _typeList As List(Of Type)
    Private _sb As New StringBuilder
    Private _rowNumber As Integer = 0
    Private _CurrencyKeyWords() As String = {"Amount", "Total", "Payment", "Subtotal", "Debit", "Credit", "Balance", "Amt", "Cost"}

    Public Sub New()
        OpenConnection()
    End Sub



    Public Sub New(reconTable As String, fieldCount As Integer, headerList As List(Of String), typeList As List(Of Type))
        Me._reconTable = reconTable
        Me._fieldCount = fieldCount
        Me._headerList = headerList
        Me._typeList = typeList
        OpenConnection()
    End Sub

    Public Function GetDataTable(selectCommand) As DataTable
        Try
            Application.Message($"Loading {selectCommand} from Rekonator")

            Dim adapter As New SqlDataAdapter(selectCommand, _connection)
            Dim dt As New DataTable()
            adapter.Fill(dt)
            Return dt

        Catch ex As Exception
            Application.ErrorMessage($"Error getting dataview for {selectCommand}: {ex.Message}")
        End Try
    End Function

    Public Function CreateTable() As Boolean
        Try
            _sb.AppendLine($"IF(OBJECT_ID('Rekonator..{_reconTable}') IS NOT NULL) DROP TABLE [{_reconTable}];")

            _sb.AppendLine($"CREATE TABLE [{_reconTable}] (")
            _sb.AppendLine("id int IDENTITY(1,1),")
            For idx = 0 To _fieldCount - 1
                _sb.Append($"[{_headerList(idx)}]")
                Select Case _typeList(idx).Name
                    Case "Double"
                        If _CurrencyKeyWords.Contains(_headerList(idx)) Then
                            _sb.AppendLine(" DECIMAL(14,2) NULL,")
                        Else
                            _sb.AppendLine(" DECIMAL(16,6) NULL,")

                        End If
                    Case "DateTime"
                        _sb.AppendLine(" DATETIME NULL,")
                    Case "String"
                        _sb.AppendLine(" NVARCHAR(4000) NULL,")
                    Case Else
                        Application.ErrorMessage("Unknown Type")
                End Select

            Next
            _sb.AppendLine(");")

            Dim command As New SqlCommand(_sb.ToString, _connection)
            command.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Application.ErrorMessage($"Error Creating Table {_reconTable}: {ex.Message}")
        End Try
        Return False
    End Function

    Public Function InsertRow(rowList As List(Of Object)) As Boolean
        Try
            _rowNumber += 1
            If _rowNumber Mod 1000 = 0 Then Application.Message(_rowNumber.ToString)
            'Starts with First Row already read from excel
            _sb.Clear()
            _sb.AppendLine($"INSERT INTO [dbo].[{_reconTable}] VALUES (")

            'Whats Faster? or string.join
            For idx = 0 To _fieldCount - 1
                If rowList(idx) Is Nothing Then
                    _sb.AppendLine("Null,")
                Else
                    Select Case _typeList(idx).Name
                        Case "Double"
                            _sb.AppendLine($"{rowList(idx)},")
                        Case "DateTime"
                            _sb.AppendLine($"'{CDate(rowList(idx)).ToString("yyyy-MM-dd hh:mm:ss")}',")
                        Case "String"
                            _sb.AppendLine($"'{rowList(idx).ToString.Replace("'", "''")}',")
                    End Select
                End If
            Next
            _sb.Replace(",", "", _sb.Length - 3, 1)
            _sb.AppendLine(");")
            Dim command As New SqlCommand(_sb.ToString, _connection)
            command.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Application.ErrorMessage($"Error Inserting Row {_reconTable} Row {_rowNumber}: {ex.Message}")
        End Try
        Return False
    End Function

    Private Sub OpenConnection()
        _connection = New SqlConnection(Application.ConnectionString)
        _connection.Open()
    End Sub

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                _connection.Close()
                _connection = Nothing
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
