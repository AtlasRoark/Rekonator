Imports System.Data
Imports System.Data.SqlClient
Imports System.Text

Public Class SQL
    Implements IDisposable

    Private _connection As SqlConnection
    Private _reconTable As String
    Private _fieldCount As Integer
    Private _headerList As List(Of String)
    Private _typeList As List(Of String)
    Private _sb As New StringBuilder
    Private _rowNumber As Integer = 0
    Private _CurrencyKeywords() As String = {"amount", "total", "payment", "subtotal", "debit", "credit", "balance", "amt", "price", "cost", "$"}
    Private _IntegerKeywords() As String = {"id", "#", "status", "No", "Num"}

    Public ReadOnly Property Connection As SqlConnection
        Get
            Connection = _connection
        End Get
    End Property

    Public Sub New()
        OpenConnection()
    End Sub



    Public Sub New(reconTable As String, fieldCount As Integer, headerList As List(Of String), typeList As List(Of String))
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
            Debug.Print(ex.Message)
            Debug.Print(selectCommand)
            Application.ErrorMessage($"Error getting dataview for {selectCommand}: {ex.Message}")
        End Try
        Return Nothing
    End Function

    Public Function CreateTable() As Boolean
        Try
            _sb.Clear()
            _sb.AppendLine($"IF(OBJECT_ID('Rekonator..{_reconTable}') IS NOT NULL) DROP TABLE [{_reconTable}];")

            _sb.AppendLine($"CREATE TABLE [{_reconTable}] (")
            _sb.AppendLine("rekonid int IDENTITY(1,1),")
            For idx = 0 To _fieldCount - 1
                _sb.Append($"[{_headerList(idx)}]")
                Select Case _typeList(idx)
                    Case "Int32", "Integer"
                        _sb.AppendLine(" INT NULL,")
                    Case "Double"
                        If IsCurrency(_headerList(idx)) Then
                            _sb.AppendLine(" DECIMAL(14,2) NULL,")
                        ElseIf IsInteger(_headerList(idx)) Then
                            _sb.AppendLine(" INT NULL,")
                        Else
                            _sb.AppendLine(" DECIMAL(16,6) NULL,")
                        End If
                    Case "Currency"
                        _sb.AppendLine(" DECIMAL(14,2) NULL,")
                    Case "Date"
                        _sb.AppendLine(" DATE NULL,")
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
    Public Function DropTable(Optional tableName As String = "") As Boolean
        If String.IsNullOrEmpty(tableName) Then tableName = _reconTable
        ExecuteNonQuery($"IF(OBJECT_ID('Rekonator..[{tableName}]') IS NOT NULL) DROP TABLE [{tableName}];")
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
                    Select Case _typeList(idx)
                        Case "Double", "Int32", "Integer", "Currency"
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

    Public Function ExecuteNonQuery(execCommand As String) As Integer
        Try
            Dim command As New SqlCommand(execCommand, _connection)
            Return command.ExecuteNonQuery()
        Catch ex As Exception
            Application.ErrorMessage($"Error Executing Non Query {execCommand}: {ex.Message}")
        End Try
    End Function

    Friend Function GetFromCommandPath(commandPath As String) As String
        Return My.Computer.FileSystem.ReadAllText(commandPath, Encoding.UTF8)
    End Function

    Private Sub OpenConnection()
        _connection = New SqlConnection(Application.ConnectionString)
        _connection.Open()
    End Sub

    Private Function IsCurrency(header As String) As Boolean
        Dim parts = header.Split(" ")
        For Each part In parts
            If _CurrencyKeywords.Contains(part.ToLower) Then
                Return True
                Exit For
            End If
        Next
        Return False
    End Function

    Private Function IsInteger(header As String) As Boolean
        Dim parts = header.Split(" ")
        For Each part In parts
            If _IntegerKeywords.Contains(part.ToLower) Then
                Return True
                Exit For
            End If
        Next
        Return False
    End Function
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
