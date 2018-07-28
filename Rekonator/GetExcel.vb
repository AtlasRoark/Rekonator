Imports ExcelDataReader
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Text

Public Class GetExcel
    Implements IDisposable

    'Private Shared Sub HasRows(ByVal connection As SqlConnection)
    '    Using connection
    '        Dim command As SqlCommand = New SqlCommand("SELECT CategoryID, CategoryName FROM Categories;", connection)
    '        connection.Open()
    '        Dim reader As SqlDataReader = command.ExecuteReader()

    '        If reader.HasRows Then

    '            While reader.Read()
    '                Console.WriteLine("{0}" & vbTab & "{1}", reader.GetInt32(0), reader.GetString(1))
    '            End While
    '        Else
    '            Console.WriteLine("No rows found.")
    '        End If

    '        reader.Close()
    '    End Using
    'End Sub

    Public Function MakeReconSource(sourcePath As String, worksheetName As String, reconSourceName As String) As ReconSource
        Application.Message($"Loading Table {reconSourceName} from Excel Worksheet {worksheetName}")
        Dim iRow As Integer = 0
        Try
            Using fileStream = File.Open(sourcePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite) 'FileShare is ReadWrite even though FileAccess is Read Only.  This allows file to be open if it open in another process e.g. Excel
                Using excelReader = ExcelReaderFactory.CreateReader(fileStream)
                    Do While excelReader.Name().ToLower <> worksheetName.ToLower
                        excelReader.NextResult()
                    Loop

                    If excelReader.Name().ToLower <> worksheetName.ToLower Then
                        Return Nothing
                    End If

                    'Get Headers
                    If Not excelReader.Read() Then Return Nothing
                    Dim headerList As New List(Of String)
                    Dim fieldCount = excelReader.FieldCount
                    For idx = 0 To fieldCount - 1
                        If excelReader.GetValue(idx) Is Nothing Then Exit For
                        headerList.Add(excelReader.GetValue(idx))
                    Next
                    fieldCount = headerList.Count

                    'Get Field Types and First Row
                    If Not excelReader.Read() Then Return Nothing
                    Dim typeList As New List(Of Type)
                    Dim rowList As New List(Of Object)
                    For idx = 0 To fieldCount - 1
                        typeList.Add(excelReader.GetFieldType(idx))
                        If typeList(idx) Is Nothing Then typeList(idx) = GetType(String)
                        rowList.Add(excelReader.GetValue(idx))
                    Next

                    'Make Table
                    Dim sb As New StringBuilder
                    sb.AppendLine($"IF(OBJECT_ID('Rekonator..{reconSourceName}') IS NOT NULL) DROP TABLE [{reconSourceName}];")

                    sb.AppendLine($"CREATE TABLE [{reconSourceName}] (")
                    For idx = 0 To fieldCount - 1
                        sb.Append($"[{headerList(idx)}]")
                        Select Case typeList(idx).Name
                            Case "Double"
                                sb.AppendLine(" DECIMAL(14,2) NULL,")
                            Case "DateTime"
                                sb.AppendLine(" DATETIME NULL,")
                            Case "String"
                                sb.AppendLine(" NVARCHAR(4000) NULL,")
                            Case Else
                                Beep()
                        End Select

                    Next
                    sb.AppendLine(");")

                    Using connection As New SqlConnection(Application.ConnectionString)
                        Dim command As New SqlCommand(sb.ToString, connection)
                        command.Connection.Open()
                        command.ExecuteNonQuery()
                    End Using

                    Do
                        iRow += 1
                        If iRow Mod 1000 = 0 Then Application.Message(iRow.ToString)
                        'Starts with First Row already read from excel
                        sb.Clear()
                        sb.AppendLine($"INSERT INTO [dbo].[{reconSourceName}] VALUES (")

                        'Whats Faster? or string.join
                        For idx = 0 To fieldCount - 1
                            If rowList(idx) Is Nothing Then
                                sb.AppendLine("Null,")
                            Else
                                Select Case typeList(idx).Name
                                    Case "Double"
                                        sb.AppendLine($"{rowList(idx)},")
                                    Case "DateTime"
                                        sb.AppendLine($"'{CDate(rowList(idx)).ToString("yyyy-MM-dd hh:mm:ss")}',")
                                    Case "String"
                                        sb.AppendLine($"'{rowList(idx).ToString.Replace("'", "''")}',")
                                End Select
                            End If
                        Next
                        sb.Replace(",", "", sb.Length - 3, 1)
                        sb.AppendLine(");")
                        Using connection As New SqlConnection(Application.ConnectionString)
                            Dim command As New SqlCommand(sb.ToString, connection)
                            command.Connection.Open()
                            command.ExecuteNonQuery()
                        End Using

                        'Get Next Row
                        If Not excelReader.Read() Then Exit Do
                        rowList.Clear()
                        For idx = 0 To fieldCount - 1
                            rowList.Add(excelReader.GetValue(idx))
                        Next
                    Loop

                End Using
            End Using
            Application.Message($"Completed: {iRow}")
            ' Return
        Catch ex As Exception
            Application.ErrorMessage($"Row: {iRow}: {ex.Message}")

        End Try
        Return Nothing

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

