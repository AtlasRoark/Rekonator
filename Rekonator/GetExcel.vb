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

    Public Function Load(reconSource As ReconSource) As Boolean
        Dim reconTable As String = reconSource.ReconTable
        Dim filePath As String = reconSource.Parameters("FilePath")
        Dim worksheetName As String = reconSource.Parameters("Worksheet")
        Application.Message($"Loading Table {reconTable} from Excel Worksheet {worksheetName}")
        Try
            Using fileStream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite) 'FileShare is ReadWrite even though FileAccess is Read Only.  This allows file to be open if it open in another process e.g. Excel
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

                    Using sql As New SQL(reconTable, fieldCount, headerList, typeList)
                        If Not sql.CreateTable() Then
                            Return False
                        End If

                        Do
                            If Not sql.InsertRow(rowList) Then
                                Return False
                            End If

                            'Get Next Row
                            If Not excelReader.Read() Then Exit Do
                            rowList.Clear()
                            For idx = 0 To fieldCount - 1
                                rowList.Add(excelReader.GetValue(idx))
                            Next
                        Loop
                    End Using

                End Using
            End Using
            Application.Message("Completed.")
            Return True
        Catch ex As Exception
            Application.ErrorMessage($"Error Loading Excel: {ex.Message}")

        End Try
        Return False

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

