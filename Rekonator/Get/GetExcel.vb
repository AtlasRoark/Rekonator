Imports ExcelDataReader
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Text

Public Class GetExcel
    Implements IDisposable

    Public Function Load(reconSource As ReconSource) As Boolean
        Dim reconTable As String = reconSource.ReconTable
        Dim filePath As String = reconSource.Parameters.GetParameter("FilePath")
        If String.IsNullOrWhiteSpace(filePath) Then filePath = reconSource.Parameters.GetParameter("Workbook")
        Dim worksheetName As String = reconSource.Parameters.GetParameter("Worksheet")
        Application.Message($"Loading Table {reconTable} from Excel Worksheet {worksheetName}")
        Dim testField As Integer = 2
        Dim fillDown As Integer() = {0, 1}
        Dim rowCount As Integer = 0
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
                    Dim fieldCount As Integer = 0
                    Dim fieldList As New List(Of String)
                    Dim typeList As New List(Of String)

                    If reconSource.Columns Is Nothing Then
                        fieldCount = excelReader.FieldCount
                        For idx = 0 To fieldCount - 1
                            If excelReader.GetValue(idx) Is Nothing Then Exit For
                            fieldList.Add(excelReader.GetValue(idx))
                        Next
                    Else
                        fieldList = reconSource.Columns.Select(Function(s) s.ColumnName).ToList
                        typeList = reconSource.Columns.Select(Function(s) s.ColumnType).ToList
                    End If
                    fieldCount = fieldList.Count

                    'Get Field Types and First Row
                    If Not excelReader.Read() Then Return Nothing
                    Dim rowList As New List(Of Object)
                    Dim lastRowList As New List(Of Object)

                    For idx = 0 To fieldCount - 1
                        If typeList.Count = 0 Then
                            typeList.Add(excelReader.GetFieldType(idx).Name)
                            If typeList(idx) Is Nothing Then typeList(idx) = "String"
                        End If
                        rowList.Add(excelReader.GetValue(idx))
                    Next

                    Using sql As New SQL(reconTable, fieldCount, fieldList, typeList)
                        If sql.CreateTable() Then
                            Application.Message($"Table: '{reconTable}' created.")
                        Else
                            Return False
                        End If

                        Do
                            If Not String.IsNullOrWhiteSpace(rowList(testField)) Then

                                If sql.InsertRow(rowList) Then
                                    rowCount += 1

                                Else
                                    Return False
                                End If
                            End If

                            'Get Next Row
                            If Not excelReader.Read() Then Exit Do
                            lastRowList.Clear()
                            lastRowList.AddRange(rowList)
                            rowList.Clear()
                            For idx = 0 To fieldCount - 1
                                Dim value As Object = excelReader.GetValue(idx)
                                If String.IsNullOrWhiteSpace(value) Then
                                    If fillDown.Contains(idx) Then
                                        value = lastRowList(idx)
                                    End If
                                End If
                                rowList.Add(value)
                            Next
                        Loop
                    End Using

                End Using
            End Using
            Application.Message($"Excel Load Completed. {rowCount} rows.")
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

