Imports System.Runtime.CompilerServices

Module ExtensionMethods

    <Extension()>
    Public Function GetParameter(ByVal parameters As List(Of Parameter), parameterName As String) As String
        If parameters.IsExist(parameterName) Then
            Return parameters.Where(Function(w) w.ParameterName = parameterName).First.ParameterValue
        Else
            Return String.Empty
        End If
    End Function

    <Extension()>
    Public Function IsExist(ByVal parameters As List(Of Parameter), lookup As String) As Boolean
        Return parameters.Exists(Function(e) e.ParameterName = lookup)
    End Function


    <Extension()>
    Public Sub AddParameter(ByVal parameters As List(Of Parameter), parameterName As String, parameterValue As String)
        parameters.Add(New Parameter With {.ParameterName = parameterName, .ParameterValue = parameterValue})
    End Sub

    <Extension()>
    Public Sub UpdateParameter(ByVal parameters As List(Of Parameter), parameterName As String, parameterValue As String)
        If parameters.IsExist(parameterName) Then
            parameters.Where(Function(w) w.ParameterName = parameterName).First.ParameterValue = parameterValue
        Else
            parameters.AddParameter(parameterName, parameterValue)
        End If
    End Sub

    <Extension()>
    Public Sub AddColumn(ByVal columns As List(Of Column), columnName As String, columnType As String)
        columns.Add(New Column With {.ColumnName = columnName, .ColumnType = columnType})
    End Sub

    <Extension()>
    Public Sub AddColumns(ByVal columns As List(Of Column), nameTypePairs As String())
        For idx As Integer = 0 To nameTypePairs.Count Step 2
            columns.Add(New Column With {.ColumnName = nameTypePairs(idx), .ColumnType = nameTypePairs(idx + 1)})
        Next
    End Sub

End Module
