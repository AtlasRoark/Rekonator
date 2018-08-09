Imports System.Globalization
Imports System.Text
Imports Rekonator.AggregateOperation

Public Class AggToTextConverter
    Implements IValueConverter

    Public Function Convert(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.Convert
        Dim aggregates As List(Of Aggregate) = TryCast(value, List(Of Aggregate))
        If aggregates Is Nothing Then
            Return ""
        Else
            Dim sb As New StringBuilder
            For Each agg As Aggregate In aggregates
                sb.Append(String.Join(",", agg.GroupByColumns) + ":")
                For Each aggop As AggregateOperation In agg.AggregateOperations
                    sb.Append(aggop.AggregateColumn + "=")
                    sb.Append(aggop.Operation.ToString + "(")
                    sb.Append(aggop.SourceColumn + "),")
                Next
                sb.Length -= 1
                sb.Append(";")
            Next
            If sb.Length > 0 Then sb.Length -= 1
            Return sb.ToString
        End If
    End Function

    Public Function ConvertBack(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.ConvertBack
        Dim aggText As String = TryCast(value, String)
        If aggText Is Nothing Then
            Return Nothing
        Else
            Dim aggregates = New List(Of Aggregate)

            Dim aggStrs As String() = aggText.Split(";")
            For Each aggStr In aggStrs
                Dim halves As String() = aggStr.Split(":")
                If halves.Count <> 2 Then Return Nothing
                Dim agg As New Aggregate

                agg.GroupByColumns = halves(0).Split(",")


                Dim aggOps As New List(Of AggregateOperation)

                Dim aggOpStrs As String() = halves(1).Split(",")
                For Each aggOpStr In aggOpStrs
                    Dim ao As New AggregateOperation
                    Dim parts As String() = aggOpStr.Split("=")
                    ao.AggregateColumn = parts(0)
                    Dim funcPartStrs As String() = parts(1).Split("(")
                    Select Case funcPartStrs(0)
                        Case "Avg"
                            ao.Operation = AggregateFunction.Avg
                        Case "Count"
                            ao.Operation = AggregateFunction.Count
                        Case "First"
                            ao.Operation = AggregateFunction.First
                        Case "Last"
                            ao.Operation = AggregateFunction.Last
                        Case "Max"
                            ao.Operation = AggregateFunction.Max
                        Case "Min"
                            ao.Operation = AggregateFunction.Min
                        Case "Sum"
                            ao.Operation = AggregateFunction.Sum
                    End Select

                    ao.SourceColumn = funcPartStrs(1).Replace(")", "")
                    aggOps.Add(ao)

                Next
                agg.AggregateOperations = aggOps
                aggregates.Add(agg)
            Next
            Return aggregates
        End If

    End Function


End Class