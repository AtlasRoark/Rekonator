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
            sb.Length -= 1
            Return sb.ToString
        End If
    End Function

    Public Function ConvertBack(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.ConvertBack
        Dim aggText As String = TryCast(value, String)
        If aggText Is Nothing Then
            Return ""
        Else
            Dim aggOps As New List(Of AggregateOperation)
            Dim aggregates As New List(Of Aggregate)
            aggOps.Add(New AggregateOperation With {.SourceColumn = "Amount", .AggregateColumn = "Total", .Operation = AggregateOperation.AggregateFunction.Sum})
            aggregates.Add(New Aggregate With {.GroupByColumns = {"TXN ID"}, .AggregateOperations = aggOps})

            Return aggregates
        End If
    End Function


End Class