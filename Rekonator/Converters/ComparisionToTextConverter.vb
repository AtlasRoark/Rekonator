
Imports System.Globalization
Imports System.Text
Imports Rekonator.Comparision

Public Class ComparisionToTextConverter
    Implements IValueConverter

    Public Function Convert(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.Convert
        Dim comparisons As List(Of Comparision) = TryCast(value, List(Of Comparision))
        If comparisons Is Nothing Then
            Return ""
        Else
            Dim sb As New StringBuilder
            For Each comp As Comparision In comparisons
                sb.Append($"{comp.LeftColumn}=")
                Select Case comp.ComparisionTest
                    Case ComparisionType.TextEquals
                        sb.Append($"~={comp.RightColumn}")
                    Case ComparisionType.TextCaseEquals
                        sb.Append($"-={comp.RightColumn}")
                    Case ComparisionType.NumberEquals
                        sb.Append($"#{comp.Percision}={comp.RightColumn}")
                    Case ComparisionType.DateEquals
                        sb.Append($"/={comp.RightColumn}")
                    Case ComparisionType.DateTimeEquals
                        sb.Append($":={comp.RightColumn}")
                End Select
                sb.Append(";")
            Next
            If sb.Length > 0 Then sb.Length -= 1
            Return sb.ToString
        End If
    End Function

    Public Function ConvertBack(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.ConvertBack
        Dim compText As String = TryCast(value, String)
        If String.IsNullOrEmpty(compText) Then
            Return Nothing
        Else
            Dim comparisions = New List(Of Comparision)

            Dim compStrs As String() = compText.Split(";")
            For Each compStr In compStrs
                Dim parts As String() = compStr.Split("=")

                Dim comp As New Comparision
                With comp
                    .LeftColumn = parts(0)
                    .RightColumn = parts(2)
                    Select Case parts(1).Substring(0, 1)
                        Case "-"
                            .ComparisionTest = ComparisionType.TextCaseEquals
                        Case "~"
                            .ComparisionTest = ComparisionType.TextEquals
                        Case "#"
                            .ComparisionTest = ComparisionType.NumberEquals
                            .Percision = CInt(parts(1).Substring(1))
                        Case "/"
                            .ComparisionTest = ComparisionType.DateEquals
                        Case ":"
                            .ComparisionTest = ComparisionType.DateTimeEquals
                    End Select
                End With
                comparisions.Add(comp)
            Next
            Return comparisions
        End If
    End Function


End Class