Imports System.Globalization

Public Class ValueComparer

    Public Shared Function CompareIntegerValues(leftValue As String, rightValue As String, options As Integer) As Boolean
        'options = 1 means has to be positive
        'options = 0 means has to be positive/negative
        'options = - 1 means has to be negative
        Dim leftResult As Integer
        Dim rightResult As Integer
        If Integer.TryParse(leftValue, leftResult) AndAlso Integer.TryParse(rightValue, rightResult) Then
            Select Case options
                Case -1
                    Return leftResult = rightResult And leftResult <= 0
                Case 1
                    Return leftResult = rightResult And leftResult >= 0
                Case Else
                    Return leftResult = rightResult
            End Select
        End If
        Return False
    End Function

    Public Shared Function CompareSingleValues(leftValue As String, rightValue As String, digits As Integer) As Boolean
        Dim leftResult As Single
        Dim rightResult As Single
        If Single.TryParse(leftValue, leftResult) AndAlso Single.TryParse(rightValue, rightResult) Then
            If digits > -1 Then
                leftResult = Math.Round(leftResult, digits)
                rightResult = Math.Round(rightResult, digits)
            End If
            Return leftResult = rightResult
        End If
        Return False
    End Function

    Public Shared Function CompareStringValues(leftValue As String, rightValue As String, isCaseSensitive As Integer) As Boolean
        'isCaseSensitive: 0 No, 1 Yes
        If CultureInfo.CurrentCulture.CompareInfo.Compare(leftValue, rightValue, IIf(isCaseSensitive = 1, CompareOptions.None, CompareOptions.IgnoreCase)) = 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Shared Function CompareDateValues(leftValue As Date, rightValue As Date, options As Integer) As Boolean
        '0 Exact
        Return (leftValue = rightValue)
    End Function
End Class
