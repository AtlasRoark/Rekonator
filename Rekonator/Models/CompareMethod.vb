<Serializable()>
Public Class CompareMethod
    Delegate Function CompareMethodDelegate(ByVal leftValue As String, ByVal rightValue As String, options As Integer) As Boolean
    Property Name As String
    Property Method As CompareMethodDelegate

    Public Shared CompareMethods As New List(Of CompareMethod)
    Public Shared Sub Add(name As String, testName As CompareMethodDelegate)
        CompareMethods.Add(New CompareMethod With {.Name = name, .Method = testName})
    End Sub
    Public Shared Function GetMethod(operationName As String) As CompareMethod
        Return CompareMethods.Where(Function(w) w.Name = operationName).FirstOrDefault
    End Function

End Class
