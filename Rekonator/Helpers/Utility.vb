Public Class Utility

    Public Shared Function FindAncestor(ByVal child As Visual, ByVal typeAncestor As Type) As Visual
        Dim parent As DependencyObject = VisualTreeHelper.GetParent(child)

        While parent IsNot Nothing AndAlso Not typeAncestor.IsInstanceOfType(parent)
            parent = VisualTreeHelper.GetParent(parent)
        End While

        Return (TryCast(parent, Visual))
    End Function

    Public Shared Sub GetLogicalChildCollection(Of T As DependencyObject)(ByVal parent As DependencyObject, ByVal logicalCollection As List(Of T))
        Dim children As IEnumerable = LogicalTreeHelper.GetChildren(parent)

        For Each child As Object In children

            If TypeOf child Is DependencyObject Then
                Dim depChild As DependencyObject = TryCast(child, DependencyObject)

                If TypeOf child Is T Then
                    logicalCollection.Add(TryCast(child, T))
                End If

                GetLogicalChildCollection(depChild, logicalCollection)
            End If
        Next
    End Sub
End Class
