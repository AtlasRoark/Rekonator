Imports System.IO
Imports System.Runtime.Serialization.Formatters.Binary
Imports System.Runtime.Serialization


<Serializable()>
Public Class Solution
    Implements ISerializable
    ' Empty constructor required to compile.
    Public Sub New()
    End Sub

    Private _solutionName As String

    Public Property SolutionName As String
        Get
            Return _solutionName
        End Get
        Set(value As String)
            _solutionName = value
        End Set
    End Property

    Private _reconciliations As List(Of Reconciliation)

    Public Property Reconciliations As List(Of Reconciliation)
        Get
            Return _reconciliations
        End Get
        Set(value As List(Of Reconciliation))
            _reconciliations = value
        End Set
    End Property

    ' Implement this method to serialize data. The method is called 
    ' on serialization.
    Public Sub getobjectdata(info As SerializationInfo, context As StreamingContext) Implements ISerializable.GetObjectData
        ' use the addvalue method to specify serialized values.
        info.AddValue("SolutionName", _solutionName, GetType(String))
        info.AddValue("Reconciliations", _reconciliations, GetType(List(Of Reconciliation)))

    End Sub

    ' The special constructor is used to deserialize values.
    Public Sub New(info As SerializationInfo, context As StreamingContext)
        ' Reset the property value using the GetValue method.
        _solutionName = DirectCast(info.GetValue("SolutionName", GetType(String)), String)
        _reconciliations = DirectCast(info.GetValue("Reconciliations", GetType(List(Of Reconciliation))), List(Of Reconciliation))
    End Sub

    Public Shared Sub SaveSolution(fileName As String, solution As Solution)
        ' Save an instance of the type and serialize it.
        'Delegates can't be saved/restored.  remove them
        For Each r As Reconciliation In solution.Reconciliations
            For Each c As Comparision In r.CompletenessComparisions.Union(r.MatchingComparisions)
                c.ComparisionMethod.Method = Nothing
            Next
        Next
        Dim fs As New FileStream(fileName, FileMode.Create)
        Dim formatter As IFormatter = New BinaryFormatter()
        formatter.Serialize(fs, solution)
        fs.Close()
    End Sub
    Public Shared Function LoadSolution(fileName As String) As Solution
        Dim fs As New FileStream(fileName, FileMode.Open)
        Dim formatter As IFormatter = New BinaryFormatter()
        Dim loadedSolution As Solution = DirectCast(formatter.Deserialize(fs), Solution)
        For Each r In loadedSolution.Reconciliations
            For Each c As Comparision In r.CompletenessComparisions.Union(r.MatchingComparisions)
                c.ComparisionMethod = CompareMethod.GetMethod(c.ComparisionMethod.Name)
            Next
        Next
    End Function
End Class

