Imports System.IO
Imports System.Runtime.Serialization.Formatters.Binary
Imports System.Runtime.Serialization

<Serializable()>
Public Class Solution
    Implements ISerializable
    ' Empty constructor required to compile.
    Public Sub New()
    End Sub

    Private _solutionName As String = "(New Solution)"

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
        Dim fs As New FileStream(fileName, FileMode.Create)
        Dim formatter As IFormatter = New BinaryFormatter()
        formatter.Serialize(fs, solution)
        fs.Close()
    End Sub

    Public Shared Function MakeNewReconcilition(Optional solution As Solution = Nothing) As Solution
        If solution Is Nothing Then solution = New Solution
        If solution.Reconciliations Is Nothing Then
            solution.Reconciliations = New List(Of Reconciliation)
            Dim r As New Reconciliation With {
                .CompletenessComparisions = New List(Of Comparision),
                .MatchingComparisions = New List(Of Comparision)
            }
            solution.Reconciliations.Add(r)
            Dim recon As New ReconSource With {
                .Aggregations = New List(Of Aggregate),
                .Columns = New List(Of Column),
                .Parameters = New List(Of Parameter)
            }
            solution.Reconciliations.Last.LeftReconSource = recon

            recon = New ReconSource With {
                .Aggregations = New List(Of Aggregate),
                .Columns = New List(Of Column),
                .Parameters = New List(Of Parameter)
                }
            solution.Reconciliations.Last.RightReconSource = recon
            Return solution
        End If
    End Function

    Public Shared Function LoadSolution(fileName As String) As Solution
        Try
            Dim fs As New FileStream(fileName, FileMode.Open)
            Dim formatter As IFormatter = New BinaryFormatter()
            Dim loadedSolution As Solution = TryCast(formatter.Deserialize(fs), Solution)
            If loadedSolution Is Nothing Then Throw New Exception("The loaded solution was empty")
            Return loadedSolution
        Catch ex As Exception
            Application.ErrorMessage($"Error opening solution file {fileName}: {ex.Message}")
        End Try
        Return Nothing
    End Function
End Class

