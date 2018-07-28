Imports System.Collections.Concurrent
Imports System.ComponentModel
Imports System.Data
Imports System.Threading
Imports Dynamitey.Dynamic

Partial Class ReconWindow
    Implements INotifyPropertyChanged

    Private _aggregates As New List(Of Aggregate)
    Private _completenessComparisions As New List(Of Comparision)
    Private _matchingComparisions As New List(Of Comparision)
    Private _solutionPath As String = String.Empty
    Private _solution As Solution
    'Cant do notify prop change on datatables.  
    Private _left As DataTable
    Private _right As DataTable
    Private _leftDetails As New DataTable
    Private _rightDetails As New DataTable
    Private _differ As New DataTable
    Private _match As New DataTable
    Private _recon As New DataTable

#Region "-- Notify Property Change --"
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    Public Sub OnPropertyChanged(propertyName As String)
        Me.CheckPropertyName(propertyName)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propertyName))
    End Sub

    <Conditional("DEBUG")>
    <DebuggerStepThrough>
    Public Sub CheckPropertyName(propertyName As String)
        If TypeDescriptor.GetProperties(Me)(propertyName) Is Nothing Then
            Throw New Exception(String.Format("Count not find property: {0}", propertyName))
        End If
    End Sub
#End Region

#Region "-- Setting Model Properties --"
    Public Property DataSources As List(Of DataSource)
        Get
            DataSources = _dataSources
        End Get
        Set(value As List(Of DataSource))
            _dataSources = value
            OnPropertyChanged("DataSources")
        End Set
    End Property
    Private _dataSources As List(Of DataSource)
    Public Property Translations As List(Of Translation)
        Get
            Translations = _translations
        End Get
        Set(value As List(Of Translation))
            _translations = value
            OnPropertyChanged("Translations")
        End Set
    End Property
    Private _translations As List(Of Translation)

    Public Property CompareMethods As List(Of CompareMethod)
        Get
            CompareMethods = _compareMethods
        End Get
        Set(value As List(Of CompareMethod))
            _compareMethods = value
        End Set
    End Property
    Private _compareMethods As List(Of CompareMethod)

#End Region

#Region "-- Solution Properties --"
    Public Property LeftSet As DataView
        Get
            LeftSet = _left.AsDataView
        End Get
        Set(value As DataView)
            OnPropertyChanged("LeftSet")
        End Set
    End Property

    Public Property RightSet As DataView
        Get
            RightSet = _right.AsDataView
        End Get
        Set(value As DataView)
            OnPropertyChanged("RightSet")
        End Set
    End Property

    Public Property DifferSet As DataView
        Get
            DifferSet = _differ.AsDataView
        End Get
        Set(value As DataView)
            OnPropertyChanged("DifferSet")
        End Set
    End Property

    Public Property MatchSet As DataView
        Get
            MatchSet = _match.AsDataView
        End Get
        Set(value As DataView)
            OnPropertyChanged("MatchSet")
        End Set
    End Property
    Public Property ReconSet As DataView
        Get
            ReconSet = _recon.AsDataView
        End Get
        Set(value As DataView)
            OnPropertyChanged("ReconSet")
        End Set
    End Property
#End Region

#Region "-- Commands --"
    Private Sub btnOpenFile_Click(sender As Object, e As RoutedEventArgs)
        Using sd As New SystemDialog
            _solutionPath = sd.OpenFile()
        End Using
        If Not String.IsNullOrEmpty(_solutionPath) Then
            _solution = Solution.LoadSolution(_solutionPath)

        End If
    End Sub

    Private Sub btnSaveFile_Click(sender As Object, e As RoutedEventArgs)
        Exit Sub

    End Sub
#End Region

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        DataContext = Me

        'Test Invoice
        '_filters.Add(New Filter With {.DataSourceName = "", .FilterColumns = {"ExportId"}, .FilterOption = .FilterOption.NonBlankOrZero})
        '_completenessComparisions.Add(New Comparision With {.LeftColumn = "QB Export ID", .RightColumn = "TxnId", .ComparisionOption = 0, .ComparisionMethod = CompareMethod.GetMethod("String Equals")})
        '_matchingComparisions.Add(New Comparision With {.LeftColumn = "Total", .RightColumn = "Subtotal", .ComparisionOption = 2, .ComparisionMethod = CompareMethod.GetMethod("Single Equals")})
        Using ge As New GetExcel
            '_left = ge.GetList("C:\Users\Peter Grillo\source\repos\Test.xlsx", "ST PL")
            '_right = ge.GetList("C:\Users\Peter Grillo\source\repos\Test.xlsx", "QB PL")
        End Using
        Task.Factory.StartNew(Sub() Test())
    End Sub

    Private Sub Test()
        _leftDetails.Reset()
        _rightDetails.Reset()
        _differ.Reset()
        DifferSet = _differ.AsDataView
        _match.Reset()
        MatchSet = _match.AsDataView
        _recon.Reset()
        ReconSet = _recon.AsDataView

        If _left.Rows.Count = 0 Or _right.Rows.Count = 0 Then Exit Sub
        'ApplyFilterX()

        'DoAggregation()
        _recon.Columns.Add(New DataColumn With {.ColumnName = "Description", .DataType = GetType(String)})
        _recon.Columns.Add(New DataColumn With {.ColumnName = "_", .DataType = GetType(String)})
        _recon.Columns.Add(New DataColumn With {.ColumnName = "__", .DataType = GetType(String)})
        _recon.Columns.Add(New DataColumn With {.ColumnName = "Count", .DataType = GetType(Integer)})
        _recon.Columns.Add(New DataColumn With {.ColumnName = "Total", .DataType = GetType(Double)})

        Dim stTotal As Double = Aggregate r In _left.AsEnumerable
                          Into Sum(CDbl(r("SubTotal")))

        AddReconRow("ServiceTitan", "", "", _left.Rows.Count, stTotal)

        Dim pending As Integer = Aggregate r In _left.AsEnumerable
                                     Where r("Status") = 2
                          Into Count
        AddReconRow("-Pending", "", "", pending, Nothing)

        LeftSet = _left.AsDataView
        RightSet = _right.AsDataView
        DifferSet = _differ.AsDataView
        MatchSet = _match.AsDataView
        ReconSet = _recon.AsDataView

    End Sub

    Private Sub AddReconRow(col1 As String, col2 As String, col3 As String, col4 As Integer, col5 As Double)
        Dim cols As New List(Of Object)
        cols.Add(col1)
        cols.Add(col2)
        cols.Add(col3)
        cols.Add(col4)
        cols.Add(col5)
        _recon.Rows.Add(cols.ToArray)
    End Sub

    Private Sub DoAggregation()
        _rightDetails = _right.Copy
        _right.Reset()
        Dim removes As New List(Of DataRow)
        'Dim colHeaders As DataColumn() = (From a As Aggregate In _aggregates
        '                                  Select New DataColumn With {
        '                          .ColumnName = a.AggregateName,
        '                          .DataType = GetType(String)
        '                          }
        '                     ).ToArray
        '_right.Columns.AddRange(colHeaders)

        For Each a In _aggregates
            'Todo Match DataSourceName

            Dim groups = From row In _rightDetails.AsEnumerable
                         Group row By GroupKey = row.Field(Of Double)("Customer ID") Into AggGroup = Group
                         Select New With {
            Key GroupKey,
            .EntryTotal = AggGroup.Sum(Function(r) r.Field(Of Double)("Entry")),
            .EntryAvg = AggGroup.Average(Function(r) r.Field(Of Double)("Entry")),
            .EntryCount = AggGroup.Count(Function(r) r.Field(Of Double)("Entry"))
            }

            Dim colHeaders As DataColumn() = (From ao As AggregateOperation In a.AggregateOperations
                                              Select New DataColumn With {
                                      .ColumnName = ao.AggregateColumn,
                                      .DataType = IIf(ao.Operation = AggregateOperation.AggregateFunction.Count, GetType(Integer), GetType(String))
                                      }
                                 ).ToArray
            _right.Columns.AddRange(colHeaders)
            _right.Columns.Add(New DataColumn With {.ColumnName = "GroupKey", .DataType = GetType(String)})
            For Each g In groups
                Dim GroupRowData As New List(Of String)
                For Each ao As AggregateOperation In a.AggregateOperations
                    GroupRowData.Add(InvokeGet(g, ao.AggregateColumn))
                Next
                GroupRowData.Add(g.GroupKey)
                _right.Rows.Add(GroupRowData.ToArray)

            Next
        Next

    End Sub



    Private Function Reconcile() As Tuple(Of Data.DataRow, Data.DataRow)
        For Each leftitem In _left.AsEnumerable
            For Each rightitem In _right.AsEnumerable
                Dim matchingResult = DoCompare(leftitem, rightitem, _matchingComparisions)
                If matchingResult.Item1 = True Then
                    InsertRow(leftitem, rightitem, _matchingComparisions, True)
                    Return New Tuple(Of Data.DataRow, Data.DataRow)(leftitem, rightitem)
                Else
                    Dim completenessResult = DoCompare(leftitem, rightitem, _completenessComparisions)
                    If completenessResult.Item1 Then  'ids match but not complete match
                        InsertRow(leftitem, rightitem, _matchingComparisions, False, matchingResult.Item2)
                        Return New Tuple(Of Data.DataRow, Data.DataRow)(leftitem, rightitem)
                    End If
                End If
            Next
        Next
        Return Nothing
    End Function

    Private Function ReconcileX() As Tuple(Of Data.DataRow, Data.DataRow)
        Parallel.ForEach(_left.AsEnumerable,
                         Sub(leftitem)
                             Parallel.ForEach(_right.AsEnumerable,
                                              Function(rightitem)
                                                  Dim matchingResult = DoCompare(leftitem, rightitem, _matchingComparisions)
                                                  If matchingResult.Item1 = True Then
                                                      InsertRow(leftitem, rightitem, _matchingComparisions, True)
                                                      Return New Tuple(Of Data.DataRow, Data.DataRow)(leftitem, rightitem)
                                                      Exit Function
                                                  Else
                                                      Dim completenessResult = DoCompare(leftitem, rightitem, _completenessComparisions)
                                                      If completenessResult.Item1 Then  'ids match but not complete match
                                                          InsertRow(leftitem, rightitem, _matchingComparisions, False, matchingResult.Item2)
                                                          Return New Tuple(Of Data.DataRow, Data.DataRow)(leftitem, rightitem)
                                                      End If
                                                  End If
                                              End Function)
                         End Sub)
        Return Nothing
    End Function

    Private Function DoCompare(leftitem As Data.DataRow, rightitem As Data.DataRow, comparisions As List(Of Comparision), Optional IsAll As Boolean = True) As Tuple(Of Boolean, List(Of String))
        Dim errorCols As New List(Of String)
        For Each compare In comparisions
            If IsDBNull(leftitem(compare.LeftColumn)) OrElse
            IsDBNull(rightitem(compare.RightColumn)) OrElse
            Not compare.ComparisionMethod.Method.Invoke(leftitem(compare.LeftColumn), rightitem(compare.RightColumn), compare.ComparisionOption) Then
                errorCols.Add(compare.LeftColumn)
                'rightitem.RowError = compare.ComparisionMethod.Name
                If Not IsAll Then Return New Tuple(Of Boolean, List(Of String))(False, errorCols)
            End If
        Next
        Return New Tuple(Of Boolean, List(Of String))(errorCols.Count = 0, errorCols)
    End Function


    Private Sub InsertRow(leftitem As Data.DataRow, rightitem As Data.DataRow, comparisions As List(Of Comparision), isMatch As Boolean, Optional errorCols As List(Of String) = Nothing)
        Dim currentTable As DataTable = IIf(isMatch, _match, _differ)
        If currentTable.Rows.Count = 0 Then
            Dim colHeaders As DataColumn() = (From c As Comparision In comparisions
                                              Select New DataColumn With {
                                  .ColumnName = c.LeftColumn + ":" + c.RightColumn,
                                  .DataType = IIf(isMatch, _left.Columns(c.LeftColumn).DataType, GetType(String))
                                  }
                             ).Distinct(New ColNameComparer).ToArray
            currentTable.Columns.AddRange(colHeaders)
        End If

        currentTable.Rows.Add(comparisions.Select(Function(s)
                                                      If isMatch Then
                                                          Return leftitem(s.LeftColumn)
                                                      Else
                                                          If errorCols Is Nothing Then
                                                              Return leftitem(s.LeftColumn)
                                                          Else
                                                              If errorCols.Contains(s.LeftColumn) Then
                                                                  Return leftitem(s.LeftColumn).ToString + "<>" + rightitem(s.RightColumn).ToString
                                                              Else
                                                                  Return leftitem(s.LeftColumn)
                                                              End If
                                                          End If
                                                      End If
                                                  End Function).ToArray)

    End Sub

    Private Sub btnLoad_Click(sender As Object, e As RoutedEventArgs)
        Dim qbfc As New GetQBFC
        LeftSet = qbfc.GetReport("P/L Detail").AsDataView
        'LeftSet = qbfc.GetList("Item")
    End Sub

    Private Sub btnMatch_Click(sender As Object, e As RoutedEventArgs)
        'DifferSet = _left
        'MatchSet = _right
        Test()
    End Sub

End Class


