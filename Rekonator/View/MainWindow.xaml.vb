Imports System.Collections.Concurrent
Imports System.ComponentModel
Imports System.Data
Imports System.Threading
Imports Dynamitey.Dynamic

Partial Class MainWindow
    Implements INotifyPropertyChanged

    Private _filters As New List(Of Filter)
    Private _aggregates As New List(Of Aggregate)
    Private _completenessComparisions As New List(Of Comparision)
    Private _matchingComparisions As New List(Of Comparision)
    Private _solutionPath As String = String.Empty
    Private _solution As Solution
    'Cant do notify prop change on datatables.  
    Private _left As New DataTable
    Private _leftRS As ReconSource
    Private _rightRS As ReconSource
    Private _right As New DataTable
    Private _leftDetails As New DataTable
    Private _rightDetails As New DataTable
    Private _differ As New DataTable
    Private _match As New DataTable

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
#End Region
    Public Property MessageLog As List(Of MessageEntry)
        Get
            MessageLog = _messages
        End Get
        Set(value As List(Of MessageEntry))
            OnPropertyChanged("MessageLog")
        End Set
    End Property
    Private _messages As New List(Of MessageEntry)

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
        Application.MessageFunc = AddressOf AddMessage
        Using m As New Mock
            Dim sm As SettingsModel = m.MockLoadSettings() 'Model for App Settings
            DataSources = sm.Datasources

            '_solution = Task.Run(Function() m.MockLoadSolutionAsync(Me)).GetAwaiter().GetResult() 'Model for Solution
        End Using

    End Sub

    Private Sub Test()
        _leftDetails.Reset()
        _rightDetails.Reset()
        _differ.Reset()
        DifferSet = _differ.AsDataView
        _match.Reset()
        MatchSet = _match.AsDataView

        If _left.Rows.Count = 0 Or _right.Rows.Count = 0 Then Exit Sub
        ApplyFilterX()


        'DoAggregation()
        Do
            Debug.Print(_left.Rows.Count.ToString)
            Dim remove As Tuple(Of Data.DataRow, Data.DataRow) = Reconcile()
            If remove Is Nothing Then Exit Do
            _left.Rows.Remove(remove.Item1)
            _right.Rows.Remove(remove.Item2)
        Loop

        LeftSet = _left.AsDataView
        RightSet = _right.AsDataView
        DifferSet = _differ.AsDataView
        MatchSet = _match.AsDataView
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

    Private Sub ApplyFilter()
        Dim removes As New List(Of DataRow)
        For Each f In _filters
            'Todo Match DataSourceName
            For Each leftitem In _left.AsEnumerable
                For Each col In f.FilterColumns
                    Select Case f.FilterOption
                        Case FilterOption.NonZero
                            Dim result As Integer
                            If Integer.TryParse(leftitem(col), result) Then
                                If result = 0 Then
                                    removes.Add(leftitem)
                                    Continue For
                                End If
                            End If
                        Case FilterOption.NonBlankOrZero
                            If IsDBNull(leftitem(col)) Then
                                removes.Add(leftitem)
                                Continue For
                            End If
                            Dim result As Integer
                            If Integer.TryParse(leftitem(col), result) Then
                                If result = 0 Then
                                    removes.Add(leftitem)
                                    Continue For
                                End If
                            End If

                            If String.IsNullOrWhiteSpace(leftitem(col)) Then
                                removes.Add(leftitem)
                                Continue For
                            End If
                    End Select
                Next
            Next
        Next

        For Each r In removes
            _left.Rows.Remove(r)
        Next
    End Sub

    Private Sub ApplyFilterX()
        Dim removes As New ConcurrentStack(Of DataRow)
        'Todo Match DataSourceName
        For Each f In _filters

            Parallel.ForEach(_left.AsEnumerable,
                             Sub(leftitem)
                                 For Each col In f.FilterColumns
                                     Select Case f.FilterOption
                                         Case FilterOption.NonZero
                                             Dim result As Integer
                                             If Integer.TryParse(leftitem(col), result) Then
                                                 If result = 0 Then
                                                     removes.Push(leftitem)
                                                     Continue For
                                                 End If
                                             End If
                                         Case FilterOption.NonBlankOrZero
                                             If IsDBNull(leftitem(col)) Then
                                                 removes.Push(leftitem)
                                                 Continue For
                                             End If
                                             Dim result As Integer
                                             If Integer.TryParse(leftitem(col), result) Then
                                                 If result = 0 Then
                                                     removes.Push(leftitem)
                                                     Continue For
                                                 End If
                                             End If

                                             If String.IsNullOrWhiteSpace(leftitem(col)) Then
                                                 removes.Push(leftitem)
                                                 Continue For
                                             End If
                                     End Select
                                 Next

                             End Sub)
        Next

        For Each r In removes
            _left.Rows.Remove(r)
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
        Task.Factory.StartNew(Sub() Configure("Huber"))
        MessageLog.Clear
    End Sub

    Private Function Configure(v As String) As Task(Of Object)
        'Test 1
        '_completenessComparisions.Add(New Comparision With {.LeftColumn = "ST Export ID", .RightColumn = "Item ID", .ComparisionOption = 1, .ComparisionMethod = CompareMethod.GetMethod("String Equals")})
        '_matchingComparisions.Add(New Comparision With {.LeftColumn = "ST Export ID", .RightColumn = "Item ID", .ComparisionOption = 1, .ComparisionMethod = CompareMethod.GetMethod("String Equals")})
        '_matchingComparisions.Add(New Comparision With {.LeftColumn = "ST Item Name", .RightColumn = "Item Name", .ComparisionOption = 0, .ComparisionMethod = CompareMethod.GetMethod("String Equals")})
        '_matchingComparisions.Add(New Comparision With {.LeftColumn = "ST Item Qty", .RightColumn = "Item Qty", .ComparisionOption = 4, .ComparisionMethod = CompareMethod.GetMethod("Single Equals")})
        '_matchingComparisions.Add(New Comparision With {.LeftColumn = "ST Item Price", .RightColumn = "Item Price", .ComparisionOption = 2, .ComparisionMethod = CompareMethod.GetMethod("Single Equals")})
        '_matchingComparisions.Add(New Comparision With {.LeftColumn = "ST Item Reorder", .RightColumn = "Item Reorder", .ComparisionOption = 1, .ComparisionMethod = CompareMethod.GetMethod("Integer Equals")})
        '_matchingComparisions.Add(New Comparision With {.LeftColumn = "ST Service Date", .RightColumn = "Service Date", .ComparisionOption = 0, .ComparisionMethod = CompareMethod.GetMethod("Date Equals")})
        'Using ge As New GetExcel
        '    _left = ge.GetList("C:\Users\Peter Grillo\source\repos\Test.xlsx", "ST Items")
        '    _right = ge.GetList("C:\Users\Peter Grillo\source\repos\Test.xlsx", "Item")
        'End Using

        'Test Agg
        '_filters.Add(New Filter With {.DataSourceName = "", .FilterColumns = {"Balance"}, .FilterOption = .FilterOption.NonBlankOrZero})
        'Dim aggOps As New List(Of AggregateOperation)
        'aggOps.Add(New AggregateOperation With {.SourceColumn = "Entry", .AggregateColumn = "EntryTotal", .Operation = AggregateOperation.AggregateFunction.Sum})
        'aggOps.Add(New AggregateOperation With {.SourceColumn = "Entry", .AggregateColumn = "EntryAvg", .Operation = AggregateOperation.AggregateFunction.Avg})
        'aggOps.Add(New AggregateOperation With {.SourceColumn = "Entry", .AggregateColumn = "EntryCount", .Operation = AggregateOperation.AggregateFunction.Count})
        '_aggregates.Add(New Aggregate With {.DataSourceName = "", .GroupByCoumns = {"Customer ID"}, .AggregateOperations = aggOps})
        '_completenessComparisions.Add(New Comparision With {.LeftColumn = "Customer ID", .RightColumn = "GroupKey", .ComparisionOption = 1, .ComparisionMethod = CompareMethod.GetMethod("String Equals")})
        '_matchingComparisions.Add(New Comparision With {.LeftColumn = "Balance", .RightColumn = "EntryTotal", .ComparisionOption = 2, .ComparisionMethod = CompareMethod.GetMethod("Single Equals")})
        '_matchingComparisions.Add(New Comparision With {.LeftColumn = "Count", .RightColumn = "EntryCount", .ComparisionOption = 0, .ComparisionMethod = CompareMethod.GetMethod("Single Equals")})
        '_matchingComparisions.Add(New Comparision With {.LeftColumn = "Avg", .RightColumn = "EntryAvg", .ComparisionOption = 9, .ComparisionMethod = CompareMethod.GetMethod("Single Equals")})
        'Using ge As New GetExcel
        '    _left = ge.GetList("C:\Users\Peter Grillo\source\repos\Test.xlsx", "Agg Balance")
        '    _right = ge.GetList("C:\Users\Peter Grillo\source\repos\Test.xlsx", "Agg Entries")
        'End Using

        'Test Huber
        Dim excelDS As DataSource = DataSource.GetDataSource("Excel")
        Dim excelParam As New Dictionary(Of String, String)
        excelParam.Add("FilePath", "C:\Users\Peter Grillo\Downloads\March Activity.xls")
        excelParam.Add("Worksheet", "ST Inv")
        _filters.Add(New Filter With {.DataSourceName = "", .FilterColumns = {"ExportId"}, .FilterOption = .FilterOption.NonBlankOrZero})

        Dim leftRS As New ReconSource With {.ReconDataSource = excelDS, .ReconTable = "stinvoice", .IsLoaded = True, .Parameters = excelParam, .Filters = _filters}
        excelParam("Worksheet") = "Mar Inv"
        Dim rightRS As New ReconSource With {.ReconDataSource = excelDS, .ReconTable = "qbinvoice", .IsLoaded = True, .Parameters = excelParam}

        _completenessComparisions.Add(New Comparision With {.LeftColumn = "QB Export ID", .RightColumn = "TxnId", .ComparisionOption = 0, .ComparisionMethod = CompareMethod.GetMethod("String Equals")})
        _matchingComparisions.Add(New Comparision With {.LeftColumn = "Total", .RightColumn = "Subtotal", .ComparisionOption = 2, .ComparisionMethod = CompareMethod.GetMethod("Single Equals")})
        Reconciliation.Add("March Invoices", leftRS, rightRS, _completenessComparisions, _matchingComparisions)

        _solution = New Solution With {.SolutionName = "Huber", .Reconciliations = Reconciliation.Reconciliations}

        If Not _solution.Reconciliations(0).LeftReconSource.IsLoaded Then
            Using ge As New GetExcel
                '_leftRS = ge.MakeReconSource("C:\Users\Peter Grillo\Downloads\March Activity.xls", "ST Inv", "stinvoice")
            End Using
        End If

        If Not _solution.Reconciliations(0).RightReconSource.IsLoaded Then
            Using ge As New GetExcel
                '_rightRS = ge.MakeReconSource("C:\Users\Peter Grillo\Downloads\March Activity.xls", "Mar Inv", "qbinvoice")
            End Using
        End If 'PL Reconcile
        'Using ge As New GetExcel
        'Dim _left = ge.MakeReconSource("C:\Users\Peter Grillo\Downloads\March Activity.xlsx", "ST PL") As ReconSource
        ''_right = ge.GetList("C:\Users\Peter Grillo\Downloads\March Activity Sorted.xlsx", "QB PL")
        'End Using


        String connectionString = "Data Source=192.168.0.192;Initial Catalog=ParallelCodes;User ID=sa;Password=789";
SqlConnection con = New SqlConnection(connectionString);
SqlCommand cmd = New SqlCommand("select * from Producttbl", con);
con.Open();
SqlDataAdapter adapter = New SqlDataAdapter(cmd);
DataTable dt = New DataTable();
adapter.Fill(dt);
myDataGrid.ItemsSource = dt.DefaultView;
cmd.Dispose();
con.Close();
        'Test()
    End Function

    Public Sub AddMessage(messageText As String, isError As Boolean)
        If Not BottomFlyout.IsOpen Then BottomFlyout.IsOpen = True
        MessageLog.Add(New MessageEntry With {.MessageText = messageText, .IsError = isError})
        OnPropertyChanged("MessageLog")
        lbMessageLog.SelectedIndex = lbMessageLog.Items.Count - 1
        lbMessageLog.ScrollIntoView(lbMessageLog.SelectedItem)
    End Sub

End Class

Public Class MessageEntry
    Property IsError As Boolean
    Property MessageText As String
End Class


