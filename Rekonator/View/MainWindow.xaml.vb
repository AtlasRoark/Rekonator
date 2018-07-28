Imports System.Collections.Concurrent
Imports System.ComponentModel
Imports System.Data
Imports System.Threading
Imports Dynamitey.Dynamic

Partial Class MainWindow
    Implements INotifyPropertyChanged

    Private _aggregates As New List(Of Aggregate)
    Private _completenessComparisions As New List(Of Comparision)
    Private _matchingComparisions As New List(Of Comparision)
    Private _solutionPath As String = String.Empty
    Private _solution As Solution
    'Cant do notify prop change on datatables.  
    Private _left As New DataTable
    Private _right As New DataTable
    Private _leftDetails As New DataTable
    Private _rightDetails As New DataTable
    Private _differ As New DataTable
    Private _match As New DataTable
    Private _isControlPressed As Boolean = False

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


        Using sql As New SQL
            Dim recon As Reconciliation = _solution.Reconciliations(0)
            _match = sql.GetDataTable(Reconciliation.GetMatchSelect(recon))
            MatchSet = _match.AsDataView
            _differ = sql.GetDataTable(Reconciliation.GetDifferSelect(recon))
            DifferSet = _match.AsDataView
            _left = sql.GetDataTable(Reconciliation.GetLeftSelect(recon))
            LeftSet = _left.AsDataView
            _right = sql.GetDataTable(Reconciliation.GetRightSelect(recon))
            RightSet = _right.AsDataView
        End Using
        'DoAggregation()
        'Do
        '    Debug.Print(_left.Rows.Count.ToString)
        '    Dim remove As Tuple(Of Data.DataRow, Data.DataRow) = Reconcile()
        '    If remove Is Nothing Then Exit Do
        '    _left.Rows.Remove(remove.Item1)
        '    _right.Rows.Remove(remove.Item2)
        'Loop

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


    'Private Function Reconcile() As Tuple(Of Data.DataRow, Data.DataRow)
    '    For Each leftitem In _left.AsEnumerable
    '        For Each rightitem In _right.AsEnumerable
    '            Dim matchingResult = DoCompare(leftitem, rightitem, _matchingComparisions)
    '            If matchingResult.Item1 = True Then
    '                InsertRow(leftitem, rightitem, _matchingComparisions, True)
    '                Return New Tuple(Of Data.DataRow, Data.DataRow)(leftitem, rightitem)
    '            Else
    '                Dim completenessResult = DoCompare(leftitem, rightitem, _completenessComparisions)
    '                If completenessResult.Item1 Then  'ids match but not complete match
    '                    InsertRow(leftitem, rightitem, _matchingComparisions, False, matchingResult.Item2)
    '                    Return New Tuple(Of Data.DataRow, Data.DataRow)(leftitem, rightitem)
    '                End If
    '            End If
    '        Next
    '    Next
    '    Return Nothing
    'End Function


    'Private Function DoCompare(leftitem As Data.DataRow, rightitem As Data.DataRow, comparisions As List(Of Comparision), Optional IsAll As Boolean = True) As Tuple(Of Boolean, List(Of String))
    '    Dim errorCols As New List(Of String)
    '    For Each compare In comparisions
    '        If IsDBNull(leftitem(compare.LeftColumn)) OrElse
    '        IsDBNull(rightitem(compare.RightColumn)) OrElse
    '        Not compare.ComparisionMethod.Method.Invoke(leftitem(compare.LeftColumn), rightitem(compare.RightColumn), compare.ComparisionOption) Then
    '            errorCols.Add(compare.LeftColumn)
    '            'rightitem.RowError = compare.ComparisionMethod.Name
    '            If Not IsAll Then Return New Tuple(Of Boolean, List(Of String))(False, errorCols)
    '        End If
    '    Next
    '    Return New Tuple(Of Boolean, List(Of String))(errorCols.Count = 0, errorCols)
    'End Function


    'Private Sub InsertRow(leftitem As Data.DataRow, rightitem As Data.DataRow, comparisions As List(Of Comparision), isMatch As Boolean, Optional errorCols As List(Of String) = Nothing)
    '    Dim currentTable As DataTable = IIf(isMatch, _match, _differ)
    '    If currentTable.Rows.Count = 0 Then
    '        Dim colHeaders As DataColumn() = (From c As Comparision In comparisions
    '                                          Select New DataColumn With {
    '                              .ColumnName = c.LeftColumn + ":" + c.RightColumn,
    '                              .DataType = IIf(isMatch, _left.Columns(c.LeftColumn).DataType, GetType(String))
    '                              }
    '                         ).Distinct(New ColNameComparer).ToArray
    '        currentTable.Columns.AddRange(colHeaders)
    '    End If

    '    currentTable.Rows.Add(comparisions.Select(Function(s)
    '                                                  If isMatch Then
    '                                                      Return leftitem(s.LeftColumn)
    '                                                  Else
    '                                                      If errorCols Is Nothing Then
    '                                                          Return leftitem(s.LeftColumn)
    '                                                      Else
    '                                                          If errorCols.Contains(s.LeftColumn) Then
    '                                                              Return leftitem(s.LeftColumn).ToString + "<>" + rightitem(s.RightColumn).ToString
    '                                                          Else
    '                                                              Return leftitem(s.LeftColumn)
    '                                                          End If
    '                                                      End If
    '                                                  End If
    '                                              End Function).ToArray)

    'End Sub

    Private Sub btnLoad_Click(sender As Object, e As RoutedEventArgs)
        Dim qbfc As New GetQBFC
        LeftSet = qbfc.GetReport("P/L Detail").AsDataView
        'LeftSet = qbfc.GetList("Item")
    End Sub

    Private Sub btnMatch_Click(sender As Object, e As RoutedEventArgs)
        Task.Factory.StartNew(Sub() Configure("Huber"))
        MessageLog.Clear()
    End Sub

    Private Function Configure(testName As String) As Task(Of Object)

        Dim leftRS As ReconSource = Nothing
        Dim rightRS As ReconSource = Nothing
        Select Case testName
            Case "Test 1"

                'Test 1
                Dim excelDS As DataSource = DataSource.GetDataSource("Excel")
                Dim excelParam As New Dictionary(Of String, String)
                excelParam.Add("FilePath", "C:\Users\Peter Grillo\source\repos\Test.xlsx")
                excelParam.Add("Worksheet", "ST Items")
                leftRS = New ReconSource With
                    {.ReconDataSource = excelDS,
                    .ReconTable = "aset",
                    .IsLoaded = False,
                    .Parameters = excelParam}

                excelParam = New Dictionary(Of String, String)
                excelParam.Add("FilePath", "C:\Users\Peter Grillo\source\repos\Test.xlsx")
                excelParam("Worksheet") = "Item"
                rightRS = New ReconSource With
                    {.ReconDataSource = excelDS,
                    .ReconTable = "bset",
                    .IsLoaded = False,
                    .Parameters = excelParam}

                _completenessComparisions.Add(New Comparision With {.LeftColumn = "ST Export ID", .RightColumn = "Item ID", .ComparisionTest = ComparisionType.TextCaseEquals})
                '_matchingComparisions.Add(New Comparision With {.LeftColumn = "ST Export ID", .RightColumn = "Item ID", .ComparisionTest = ComparisionType.TextCaseEquals})
                _matchingComparisions.Add(New Comparision With {.LeftColumn = "ST Item Name", .RightColumn = "Item Name", .ComparisionTest = ComparisionType.TextEquals})
                _matchingComparisions.Add(New Comparision With {.LeftColumn = "ST Item Qty", .RightColumn = "Item Qty", .Percision = 4, .ComparisionTest = ComparisionType.NumberEquals})
                _matchingComparisions.Add(New Comparision With {.LeftColumn = "ST Item Price", .RightColumn = "Item Price", .Percision = 2, .ComparisionTest = ComparisionType.NumberEquals})
                _matchingComparisions.Add(New Comparision With {.LeftColumn = "ST Item Reorder", .RightColumn = "Item Reorder", .ComparisionTest = ComparisionType.NumberEquals})
                _matchingComparisions.Add(New Comparision With {.LeftColumn = "ST Service Date", .RightColumn = "Service Date", .ComparisionTest = ComparisionType.DateEquals})
                Reconciliation.Add("Test Set", leftRS, rightRS, _completenessComparisions, _matchingComparisions)
            Case "Huber"
                Dim excelDS As DataSource = DataSource.GetDataSource("Excel")
                Dim excelParam As New Dictionary(Of String, String)
                excelParam.Add("FilePath", "C:\Users\Peter Grillo\Downloads\March Activity.xls")
                excelParam.Add("Worksheet", "ST Inv")
                leftRS = New ReconSource With
                    {.ReconDataSource = excelDS,
                    .ReconTable = "stinvoice",
                    .IsLoaded = True,
                    .Parameters = excelParam,
                    .Where = "NOT (ISNULL(x!.[ExportId], '')='' OR x!.[ExportId]='0')"}

                excelParam = New Dictionary(Of String, String)
                excelParam.Add("FilePath", "C:\Users\Peter Grillo\Downloads\March Activity.xls")
                excelParam("Worksheet") = "Mar Inv"
                rightRS = New ReconSource With
                    {.ReconDataSource = excelDS,
                    .ReconTable = "qbinvoice",
                    .IsLoaded = True,
                    .Parameters = excelParam,
                    .Where = "NOT (ISNULL(x!.[SubTotal], '')='' OR x!.[SubTotal]='0')"}

                _completenessComparisions.Add(New Comparision With {.LeftColumn = "QB Export ID", .RightColumn = "TxnId", .ComparisionTest = ComparisionType.TextCaseEquals})
                _matchingComparisions.Add(New Comparision With {.LeftColumn = "Total", .RightColumn = "Subtotal", .Percision = 2, .ComparisionTest = ComparisionType.NumberEquals})
                Reconciliation.Add("March Invoices", leftRS, rightRS, _completenessComparisions, _matchingComparisions)

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
        End Select

        _solution = New Solution With {.SolutionName = testName, .Reconciliations = Reconciliation.Reconciliations}
        If Not leftRS.IsLoaded Then
            Using ge As New GetExcel
                leftRS.IsLoaded = ge.Load(leftRS)
            End Using
        End If
        If Not rightRS.IsLoaded Then
            Using ge As New GetExcel
                rightRS.IsLoaded = ge.Load(rightRS)
            End Using
        End If 'PL Reconcile
        'Using ge As New GetExcel
        'Dim _left = ge.MakeReconSource("C:\Users\Peter Grillo\Downloads\March Activity.xlsx", "ST PL") As ReconSource
        ''_right = ge.GetList("C:\Users\Peter Grillo\Downloads\March Activity Sorted.xlsx", "QB PL")
        'End Using
        Using sql As New SQL
            Dim rs As ReconSource = _solution.Reconciliations(0).LeftReconSource
            _left = sql.GetDataTable(ReconSource.GetSelect(rs))
            LeftSet = _left.AsDataView
            rs = _solution.Reconciliations(0).RightReconSource
            _right = sql.GetDataTable(ReconSource.GetSelect(rs))
            RightSet = _right.AsDataView
        End Using


        Test()
    End Function

    Public Sub AddMessage(messageText As String, isError As Boolean)
        If Not BottomFlyout.IsOpen Then BottomFlyout.IsOpen = True
        MessageLog.Add(New MessageEntry With {.MessageText = messageText, .IsError = isError})
        OnPropertyChanged("MessageLog")
        lbMessageLog.SelectedIndex = lbMessageLog.Items.Count - 1
        lbMessageLog.ScrollIntoView(lbMessageLog.SelectedItem)
    End Sub

    Private Sub DataGridCell_PreviewKeyDown(sender As Object, e As KeyEventArgs)

        Dim cell As DataGridCell = TryCast(sender, DataGridCell)

        If e.Key = Key.LeftCtrl OrElse e.Key = Key.RightCtrl Then
            _isControlPressed = True
        End If

        If _isControlPressed AndAlso e.Key = Key.C Then
            If TypeOf cell.Content Is TextBlock Then Clipboard.SetText((TryCast(cell.Content, TextBlock)).Text)
            _isControlPressed = False
            e.Handled = True
        End If
    End Sub
End Class



