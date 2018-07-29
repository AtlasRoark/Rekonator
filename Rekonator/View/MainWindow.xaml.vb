Imports System.Collections.Concurrent
Imports System.ComponentModel
Imports System.Data
Imports System.Threading
Imports Dynamitey.Dynamic

Partial Class MainWindow
    Implements INotifyPropertyChanged

    Private _solutionPath As String = String.Empty
    Private _solution As Solution 'active solution
    Private _reconciliation As Reconciliation ' active reconciliation from _solution
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
            DataSources = DataSource.DataSources
        End Get
        Set(value As List(Of DataSource))
            DataSource.DataSources = value
            OnPropertyChanged("DataSources")
        End Set
    End Property
    'Private _dataSources As List(Of DataSource)
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

    Private Enum Sets
        Left
        Right
    End Enum
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
            _match = sql.GetDataTable(Reconciliation.GetMatchSelect(_reconciliation))
            MatchSet = _match.AsDataView
            _differ = sql.GetDataTable(Reconciliation.GetDifferSelect(_reconciliation))
            DifferSet = _match.AsDataView
            _left = sql.GetDataTable(Reconciliation.GetLeftSelect(_reconciliation))
            LeftSet = _left.AsDataView
            _right = sql.GetDataTable(Reconciliation.GetRightSelect(_reconciliation))
            RightSet = _right.AsDataView
        End Using

        If _reconciliation.LeftReconSource.Aggregations IsNot Nothing Then DoAggregation(_reconciliation.LeftReconSource)
        If _reconciliation.RightReconSource.Aggregations IsNot Nothing Then DoAggregation(_reconciliation.RightReconSource)
    End Sub

    Private Sub DoAggregation(reconSource As ReconSource)

        _rightDetails = _right.Copy
        _right.Reset()
        'Dim removes As New List(Of DataRow)
        ''Dim colHeaders As DataColumn() = (From a As Aggregate In _aggregates
        ''                                  Select New DataColumn With {
        ''                          .ColumnName = a.AggregateName,
        ''                          .DataType = GetType(String)
        ''                          }
        ''                     ).ToArray
        ''_right.Columns.AddRange(colHeaders)

        'For Each a In _aggregates
        '    'Todo Match DataSourceName

        '    Dim groups = From row In _rightDetails.AsEnumerable
        '                 Group row By GroupKey = row.Field(Of Double)("Customer ID") Into AggGroup = Group
        '                 Select New With {
        '    Key GroupKey,
        '    .EntryTotal = AggGroup.Sum(Function(r) r.Field(Of Double)("Entry")),
        '    .EntryAvg = AggGroup.Average(Function(r) r.Field(Of Double)("Entry")),
        '    .EntryCount = AggGroup.Count(Function(r) r.Field(Of Double)("Entry"))
        '    }

        '    Dim colHeaders As DataColumn() = (From ao As AggregateOperation In a.AggregateOperations
        '                                      Select New DataColumn With {
        '                              .ColumnName = ao.AggregateColumn,
        '                              .DataType = IIf(ao.Operation = AggregateOperation.AggregateFunction.Count, GetType(Integer), GetType(String))
        '                              }
        '                         ).ToArray
        '    _right.Columns.AddRange(colHeaders)
        '    _right.Columns.Add(New DataColumn With {.ColumnName = "GroupKey", .DataType = GetType(String)})
        '    For Each g In groups
        '        Dim GroupRowData As New List(Of String)
        '        For Each ao As AggregateOperation In a.AggregateOperations
        '            GroupRowData.Add(InvokeGet(g, ao.AggregateColumn))
        '        Next
        '        GroupRowData.Add(g.GroupKey)
        '        _right.Rows.Add(GroupRowData.ToArray)

        '    Next
        'Next

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

    Private Sub btnMatch_Click(sender As Object, e As RoutedEventArgs)
        MessageLog.Clear()
        Task.Factory.StartNew(Sub() Configure("Test Agg")).
                                  ContinueWith(Sub() LoadReconSources()).
                                  ContinueWith(Sub() Test())
    End Sub

    Private Sub Configure(testName As String)
        Dim _completenessComparisions As New List(Of Comparision)
        Dim _matchingComparisions As New List(Of Comparision)
        Dim leftRS As ReconSource = Nothing
        Dim rightRS As ReconSource = Nothing
        Select Case testName
            Case "QB P/L Detail"
                Dim sqlDS As DataSource = DataSource.GetDataSource("SQL")
                Dim sqlParam As New Dictionary(Of String, String)
                sqlParam.Add("connectionstring", "Data Source=dbvipmaster;Initial Catalog=Prod-Lahydrojet;User ID=linxlogic;Password=6a3r3a0$")
                sqlParam.Add("schema", "lahydrojet1")
                'sqlParam.Add("commandtext", "SELECT * FROM foo")
                sqlParam.Add("commandpath", "C:\Users\Peter Grillo\Documents\SQL Server Management Studio\lah_sql_accountdetail.sql")
                sqlParam.Add("create", "CREATE TABLE lah_sql_accountdetail ( [GL Type] nvarchar(255), [GL Number] nvarchar(4000), [GL Account] nvarchar(255), [Reference] nvarchar(50), [Date] datetime2(7), [Amount] decimal(9,2), [TXN ID] nvarchar(4000), [Id] bigint, [Business Unit] nvarchar(255) )") 'Right click in SSMS result, change table name
                leftRS = New ReconSource With
                    {.ReconDataSource = sqlDS,
                    .ReconTable = "lah_sql_accountdetail",
                    .IsLoaded = True,
                    .Parameters = sqlParam}

                Dim qbDS As DataSource = DataSource.GetDataSource("QuickBooks")
                sqlParam = New Dictionary(Of String, String)
                sqlParam.Add("request", "AppendGeneralDetailReportQueryRq")
                sqlParam.Add("detailreporttype", "gdrtProfitAndLossDetail")
                rightRS = New ReconSource With
                    {.ReconDataSource = qbDS,
                    .ReconTable = "lah_qb_pldetail",
                    .IsLoaded = True,
                    .Parameters = sqlParam}

                _completenessComparisions.Add(New Comparision With {.LeftColumn = "TXN ID", .RightColumn = "TxnID", .ComparisionTest = ComparisionType.TextCaseEquals})
                _matchingComparisions.Add(New Comparision With {.LeftColumn = "Reference", .RightColumn = "Number", .ComparisionTest = ComparisionType.TextEquals})
                Reconciliation.Add("Test Set", leftRS, rightRS, _completenessComparisions, _matchingComparisions, #6/1/2018#, #6/30/2018#)

            Case "Account Detail"
                Dim sqlDS As DataSource = DataSource.GetDataSource("SQL")
                Dim sqlParam As New Dictionary(Of String, String)
                sqlParam.Add("connectionstring", "Data Source=dbvipmaster;Initial Catalog=Prod-Lahydrojet;User ID=linxlogic;Password=6a3r3a0$")
                sqlParam.Add("schema", "lahydrojet1")
                'sqlParam.Add("commandtext", "SELECT * FROM foo")
                sqlParam.Add("commandpath", "C:\Users\Peter Grillo\Documents\SQL Server Management Studio\lah_sql_accountdetail.sql")
                sqlParam.Add("create", "CREATE TABLE lah_sql_accountdetail ( [GL Name] nvarchar(255), [GL Number] nvarchar(4000), [Name] nvarchar(255), [Number] nvarchar(50), [InvoicedOn] datetime2(7), [Total] decimal(9,2), [ExportId] nvarchar(4000), [Id] bigint, [Active] bit )") 'Right click in SSMS result
                leftRS = New ReconSource With
                    {.ReconDataSource = sqlDS,
                    .ReconTable = "lah_sql_accountdetail",
                    .IsLoaded = False,
                    .Parameters = sqlParam}

                Dim excelDS As DataSource = DataSource.GetDataSource("Excel")
                Dim excelParam As New Dictionary(Of String, String)
                excelParam.Add("FilePath", "C:\Users\Peter Grillo\Downloads\1_08776f78-eaac-468c-89d4-aeedf5b65673_AccountingDetail.xlsx")
                excelParam("Worksheet") = "Combined"
                rightRS = New ReconSource With
                    {.ReconDataSource = excelDS,
                    .ReconTable = "lah_accountdetail",
                    .IsLoaded = False,
                    .Parameters = excelParam}

                _completenessComparisions.Add(New Comparision With {.LeftColumn = "ExportID", .RightColumn = "Item ID", .ComparisionTest = ComparisionType.TextCaseEquals})
                _matchingComparisions.Add(New Comparision With {.LeftColumn = "ST Item Name", .RightColumn = "Item Name", .ComparisionTest = ComparisionType.TextEquals})
                _matchingComparisions.Add(New Comparision With {.LeftColumn = "ST Item Qty", .RightColumn = "Item Qty", .Percision = 4, .ComparisionTest = ComparisionType.NumberEquals})
                _matchingComparisions.Add(New Comparision With {.LeftColumn = "ST Item Price", .RightColumn = "Item Price", .Percision = 2, .ComparisionTest = ComparisionType.NumberEquals})
                _matchingComparisions.Add(New Comparision With {.LeftColumn = "ST Item Reorder", .RightColumn = "Item Reorder", .ComparisionTest = ComparisionType.NumberEquals})
                _matchingComparisions.Add(New Comparision With {.LeftColumn = "ST Service Date", .RightColumn = "Service Date", .ComparisionTest = ComparisionType.DateEquals})
                Reconciliation.Add("Test Set", leftRS, rightRS, _completenessComparisions, _matchingComparisions, #5/1/2018#, #5/31/2018#)
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

            Case "Test Agg"
                Dim excelDS As DataSource = DataSource.GetDataSource("Excel")
                Dim excelParam As New Dictionary(Of String, String)
                excelParam.Add("FilePath", "C:\Users\Peter Grillo\source\repos\Test.xlsx")
                excelParam.Add("Worksheet", "Agg Balance")
                leftRS = New ReconSource With
                    {.ReconDataSource = excelDS,
                    .ReconTable = "aggbalance",
                    .IsLoaded = False,
                    .Parameters = excelParam,
                    .Where = "NOT (ISNULL(x!.[Balance], '')='' OR x!.[Balance]='0')"}

                Dim aggOps As New List(Of AggregateOperation)
                aggOps.Add(New AggregateOperation With {.SourceColumn = "Entry", .AggregateColumn = "EntryTotal", .Operation = AggregateOperation.AggregateFunction.Sum})
                aggOps.Add(New AggregateOperation With {.SourceColumn = "Entry", .AggregateColumn = "EntryAvg", .Operation = AggregateOperation.AggregateFunction.Avg})
                aggOps.Add(New AggregateOperation With {.SourceColumn = "Entry", .AggregateColumn = "EntryCount", .Operation = AggregateOperation.AggregateFunction.Count})
                Dim _aggregates As New List(Of Aggregate)

                _aggregates.Add(New Aggregate With {.DataSourceName = "", .GroupByColumns = {"Customer ID"}, .AggregateOperations = aggOps})

                excelParam = New Dictionary(Of String, String)
                excelParam.Add("FilePath", "C:\Users\Peter Grillo\source\repos\Test.xlsx")
                excelParam("Worksheet") = "Agg Entries"
                rightRS = New ReconSource With
                    {.ReconDataSource = excelDS,
                    .ReconTable = "aggentries",
                    .IsLoaded = False,
                    .Parameters = excelParam,
                    .Aggregations = _aggregates}


                _completenessComparisions.Add(New Comparision With {.LeftColumn = "Customer ID", .RightColumn = "GroupKey", .ComparisionTest = ComparisionType.TextCaseEquals})
                _matchingComparisions.Add(New Comparision With {.LeftColumn = "Balance", .RightColumn = "EntryTotal", .Percision = 2, .ComparisionTest = ComparisionType.NumberEquals})
                _matchingComparisions.Add(New Comparision With {.LeftColumn = "Count", .RightColumn = "EntryCount", .Percision = 0, .ComparisionTest = ComparisionType.NumberEquals})
                _matchingComparisions.Add(New Comparision With {.LeftColumn = "Ang", .RightColumn = "EntryAvg", .Percision = 9, .ComparisionTest = ComparisionType.NumberEquals})
                Reconciliation.Add("March Invoices", leftRS, rightRS, _completenessComparisions, _matchingComparisions)

        End Select

        _solution = New Solution With {.SolutionName = testName, .Reconciliations = Reconciliation.Reconciliations}
        _reconciliation = _solution.Reconciliations(0)

    End Sub

    Private Sub LoadReconSources()
        If Not _reconciliation.LeftReconSource.IsLoaded Then LoadReconSource(_reconciliation.LeftReconSource)
        If Not _reconciliation.RightReconSource.IsLoaded Then LoadReconSource(_reconciliation.RightReconSource)

        Using sql As New SQL
            Dim rs As ReconSource = _reconciliation.LeftReconSource
            _left = sql.GetDataTable(ReconSource.GetSelect(rs)) ', _reconciliation.FromDate, _reconciliation.ToDate)
            LeftSet = _left.AsDataView
            rs = _reconciliation.RightReconSource
            _right = sql.GetDataTable(ReconSource.GetSelect(rs)) ', _reconciliation.FromDate, _reconciliation.ToDate)
            RightSet = _right.AsDataView
        End Using
    End Sub

    Private Sub LoadReconSource(reconSource As ReconSource)
        Select Case reconSource.ReconDataSource.DataSourceName
            Case "Excel"
                Using excel As New GetExcel
                    reconSource.IsLoaded = excel.Load(reconSource)
                End Using
            Case "SQL"
                Using sql As New GetSQL
                    reconSource.IsLoaded = sql.Load(reconSource, _reconciliation.FromDate, _reconciliation.ToDate)
                End Using
            Case "QuickBooks"
                Using qbd As New GetQBD
                    reconSource.IsLoaded = qbd.LoadReport(reconSource, _reconciliation.FromDate, _reconciliation.ToDate)
                End Using
        End Select
    End Sub

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



