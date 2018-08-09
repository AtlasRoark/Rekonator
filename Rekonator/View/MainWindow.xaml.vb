Imports System.Data
Imports MahApps.Metro.Controls.Dialogs
Imports Rekonator

Partial Class MainWindow

    Private _vm As New MainViewModel
    Private _solutionPath As String = String.Empty
    'Cant do notify prop change on datatables.  
    Private _left As New DataTable
    Private _leftDetails As New DataTable
    Private _right As New DataTable
    Private _rightDetails As New DataTable
    Private _differ As New DataTable
    Private _match As New DataTable
    Private _isControlPressed As Boolean = False



#Region "-- Solution Properties --"

    'Public Property ObsSolution As ObservableCollection(Of Solution) 'active solution
    '    Get
    '        ObsSolution = _obsSolution
    '    End Get
    '    Set(value As ObservableCollection(Of Solution))
    '        _obsSolution = value
    '        Reconciliation = _solution.Reconciliations(0)
    '        Me.DataContext = Me
    '        Me.dspLeft.DataContext = _solution.Reconciliations(0).LeftReconSource
    '        Me.dspRight.DataContext = _solution.Reconciliations(0).RightReconSource
    '        OnPropertyChanged("ObsSolution")
    '    End Set
    'End Property
    'Private _obsSolution As ObservableCollection(Of Solution)




#End Region



#Region "-- Commands --"
    Private Sub ButtonNew_Click(sender As Object, e As RoutedEventArgs)
        _vm.Solution = Solution.MakeNewReconcilition()
        Me.TopFlyout.IsOpen = True
    End Sub

    Private Sub ButtonOpenFile_Click(sender As Object, e As RoutedEventArgs)
        Using sd As New SystemDialog
            _solutionPath = sd.OpenFile()
        End Using
        If Not String.IsNullOrWhiteSpace(_solutionPath) Then LoadSolution()
    End Sub

    Private Sub LoadSolution()
        If Not String.IsNullOrEmpty(_solutionPath) Then
            _vm.Solution = Solution.LoadSolution(_solutionPath)
        End If
        Me.TopFlyout.IsOpen = True
    End Sub

    Private Sub ButtonSaveFile_Click(sender As Object, e As RoutedEventArgs)
        If _vm IsNot Nothing AndAlso _vm.Solution IsNot Nothing Then
            If String.IsNullOrEmpty(_solutionPath) Then
                Using sd As New SystemDialog
                    _solutionPath = sd.SaveFile
                End Using
            End If
            If Not String.IsNullOrEmpty(_solutionPath) Then
                Solution.SaveSolution(_solutionPath, _vm.Solution)
                Application.Message($"{_solutionPath} saved.")
            End If
        End If
    End Sub

    Public Sub LoadReconSource(side As ReconSource.SideName)
        Dim reconSource As ReconSource
        If side.Equals(ReconSource.SideName.Left) Then
            reconSource = _vm.Reconciliation.LeftReconSource
        Else
            reconSource = _vm.Reconciliation.RightReconSource
        End If

        Dim r As MessageDialogResult = Nothing
        If reconSource.IsLoaded And reconSource.ReconDataSource.IsSlowLoading Then
            Application.Current.Dispatcher.BeginInvoke(Async Function()
                                                           r = Await ShowMessageAsync("Rekonator", "Do you want to reload data source?", MessageDialogStyle.AffirmativeAndNegative)
                                                       End Function)
            If r = MessageDialogResult.Negative Then Exit Sub
        End If

        Select Case reconSource.ReconDataSource.DataSourceName
            Case "Excel"
                Using excel As New GetExcel
                    reconSource.IsLoaded = excel.Load(reconSource)
                End Using
            Case "SQL"
                Using sql As New GetSQL
                    reconSource.IsLoaded = sql.Load(reconSource, _vm.Reconciliation.FromDate, _vm.Reconciliation.ToDate)
                End Using
            Case "QuickBooks"
                If reconSource.IsLoaded = True And reconSource.ReconDataSource.IsSlowLoading Then
                    If reconSource.ReconDataSource.IsSlowLoading Then
                    End If
                End If
                Using qbd As New GetQBD
                    reconSource.IsLoaded = qbd.LoadReport(reconSource, _vm.Reconciliation.FromDate, _vm.Reconciliation.ToDate)
                End Using
        End Select
    End Sub


    Private Sub btnRight_Click(sender As Object, e As RoutedEventArgs)
        _vm.MessageLog.Clear()
        '_reconciliation.RightReconSource.IsLoaded = False
        'btnMatch_Click(sender, e)
    End Sub
    Private Sub ButtonDiffer_Click(sender As Object, e As RoutedEventArgs)

    End Sub
    Private Sub ButtonMatch_Click(sender As Object, e As RoutedEventArgs)
        _vm.MessageLog.Clear() 'Test Agg QB P/L Detail
        Task.Factory.StartNew(Sub() Rekonate())

        'Task.Factory.StartNew(Sub() Configure("QB P/L Detail")).
        '                          ContinueWith(Sub() LoadReconSources()).
        '                          ContinueWith(Sub() Test())
    End Sub

    Private Sub DataGridRow_MouseRightButtonDown(sender As Object, e As MouseButtonEventArgs)
        Dim dgr As DataGridRow = TryCast(sender, DataGridRow)
        If dgr IsNot Nothing Then
            Dim dr As DataRow = TryCast(dgr.Item.Row, DataRow)
            If dr IsNot Nothing Then
                Using sql As New SQL

                    _left = sql.GetDataTable($"Select * From [dbo].[{_vm.Reconciliation.LeftReconSource.ReconTable}] Where [Txn ID] = '{dr.ItemArray(3)}' AND [GL ACCOUNT] = '{dr.ItemArray(4)}'")
                    _vm.LeftSet = _left.AsDataView
                    _right = sql.GetDataTable($"Select * From [dbo].[{_vm.Reconciliation.RightReconSource.ReconTable}] Where [TxnID] = '{dr.ItemArray(5)}' AND [Account] = '{dr.ItemArray(6)}'")
                    _vm.RightSet = _right.AsDataView
                End Using

            End If
        End If
    End Sub

    Private Sub DataGridRow_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs)
        Dim client As New Client
        client.Show()
        Me.Close()
    End Sub

    Friend Sub ChangeReconciliation(rc As Reconciliation)
        If rc IsNot Nothing Then
            _vm.Reconciliation = rc
            'Task.Factory.StartNew(Sub() Configure(solutionName)).
            '                  ContinueWith(Sub() LoadReconSources()).
            '                  ContinueWith(Sub() Test())
        End If
    End Sub
#End Region

    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()
        DataContext = _vm
        Application.MessageFunc = AddressOf AddMessage
    End Sub

    Private Sub Rekonate()
        _leftDetails.Reset()
        _rightDetails.Reset()
        _differ.Reset()
        _vm.DifferSet = _differ.AsDataView
        _match.Reset()
        _vm.MatchSet = _match.AsDataView

        'If _left.Rows.Count = 0 Or _right.Rows.Count = 0 Then Exit Sub

        Using sql As New SQL
            sql.DropTables({"Left", "Right", "Match", "Differ"})
        End Using
        Using sql As New SQL

            _match = sql.GetDataTable(Reconciliation.GetMatchSelect(_vm.Reconciliation))
            If _match.AsDataView Is Nothing Then Exit Sub
            _vm.MatchSet = _match.AsDataView
            _differ = sql.GetDataTable(Reconciliation.GetDifferSelect(_vm.Reconciliation))
            If _differ.AsDataView Is Nothing Then Exit Sub
            _vm.DifferSet = _match.AsDataView
            _left = sql.GetDataTable(Reconciliation.GetLeftSelect(_vm.Reconciliation))
            If _left.AsDataView Is Nothing Then Exit Sub
            _vm.LeftSet = _left.AsDataView
            _right = sql.GetDataTable(Reconciliation.GetRightSelect(_vm.Reconciliation))
            _vm.RightSet = _right.AsDataView
        End Using

        'If _reconciliation.LeftReconSource.Aggregations IsNot Nothing Then DoAggregation(_reconciliation.LeftReconSource)
        'If _reconciliation.RightReconSource.Aggregations IsNot Nothing Then DoAggregation(_reconciliation.RightReconSource)
    End Sub

    'Private Sub DoAggregation(reconSource As ReconSource)

    '    _rightDetails = _right.Copy
    '    _right.Reset()
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

    'End Sub


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



    Private Sub Configure(testName As String)
        Dim completenessComparisions As New List(Of Comparision)
        Dim matchingComparisions As New List(Of Comparision)
        Dim aggOps As List(Of AggregateOperation)
        Dim aggregates As List(Of Aggregate)
        Dim leftRS As ReconSource = Nothing
        Dim rightRS As ReconSource = Nothing
        Reconciliation.Clear()

        Select Case testName
            Case "DMR1"
                'Left
                aggOps = New List(Of AggregateOperation)
                aggOps.Add(New AggregateOperation With {.SourceColumn = "Amount", .AggregateColumn = "Total", .Operation = AggregateOperation.AggregateFunction.Sum})
                aggregates = New List(Of Aggregate)
                aggregates.Add(New Aggregate With {.GroupByColumns = {"GL Account"}, .AggregateOperations = aggOps})

                Dim sqlDS As DataSource = DataSource.GetDataSource("SQL")
                Dim sqlParams As New List(Of Parameter)
                sqlParams.AddParameter("connectionstring", "Data Source=dbvipmaster;Initial Catalog=Prod-Lahydrojet;User ID=linxlogic;Password=6a3r3a0$")
                sqlParams.AddParameter("schema", "lahydrojet1")
                'sqlParams.AddParameter("commandtext", "SELECT * FROM foo")
                sqlParams.AddParameter("commandpath", "C:\Users\Peter Grillo\Documents\SQL Server Management Studio\lah_sql_transdetail.sql")
                sqlParams.AddParameter("create", "CREATE TABLE #temptable ( [HV_DCR_DemandCallsRun] bigint, [HV_DemandRevenue] decimal(12,2), [HV_DCR_NoClubVisits] bigint, [HV_DCR_ClubRevenue] bigint, [HV_DCR_NoOfMaintenanceVisits] bigint, [HV_DCR_MaintenanceRevenue] decimal(12,2), [HV_ZeroRevenueCalls] bigint, [HV_DiagnosticFeeOnlyCalls] bigint, [PB_DCR_DemandCallsRun] bigint, [PB_DemandRevenue] decimal(12,2), [PB_DCR_NoClubVisits] bigint, [PB_DCR_ClubRevenue] bigint, [PB_DCR_NoOfMaintenanceVisits] bigint, [PB_DCR_MaintenanceRevenue] decimal(12,2), [PB_ZeroRevenueCalls] bigint, [PB_DiagnosticFeeOnlyCalls] bigint, [EL_DCR_DemandCallsRun] bigint, [EL_DemandRevenue] decimal(12,2), [EL_DCR_NoClubVisits] bigint, [EL_DCR_ClubRevenue] bigint, [EL_DCR_NoOfMaintenanceVisits] bigint, [EL_DCR_MaintenanceRevenue] decimal(12,2), [EL_ZeroRevenueCalls] bigint, [EL_DiagnosticFeeOnlyCalls] bigint )") 'Right click in SSMS result, change table name
                leftRS = New ReconSource With
                    {.ReconDataSource = sqlDS,
                    .ReconTable = "lah_sql_transdetail",
                    .IsLoaded = False,
                    .Parameters = sqlParams,
                    .Aggregations = aggregates}

                'Right
                aggOps = New List(Of AggregateOperation)
                aggOps.Add(New AggregateOperation With {.SourceColumn = "Amount", .AggregateColumn = "Total", .Operation = AggregateOperation.AggregateFunction.Sum})
                aggregates = New List(Of Aggregate)
                aggregates.Add(New Aggregate With {.GroupByColumns = {"Account"}, .AggregateOperations = aggOps})

                Dim qbDS As DataSource = DataSource.GetDataSource("SQL")
                sqlParams = New List(Of Parameter)
                sqlParams.AddParameter("connectionstring", "Data Source=dbvipmaster;Initial Catalog=Prod-Lahydrojet;User ID=linxlogic;Password=6a3r3a0$")
                sqlParams.AddParameter("schema", "lahydrojet1")
                'sqlParams.AddParameter("commandtext", "SELECT * FROM foo")
                sqlParams.AddParameter("commandpath", "C:\Users\Peter Grillo\Documents\SQL Server Management Studio\lah_sql_transdetail.sql")
                sqlParams.AddParameter("create", "CREATE TABLE #temptable ( [HV_DCR_DemandCallsRun] bigint, [HV_DemandRevenue] decimal(12,2), [HV_DCR_NoClubVisits] bigint, [HV_DCR_ClubRevenue] bigint, [HV_DCR_NoOfMaintenanceVisits] bigint, [HV_DCR_MaintenanceRevenue] decimal(12,2), [HV_ZeroRevenueCalls] bigint, [HV_DiagnosticFeeOnlyCalls] bigint, [PB_DCR_DemandCallsRun] bigint, [PB_DemandRevenue] decimal(12,2), [PB_DCR_NoClubVisits] bigint, [PB_DCR_ClubRevenue] bigint, [PB_DCR_NoOfMaintenanceVisits] bigint, [PB_DCR_MaintenanceRevenue] decimal(12,2), [PB_ZeroRevenueCalls] bigint, [PB_DiagnosticFeeOnlyCalls] bigint, [EL_DCR_DemandCallsRun] bigint, [EL_DemandRevenue] decimal(12,2), [EL_DCR_NoClubVisits] bigint, [EL_DCR_ClubRevenue] bigint, [EL_DCR_NoOfMaintenanceVisits] bigint, [EL_DCR_MaintenanceRevenue] decimal(12,2), [EL_ZeroRevenueCalls] bigint, [EL_DiagnosticFeeOnlyCalls] bigint )") 'Right click in SSMS result, change table name
                rightRS = New ReconSource With
                    {.ReconDataSource = qbDS,
                    .ReconTable = "DMR1new",
                    .IsLoaded = True,
                    .Parameters = sqlParams,
                    .Aggregations = aggregates}

                completenessComparisions.Add(New Comparision With {.LeftColumn = "GL Account", .RightColumn = "Account", .ComparisionTest = ComparisionType.TextCaseEquals}) '.RightFunction = "SUBSTRING({RightColumn}, 11,  LEN({RightColumn}) -10)"
                matchingComparisions.Add(New Comparision With {.LeftColumn = "Total", .RightColumn = "Total", .Percision = 2, .ComparisionTest = ComparisionType.NumberEquals})
                Reconciliation.Add("PL Summary", leftRS, rightRS, completenessComparisions, matchingComparisions, #6/1/2018#, #6/30/2018#)
            Case "P/L Summary"
                'Left
                aggOps = New List(Of AggregateOperation)
                aggOps.Add(New AggregateOperation With {.SourceColumn = "Amount", .AggregateColumn = "Total", .Operation = AggregateOperation.AggregateFunction.Sum})
                aggregates = New List(Of Aggregate)
                aggregates.Add(New Aggregate With {.GroupByColumns = {"GL Account"}, .AggregateOperations = aggOps})

                Dim sqlDS As DataSource = DataSource.GetDataSource("SQL")
                Dim sqlParams As New List(Of Parameter)
                sqlParams.AddParameter("connectionstring", "Data Source=dbvipmaster;Initial Catalog=Prod-Lahydrojet;User ID=linxlogic;Password=6a3r3a0$")
                sqlParams.AddParameter("schema", "lahydrojet1")
                'sqlParams.AddParameter("commandtext", "SELECT * FROM foo")
                sqlParams.AddParameter("commandpath", "C:\Users\Peter Grillo\Documents\SQL Server Management Studio\lah_sql_accountdetail.sql")
                sqlParams.AddParameter("create", "CREATE TABLE lah_sql_accountdetail ( [GL Type] nvarchar(255), [GL Number] nvarchar(4000), [GL Account] nvarchar(255), [Reference] nvarchar(50), [Date] datetime2(7), [Amount] decimal(9,2), [TXN ID] nvarchar(4000), [Id] bigint, [Business Unit] nvarchar(255) )") 'Right click in SSMS result, change table name
                leftRS = New ReconSource With
                    {.ReconDataSource = sqlDS,
                    .ReconTable = "lah_sql_accountdetail",
                    .IsLoaded = True,
                    .Parameters = sqlParams,
                    .Aggregations = aggregates}

                'Right
                aggOps = New List(Of AggregateOperation)
                aggOps.Add(New AggregateOperation With {.SourceColumn = "Amount", .AggregateColumn = "Total", .Operation = AggregateOperation.AggregateFunction.Sum})
                aggregates = New List(Of Aggregate)
                aggregates.Add(New Aggregate With {.GroupByColumns = {"Account"}, .AggregateOperations = aggOps})

                Dim qbDS As DataSource = DataSource.GetDataSource("QuickBooks")
                sqlParams = New List(Of Parameter)
                sqlParams.AddParameter("request", "AppendGeneralDetailReportQueryRq")
                sqlParams.AddParameter("detailreporttype", "gdrtProfitAndLossDetail")
                rightRS = New ReconSource With
                    {.ReconDataSource = qbDS,
                    .ReconTable = "lah_qb_pldetail",
                    .IsLoaded = True,
                    .Parameters = sqlParams,
                    .Aggregations = aggregates}

                completenessComparisions.Add(New Comparision With {.LeftColumn = "GL Account", .RightColumn = "Account", .ComparisionTest = ComparisionType.TextCaseEquals, .RightFunction = "SUBSTRING({RightColumn}, 11,  LEN({RightColumn}) -10)"})
                matchingComparisions.Add(New Comparision With {.LeftColumn = "Total", .RightColumn = "Total", .Percision = 2, .ComparisionTest = ComparisionType.NumberEquals})
                Reconciliation.Add("PL Summary", leftRS, rightRS, completenessComparisions, matchingComparisions, #6/1/2018#, #6/30/2018#)

            Case "QB P/L Detail"
                'Left
                aggOps = New List(Of AggregateOperation)
                aggOps.Add(New AggregateOperation With {.SourceColumn = "Amount", .AggregateColumn = "Total", .Operation = AggregateOperation.AggregateFunction.Sum})
                aggregates = New List(Of Aggregate)
                aggregates.Add(New Aggregate With {.GroupByColumns = {"TXN ID", "GL Account"}, .AggregateOperations = aggOps})

                Dim sqlDS As DataSource = DataSource.GetDataSource("SQL")
                Dim sqlParams As New List(Of Parameter)
                sqlParams.AddParameter("connectionstring", "Data Source=dbvipmaster;Initial Catalog=Prod-Lahydrojet;User ID=linxlogic;Password=6a3r3a0$")
                sqlParams.AddParameter("schema", "lahydrojet1")
                'sqlParams.AddParameter("commandtext", "SELECT * FROM foo")
                sqlParams.AddParameter("commandpath", "C:\Users\Peter Grillo\Documents\SQL Server Management Studio\lah_sql_accountdetail.sql")
                sqlParams.AddParameter("create", "CREATE TABLE lah_sql_accountdetail ( [GL Type] nvarchar(255), [GL Number] nvarchar(4000), [GL Account] nvarchar(255), [Reference] nvarchar(50), [Date] datetime2(7), [Amount] decimal(9,2), [TXN ID] nvarchar(4000), [Id] bigint, [Business Unit] nvarchar(255) )") 'Right click in SSMS result, change table name
                leftRS = New ReconSource With
                    {.ReconDataSource = sqlDS,
                    .ReconTable = "lah_sql_accountdetail",
                    .IsLoaded = False,
                    .Parameters = sqlParams,
                    .Aggregations = aggregates}

                'Right
                aggOps = New List(Of AggregateOperation)
                aggOps.Add(New AggregateOperation With {.SourceColumn = "Amount", .AggregateColumn = "Total", .Operation = AggregateOperation.AggregateFunction.Sum})
                aggregates = New List(Of Aggregate)
                aggregates.Add(New Aggregate With {.GroupByColumns = {"TxnID", "Account"}, .AggregateOperations = aggOps})

                Dim qbDS As DataSource = DataSource.GetDataSource("QuickBooks")
                sqlParams = New List(Of Parameter)
                sqlParams.AddParameter("request", "AppendGeneralDetailReportQueryRq")
                sqlParams.AddParameter("detailreporttype", "gdrtProfitAndLossDetail")
                rightRS = New ReconSource With
                    {.ReconDataSource = qbDS,
                    .ReconTable = "lah_qb_pldetail",
                    .IsLoaded = True,
                    .Parameters = sqlParams,
                    .Aggregations = aggregates}

                completenessComparisions.Add(New Comparision With {.LeftColumn = "TXN ID", .RightColumn = "TxnID", .ComparisionTest = ComparisionType.TextCaseEquals})
                completenessComparisions.Add(New Comparision With {.LeftColumn = "GL Account", .RightColumn = "Account", .ComparisionTest = ComparisionType.TextCaseEquals})
                matchingComparisions.Add(New Comparision With {.LeftColumn = "Total", .RightColumn = "Total", .Percision = 2, .ComparisionTest = ComparisionType.NumberEquals})
                Reconciliation.Add("Test Set", leftRS, rightRS, completenessComparisions, matchingComparisions, #6/1/2018#, #6/30/2018#)

            Case "Account Detail"
                'Left
                Dim sqlDS As DataSource = DataSource.GetDataSource("SQL")
                Dim sqlParams As New List(Of Parameter)
                sqlParams.AddParameter("connectionstring", "Data Source=dbvipmaster;Initial Catalog=Prod-Lahydrojet;User ID=linxlogic;Password=6a3r3a0$")
                sqlParams.AddParameter("schema", "lahydrojet1")
                'sqlParams.AddParameter("commandtext", "SELECT * FROM foo")
                sqlParams.AddParameter("commandpath", "C:\Users\Peter Grillo\Documents\SQL Server Management Studio\lah_sql_accountdetail.sql")
                sqlParams.AddParameter("create", "CREATE TABLE lah_sql_accountdetail ( [GL Name] nvarchar(255), [GL Number] nvarchar(4000), [Name] nvarchar(255), [Number] nvarchar(50), [InvoicedOn] datetime2(7), [Total] decimal(9,2), [ExportId] nvarchar(4000), [Id] bigint, [Active] bit )") 'Right click in SSMS result
                leftRS = New ReconSource With
                    {.ReconDataSource = sqlDS,
                    .ReconTable = "lah_sql_accountdetail",
                    .IsLoaded = True,
                    .Parameters = sqlParams}

                'Right

                Dim excelDS As DataSource = DataSource.GetDataSource("Excel")
                Dim excelParams As New List(Of Parameter)
                excelParams.AddParameter("FilePath", "C:\Users\Peter Grillo\Downloads\1_08776f78-eaac-468c-89d4-aeedf5b65673_AccountingDetail.xlsx")
                excelParams.AddParameter("Worksheet", "Combined")
                rightRS = New ReconSource With
                    {.ReconDataSource = excelDS,
                    .ReconTable = "lah_accountdetail",
                    .IsLoaded = True,
                    .Parameters = excelParams}

                completenessComparisions.Add(New Comparision With {.LeftColumn = "ExportID", .RightColumn = "Item ID", .ComparisionTest = ComparisionType.TextCaseEquals})
                matchingComparisions.Add(New Comparision With {.LeftColumn = "ST Item Name", .RightColumn = "Item Name", .ComparisionTest = ComparisionType.TextEquals})
                matchingComparisions.Add(New Comparision With {.LeftColumn = "ST Item Qty", .RightColumn = "Item Qty", .Percision = 4, .ComparisionTest = ComparisionType.NumberEquals})
                matchingComparisions.Add(New Comparision With {.LeftColumn = "ST Item Price", .RightColumn = "Item Price", .Percision = 2, .ComparisionTest = ComparisionType.NumberEquals})
                matchingComparisions.Add(New Comparision With {.LeftColumn = "ST Item Reorder", .RightColumn = "Item Reorder", .ComparisionTest = ComparisionType.NumberEquals})
                matchingComparisions.Add(New Comparision With {.LeftColumn = "ST Service Date", .RightColumn = "Service Date", .ComparisionTest = ComparisionType.DateEquals})
                Reconciliation.Add("Test Set", leftRS, rightRS, completenessComparisions, matchingComparisions, #5/1/2018#, #5/31/2018#)
            Case "Test 1"


                'Test 1
                Dim excelDS As DataSource = DataSource.GetDataSource("Excel")
                Dim excelParams As New List(Of Parameter)
                excelParams.AddParameter("FilePath", "C:\Users\Peter Grillo\source\repos\Test.xlsx")
                excelParams.AddParameter("Worksheet", "ST Items")
                leftRS = New ReconSource With
                    {.ReconDataSource = excelDS,
                    .ReconTable = "aset",
                    .IsLoaded = False,
                    .Parameters = excelParams}

                excelParams = New List(Of Parameter)
                excelParams.AddParameter("FilePath", "C:\Users\Peter Grillo\source\repos\Test.xlsx")
                excelParams.AddParameter("Worksheet", "Item")
                rightRS = New ReconSource With
                    {.ReconDataSource = excelDS,
                    .ReconTable = "bset",
                    .IsLoaded = False,
                    .Parameters = excelParams}

                completenessComparisions.Add(New Comparision With {.LeftColumn = "ST Export ID", .RightColumn = "Item ID", .ComparisionTest = ComparisionType.TextCaseEquals})
                matchingComparisions.Add(New Comparision With {.LeftColumn = "ST Item Name", .RightColumn = "Item Name", .ComparisionTest = ComparisionType.TextEquals})
                matchingComparisions.Add(New Comparision With {.LeftColumn = "ST Item Qty", .RightColumn = "Item Qty", .Percision = 4, .ComparisionTest = ComparisionType.NumberEquals})
                matchingComparisions.Add(New Comparision With {.LeftColumn = "ST Item Price", .RightColumn = "Item Price", .Percision = 2, .ComparisionTest = ComparisionType.NumberEquals})
                matchingComparisions.Add(New Comparision With {.LeftColumn = "ST Item Reorder", .RightColumn = "Item Reorder", .ComparisionTest = ComparisionType.NumberEquals})
                matchingComparisions.Add(New Comparision With {.LeftColumn = "ST Service Date", .RightColumn = "Service Date", .ComparisionTest = ComparisionType.DateEquals})
                Reconciliation.Add("Test Set", leftRS, rightRS, completenessComparisions, matchingComparisions)
            Case "Huber"
                Dim excelDS As DataSource = DataSource.GetDataSource("Excel")
                Dim excelParams As New List(Of Parameter)
                excelParams.AddParameter("FilePath", "C:\Users\Peter Grillo\Downloads\March Activity.xls")
                excelParams.AddParameter("Worksheet", "ST Inv")
                leftRS = New ReconSource With
                    {.ReconDataSource = excelDS,
                    .ReconTable = "stinvoice",
                    .IsLoaded = True,
                    .Parameters = excelParams,
                    .WhereClause = "NOT (ISNULL(x!.[ExportId], '')='' OR x!.[ExportId]='0')"}

                excelParams = New List(Of Parameter)
                excelParams.AddParameter("FilePath", "C:\Users\Peter Grillo\Downloads\March Activity.xls")
                excelParams.AddParameter("Worksheet", "Mar Inv")
                rightRS = New ReconSource With
                    {.ReconDataSource = excelDS,
                    .ReconTable = "qbinvoice",
                    .IsLoaded = True,
                    .Parameters = excelParams,
                    .WhereClause = "NOT (ISNULL(x!.[SubTotal], '')='' OR x!.[SubTotal]='0')"}

                completenessComparisions.Add(New Comparision With {.LeftColumn = "QB Export ID", .RightColumn = "TxnId", .ComparisionTest = ComparisionType.TextCaseEquals})
                matchingComparisions.Add(New Comparision With {.LeftColumn = "Total", .RightColumn = "Subtotal", .Percision = 2, .ComparisionTest = ComparisionType.NumberEquals})
                Reconciliation.Add("March Invoices", leftRS, rightRS, completenessComparisions, matchingComparisions)

            Case "Test Agg"
                'Left
                Dim excelDS As DataSource = DataSource.GetDataSource("Excel")
                Dim excelParams As List(Of Parameter)
                excelParams.AddParameter("FilePath", "C:\Users\Peter Grillo\source\repos\Test.xlsx")
                excelParams.AddParameter("Worksheet", "Agg Balance")
                Dim columns As New List(Of Column)
                columns.AddColumns({"Customer ID", "Integer", "Balance", "Currency", "Count", "Integer", "Avg", "Single", "Test", "String"})
                leftRS = New ReconSource With
                    {.ReconDataSource = excelDS,
                    .ReconTable = "aggbalance",
                    .IsLoaded = False,
                    .Parameters = excelParams,
                    .WhereClause = "ISNULL(x!.[Balance], 0) <> 0",
                    .Columns = columns
                }

                'Right
                aggOps = New List(Of AggregateOperation)
                aggOps.Add(New AggregateOperation With {.SourceColumn = "Entry", .AggregateColumn = "EntryTotal", .Operation = AggregateOperation.AggregateFunction.Sum})
                aggOps.Add(New AggregateOperation With {.SourceColumn = "Entry", .AggregateColumn = "EntryAvg", .Operation = AggregateOperation.AggregateFunction.Avg})
                aggOps.Add(New AggregateOperation With {.SourceColumn = "Entry", .AggregateColumn = "EntryCount", .Operation = AggregateOperation.AggregateFunction.Count})

                aggregates = New List(Of Aggregate)
                aggregates.Add(New Aggregate With {.GroupByColumns = {"Customer ID"}, .AggregateOperations = aggOps})

                excelParams = New List(Of Parameter)
                excelParams.AddParameter("FilePath", "C:\Users\Peter Grillo\source\repos\Test.xlsx")
                excelParams.AddParameter("Worksheet", "Agg Entries")
                columns = New List(Of Column)
                columns.AddColumns({"Customer ID", "Integer", "Entry", "Currency", "Test", "String"})
                rightRS = New ReconSource With
                    {.ReconDataSource = excelDS,
                    .ReconTable = "aggentries",
                    .IsLoaded = False,
                    .Parameters = excelParams,
                    .Aggregations = aggregates,
                    .WhereClause = "ISNULL(x!.[Entry], 0) <> 0",
                    .Columns = columns
                }

                completenessComparisions.Add(New Comparision With {.LeftColumn = "Customer ID", .RightColumn = "Customer ID", .ComparisionTest = ComparisionType.NumberEquals})
                matchingComparisions.Add(New Comparision With {.LeftColumn = "Balance", .RightColumn = "EntryTotal", .Percision = 2, .ComparisionTest = ComparisionType.NumberEquals})
                matchingComparisions.Add(New Comparision With {.LeftColumn = "Count", .RightColumn = "EntryCount", .Percision = 0, .ComparisionTest = ComparisionType.NumberEquals})
                matchingComparisions.Add(New Comparision With {.LeftColumn = "Avg", .RightColumn = "EntryAvg", .Percision = 9, .ComparisionTest = ComparisionType.NumberEquals})
                Reconciliation.Add("March Invoices", leftRS, rightRS, completenessComparisions, matchingComparisions)

        End Select

        '_solution = New Solution With {.SolutionName = testName, .Reconciliations = Reconciliation.Reconciliations}
        '_reconciliation = _solution.Reconciliations(0)

    End Sub


    Public Sub AddMessage(messageText As String, Optional isError As Boolean = False)
        Try
            If Not BottomFlyout.IsOpen Then BottomFlyout.IsOpen = True
            _vm.MessageLog.Add(New MessageEntry With {.MessageText = messageText, .IsError = isError})
            With UserControlMessageLog.ListBoxMessageLog
                .SelectedIndex = .Items.Count - 1
                .ScrollIntoView(.SelectedItem)
            End With
        Catch

        End Try
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



    Private Sub MainWindow_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        'UserControlMessageLog.DataContext = _vm 'Model for App Settings
        'Top UserControlSolution bound to vm.Reconciliation
        Using m As New Mock
            _vm.DataSources = m.MockLoadDataSources()
            '_vm.Solution = Task.Run(Function() m.MockLoadSolutionAsync(1)).GetAwaiter().GetResult() 'Model for Solution
        End Using

        _solutionPath = "C:\Users\Peter Grillo\Documents\dmr.rek"
        LoadSolution()
    End Sub

    Private Sub dgDiffer_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles dgDiffer.SelectionChanged

    End Sub
End Class
