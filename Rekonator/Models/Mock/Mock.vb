Imports Rekonator

Public Class Mock
    Implements IDisposable

    Public Function MockLoadSettings(mainWindow As MainWindow) As AppViewModel
        Dim appVM As New AppViewModel
        'CompareMethod.Add("Integer Equals", AddressOf ValueComparer.CompareIntegerValues)
        'CompareMethod.Add("Single Equals", AddressOf ValueComparer.CompareSingleValues)
        'CompareMethod.Add("String Equals", AddressOf ValueComparer.CompareStringValues)
        'CompareMethod.Add("Date Equals", AddressOf ValueComparer.CompareDateValues)
        'sm.CompareMethods = CompareMethod.CompareMethods

        Dim ds As New List(Of DataSource)
        ds.Add(New DataSource With {.DataSourceName = "Excel"})
        ds.Add(New DataSource With {.DataSourceName = "Intact"})
        ds.Add(New DataSource With {.DataSourceName = "QuickBooks"})
        ds.Add(New DataSource With {.DataSourceName = "ServiceTitan"})
        ds.Add(New DataSource With {.DataSourceName = "SQL"})
        appVM.DataSources = ds

        appVM.MainWindow = mainWindow
        Return appVM
    End Function

    Public Async Function MockLoadSolutionAsync(solutionNumber As Integer) As Task(Of Solution)
        'Public Function MockLoadSolution() As Solution

        Try
            Dim aggOps As List(Of AggregateOperation)
            Dim aggregates As List(Of Aggregate)
            Dim leftRS As ReconSource
            Dim rightRS As ReconSource
            Dim completenessComparisions As List(Of Comparision)
            Dim matchingComparisions As List(Of Comparision)

            Select Case solutionNumber
                Case 1 ' "Invoice Completeness"
                    'Left
                    aggOps = New List(Of AggregateOperation)
                    aggOps.Add(New AggregateOperation With {.SourceColumn = "Amount", .AggregateColumn = "Total", .Operation = AggregateOperation.AggregateFunction.Sum})
                    aggregates = New List(Of Aggregate)
                    aggregates.Add(New Aggregate With {.GroupByColumns = {"TXN ID"}, .AggregateOperations = aggOps})

                    Dim sqlDS As DataSource = DataSource.GetDataSource("SQL")
                    Dim sqlParams As New List(Of Parameter)
                    sqlParams.AddParameter("connectionstring", "Data Source=dbvipmaster;Initial Catalog=Prod-Lahydrojet;User ID=linxlogic;Password=6a3r3a0$")
                    sqlParams.AddParameter("schema", "lahydrojet1")
                    'sqlParams.AddParameter("commandtext", "SELECT * FROM foo")
                    sqlParams.AddParameter("commandpath", "C:\Users\Peter Grillo\Documents\SQL Server Management Studio\lah_sql_paymenttransdetail.sql")
                    sqlParams.AddParameter("create", "CREATE TABLE lah_sql_trans ( [GL Type] nvarchar(255), [GL Number] nvarchar(4000), [GL Account] nvarchar(255), [Reference] nvarchar(50), [Date] datetime2(7), [Amount] decimal(9,2), [TXN ID] nvarchar(4000), [Id] bigint, [Business Unit] nvarchar(255) )") 'Right click in SSMS result, change table name
                    leftRS = New ReconSource With
                    {.ReconDataSource = sqlDS,
                    .ReconTable = "lah_sql_trans",
                    .IsLoaded = True,
                    .Parameters = sqlParams,
                    .Aggregations = aggregates}

                    'Right
                    aggOps = New List(Of AggregateOperation)
                    aggOps.Add(New AggregateOperation With {.SourceColumn = "Amount", .AggregateColumn = "Total", .Operation = AggregateOperation.AggregateFunction.Sum})
                    aggregates = New List(Of Aggregate)
                    aggregates.Add(New Aggregate With {.GroupByColumns = {"TxnID"}, .AggregateOperations = aggOps})

                    Dim qbDS As DataSource = DataSource.GetDataSource("QuickBooks")
                    sqlParams = New List(Of Parameter)
                    sqlParams.AddParameter("detailreporttype", "gdrtTxnListByDate")
                    rightRS = New ReconSource With
                    {.ReconDataSource = qbDS,
                    .ReconTable = "lah_qb_transdetail",
                    .IsLoaded = True,
                    .Parameters = sqlParams,
                    .Aggregations = aggregates}

                    completenessComparisions = New List(Of Comparision)
                    completenessComparisions.Add(New Comparision With {.LeftColumn = "TXN ID", .RightColumn = "TxnID", .ComparisionTest = ComparisionType.TextCaseEquals}) '.RightFunction = "SUBSTRING({RightColumn}, 11,  LEN({RightColumn}) -10)"
                    matchingComparisions = New List(Of Comparision)
                    matchingComparisions.Add(New Comparision With {.LeftColumn = "Total", .RightColumn = "Total", .Percision = 2, .ComparisionTest = ComparisionType.NumberEquals})
                    Reconciliation.Add("Trans Detail", leftRS, rightRS, completenessComparisions, matchingComparisions, #1/1/2018#, #6/30/2018#)
                    Return New Solution With {.SolutionName = "ST/QB Trans", .Reconciliations = Reconciliation.Reconciliations}
                Case 2

                    ' "QB P/L Detail"
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
                    .IsLoaded = True,
                    .Parameters = sqlParams,
                    .Aggregations = aggregates,
                    .InstantiatedSide = ReconSource.Side.Left}

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
                    .Aggregations = aggregates,
                    .InstantiatedSide = ReconSource.Side.Rigth}

                    completenessComparisions = New List(Of Comparision)
                    completenessComparisions.Add(New Comparision With {.LeftColumn = "TXN ID", .RightColumn = "TxnID", .ComparisionTest = ComparisionType.TextCaseEquals})
                    completenessComparisions.Add(New Comparision With {.LeftColumn = "GL Account", .RightColumn = "Account", .ComparisionTest = ComparisionType.TextCaseEquals})
                    matchingComparisions = New List(Of Comparision)
                    matchingComparisions.Add(New Comparision With {.LeftColumn = "Total", .RightColumn = "Total", .Percision = 2, .ComparisionTest = ComparisionType.NumberEquals})
                    Reconciliation.Add("Test Set", leftRS, rightRS, completenessComparisions, matchingComparisions, #1/1/2018#, #6/30/2018#)
                    Return New Solution With {.SolutionName = "QB Reports", .Reconciliations = Reconciliation.Reconciliations}
            End Select
        Catch ex As Exception
            Application.ErrorMessage("Error Mocking Solution")
        End Try
    End Function



#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
        ' TODO: uncomment the following line if Finalize() is overridden above.
        ' GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class

