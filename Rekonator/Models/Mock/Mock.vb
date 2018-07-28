Imports Rekonator

Public Class SettingsModel
    Property Datasources As List(Of DataSource)
    Property CompareMethods As List(Of CompareMethod)
End Class
Public Class Mock
    Implements IDisposable

    Public Function MockLoadSettings() As SettingsModel
        Dim sm As New SettingsModel
        CompareMethod.Add("Integer Equals", AddressOf ValueComparer.CompareIntegerValues)
        CompareMethod.Add("Single Equals", AddressOf ValueComparer.CompareSingleValues)
        CompareMethod.Add("String Equals", AddressOf ValueComparer.CompareStringValues)
        CompareMethod.Add("Date Equals", AddressOf ValueComparer.CompareDateValues)
        sm.CompareMethods = CompareMethod.CompareMethods

        Dim ds As New List(Of DataSource)
        ds.Add(New DataSource With {.DataSourceName = "Excel"})
        ds.Add(New DataSource With {.DataSourceName = "Intact"})
        ds.Add(New DataSource With {.DataSourceName = "QuickBooks"})
        ds.Add(New DataSource With {.DataSourceName = "ServiceTitan"})
        ds.Add(New DataSource With {.DataSourceName = "SQL"})
        sm.Datasources = ds

        Return sm
    End Function

    Public Async Function MockLoadSolutionAsync(mainWindow As MainWindow) As Task(Of Solution)
        Try

            Dim excelDS As DataSource = mainWindow.DataSources.Where(Function(w) w.DataSourceName = "Excel").FirstOrDefault
            Dim excelParam As New Dictionary(Of String, String)
            excelParam.Add("FilePath", "C:\Users\Peter Grillo\source\repos\Test 1.xlsx")
            excelParam.Add("Worksheet", "ST Items")
            Dim leftRS As New ReconSource With {.ReconDataSource = excelDS, .Parameters = excelParam}
            excelParam("Worksheet") = "Items"
            Dim rightRS As New ReconSource With {.ReconDataSource = excelDS, .Parameters = Nothing}

            Dim completenessComparisions As New List(Of Comparision)
            completenessComparisions.Add(New Comparision With {.LeftColumn = "ST Export ID", .RightColumn = "Item ID", .ComparisionOption = 1, .ComparisionMethod = CompareMethod.GetMethod("String Equals")})

            Dim matchingComparisions As New List(Of Comparision)
            matchingComparisions.Add(New Comparision With {.LeftColumn = "ST Export ID", .RightColumn = "Item ID", .ComparisionOption = 1, .ComparisionMethod = CompareMethod.GetMethod("String Equals")})
            matchingComparisions.Add(New Comparision With {.LeftColumn = "ST Item Name", .RightColumn = "Item Name", .ComparisionOption = 0, .ComparisionMethod = CompareMethod.GetMethod("String Equals")})
            matchingComparisions.Add(New Comparision With {.LeftColumn = "ST Item Qty", .RightColumn = "Item Qty", .ComparisionOption = 4, .ComparisionMethod = CompareMethod.GetMethod("Single Equals")})
            matchingComparisions.Add(New Comparision With {.LeftColumn = "ST Item Price", .RightColumn = "Item Price", .ComparisionOption = 2, .ComparisionMethod = CompareMethod.GetMethod("Single Equals")})
            matchingComparisions.Add(New Comparision With {.LeftColumn = "ST Item Reorder", .RightColumn = "Item Reorder", .ComparisionOption = 1, .ComparisionMethod = CompareMethod.GetMethod("Integer Equals")})
            matchingComparisions.Add(New Comparision With {.LeftColumn = "ST Service Date", .RightColumn = "Service Date", .ComparisionOption = 0, .ComparisionMethod = CompareMethod.GetMethod("Date Equals")})

            Reconciliation.Add("Test 1", leftRS, rightRS, completenessComparisions, matchingComparisions)

            Reconciliation.Add("Test Agg", leftRS, rightRS, completenessComparisions, matchingComparisions)

            Dim currentSolution As New Solution With {.SolutionName = "Test", .Reconciliations = Reconciliation.Reconciliations}

            'test Save
            'Dim fileName As String = "C:\Users\Peter Grillo\Downloads\HighTower.rek"
            'DataSource.Add("QuickBooks", "P/L Report", "")
            'DataSource.Add("ServiceTitan", "P/L Report", "")
            'Reconciliation.Add("P/L", DataSource.DataSources(0), DataSource.DataSources(1), _completenessComparisions, _matchingComparisions)
            'DataSource.Add("QuickBooks", "AR Report", "")
            'DataSource.Add("ServiceTitan", "AR Report", "")
            'Reconciliation.Add("AR", DataSource.DataSources(2), DataSource.DataSources(3), _completenessComparisions, _matchingComparisions)
            'Dim currentSolution As New Solution With {.SolutionName = "HighTower", .Reconciliations = Reconciliation.Reconciliations}
            'Solution.SaveSolution(fileName, currentSolution)
            'Dim aSolution As Solution = Solution.LoadSolution(fileName)
            'If currentSolution.SolutionName = aSolution.SolutionName Then
            'Beep()
            'End If


            '_aggregates.Add(New Aggregate With {.DataSourceName = "TestRight", .GroupBy = {"Invoice ID", "Invoice Date"}), .AggregateTarget = .Columns=z
            'LeftSet = _left.AsDataView
            'RightSet = _right.AsDataView
            Await LoadAsync()

        Catch ex As Exception

        End Try
    End Function

    Private Function LoadAsync() As Task
        Throw New NotImplementedException()
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
