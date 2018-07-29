Imports System.Text

<Serializable()>
Public Class Reconciliation
    Property ReconciliationName As String
    Property LeftReconSource As ReconSource
    Property RightReconSource As ReconSource
    Property CompletenessComparisions As List(Of Comparision)
    Property MatchingComparisions As List(Of Comparision)
    Property FromDate As DateTime = DateTime.MinValue
    Property ToDate As DateTime = DateTime.MinValue

    Private Shared _sb As New StringBuilder

    Public Shared Reconciliations As New List(Of Reconciliation)
    Public Shared Sub Add(reconciliationName As String,
                          leftDataSource As ReconSource,
                          rightDataSource As ReconSource,
                          completenessComparision As List(Of Comparision),
                          matchingComparision As List(Of Comparision),
                          Optional fromDate As DateTime = Nothing,
                          Optional toDate As DateTime = Nothing)
        Reconciliations.Add(New Reconciliation With {
                            .ReconciliationName = reconciliationName,
                            .LeftReconSource = leftDataSource,
                            .RightReconSource = rightDataSource,
                            .CompletenessComparisions = completenessComparision,
                            .MatchingComparisions = matchingComparision,
                            .FromDate = fromDate,
                            .ToDate = toDate}
                            )
    End Sub

    Public Shared Sub Clear()
        Reconciliations.Clear()
    End Sub
    Public Shared Function GetReconciliation(reconciliationName As String) As Reconciliation
        Return Reconciliations.Where(Function(w) w.ReconciliationName = reconciliationName).FirstOrDefault
    End Function

    Public Shared Function GetMatchSelect(recon As Reconciliation) As String
        Dim isFirst As Boolean = True
        _sb.Clear()
        _sb.AppendLine("IF(OBJECT_ID('Rekonator..[Match]') IS NOT NULL) DROP TABLE [Match];")
        _sb.AppendLine()

        Dim cteTable As String = String.Empty
        Dim isAggA As Boolean = (recon.LeftReconSource.Aggregations IsNot Nothing)
        Dim aTable As String = $"[dbo].[{recon.LeftReconSource.ReconTable}] a"
        If isAggA Then
            cteTable = $"cte_{recon.LeftReconSource.ReconTable}_grp"
            _sb.AppendLine(MakeGroupBy(recon.LeftReconSource, cteTable, "a"))
            aTable = $"{cteTable} a"
        End If
        Dim isAggB As Boolean = (recon.RightReconSource.Aggregations IsNot Nothing)
        Dim bTable As String = $"[dbo].[{recon.RightReconSource.ReconTable}] b"
        If isAggB Then
            cteTable = $"cte_{recon.RightReconSource.ReconTable}_grp"
            _sb.AppendLine(MakeGroupBy(recon.RightReconSource, cteTable, "b"))
            bTable = $"{cteTable} b"
        End If

        _sb.AppendLine("SELECT")
        For Each c As Comparision In recon.CompletenessComparisions.Union(recon.MatchingComparisions)
            If Not isFirst Then
                _sb.Append(",")
            End If
            _sb.AppendLine($"[{c.LeftColumn}={c.RightColumn}] = a.[{c.LeftColumn}]")
            isFirst = False
        Next
        If isAggA Then _sb.AppendLine(MakeGroupByColumns(recon.LeftReconSource.Aggregations(0).GroupByColumns, "a", True))
        If isAggB Then _sb.AppendLine(MakeGroupByColumns(recon.RightReconSource.Aggregations(0).GroupByColumns, "b", True))
        If Not isAggA Then _sb.AppendLine(",[idA] = a.[rekonid]")
        If Not isAggB Then _sb.AppendLine(", [idB] = b.[rekonid]")
        _sb.AppendLine("INTO [Match]")
        _sb.AppendLine($"FROM {aTable}, {bTable}")
        _sb.AppendLine("WHERE")

        isFirst = True
        For Each c As Comparision In recon.CompletenessComparisions.Union(recon.MatchingComparisions)
            If Not isFirst Then
                _sb.Append("And ")
            End If
            Select Case c.ComparisionTest
                Case ComparisionType.TextCaseEquals
                    _sb.AppendLine($"ISNULL(a.[{c.LeftColumn}],'') = ISNULL(b.[{c.RightColumn}],'')")
                Case ComparisionType.TextEquals
                    _sb.AppendLine($"LOWER(ISNULL(a.[{c.LeftColumn}],'')) = LOWER(ISNULL(b.[{c.RightColumn}],''))")
                Case ComparisionType.NumberEquals
                    _sb.AppendLine($"CONVERT(DECIMAL(14,{c.Percision}), ISNULL(a.[{c.LeftColumn}],0)) = CONVERT(DECIMAL(14,{c.Percision}), ISNULL(b.[{c.RightColumn}],0))")
                Case ComparisionType.DateEquals
                    _sb.AppendLine($"CONVERT(DATE, a.[{c.LeftColumn}]) = CONVERT(DATE, b.[{c.RightColumn}])")
                Case ComparisionType.DateTimeEquals
                    _sb.AppendLine($"CONVERT(DATETIME, a.[{c.LeftColumn}]) = CONVERT(DATETIME, b.[{c.RightColumn}])")
            End Select
            isFirst = False
        Next
        If Not String.IsNullOrWhiteSpace(recon.LeftReconSource.Where) And Not isAggA Then
            _sb.AppendLine($"AND {recon.LeftReconSource.Where.Replace("x!", "a")}")
        End If
        If Not String.IsNullOrWhiteSpace(recon.RightReconSource.Where) And Not isAggB Then
            _sb.AppendLine($"AND {recon.RightReconSource.Where.Replace("x!", "b")}")
        End If

        _sb.AppendLine()
        _sb.AppendLine("SELECT * FROM [Match]")
        Return _sb.ToString
    End Function

    Public Shared Function GetDifferSelect(recon As Reconciliation) As String
        Dim isFirst As Boolean = True
        _sb.Clear()
        _sb.AppendLine("IF(OBJECT_ID('Rekonator..[Differ]') IS NOT NULL) DROP TABLE [Differ];")
        _sb.AppendLine()

        Dim cteTable As String = String.Empty
        Dim isAggA As Boolean = (recon.LeftReconSource.Aggregations IsNot Nothing)
        Dim aTable As String = $"[dbo].[{recon.LeftReconSource.ReconTable}] a"
        If isAggA Then
            cteTable = $"cte_{recon.LeftReconSource.ReconTable}_grp"
            _sb.AppendLine(MakeGroupBy(recon.LeftReconSource, cteTable, "a"))
            aTable = $"{cteTable} a"
        End If
        Dim isAggB As Boolean = (recon.RightReconSource.Aggregations IsNot Nothing)
        Dim bTable As String = $"[dbo].[{recon.RightReconSource.ReconTable}] b"
        If isAggB Then
            cteTable = $"cte_{recon.RightReconSource.ReconTable}_grp"
            _sb.AppendLine(MakeGroupBy(recon.RightReconSource, cteTable, "b"))
            bTable = $"{cteTable} b"
        End If
        _sb.AppendLine("SELECT")
        For Each c As Comparision In recon.CompletenessComparisions.Union(recon.MatchingComparisions)
            If Not isFirst Then
                _sb.Append(",")
            End If

            Dim test As String = String.Empty
            Select Case c.ComparisionTest
                Case ComparisionType.TextCaseEquals
                    test = $"ISNULL(a.[{c.LeftColumn}],'') = ISNULL(b.[{c.RightColumn}],'')"
                Case ComparisionType.TextEquals
                    test = $"LOWER(ISNULL(a.[{c.LeftColumn}],'')) = LOWER(ISNULL(b.[{c.RightColumn}],''))"
                Case ComparisionType.NumberEquals
                    test = $"CONVERT(DECIMAL(14,{c.Percision}), ISNULL(a.[{c.LeftColumn}],0)) = CONVERT(DECIMAL(14,{c.Percision}), ISNULL(b.[{c.RightColumn}],0))"
                Case ComparisionType.DateEquals
                    test = $"CONVERT(DATE, a.[{c.LeftColumn}]) = CONVERT(DATE, b.[{c.RightColumn}])"
                Case ComparisionType.DateTimeEquals
                    test = $"CONVERT(DATETIME, a.[{c.LeftColumn}]) = CONVERT(DATETIME, b.[{c.RightColumn}])"
            End Select

            _sb.AppendLine($"[{c.LeftColumn}:{c.RightColumn}] = CONCAT(a.[{c.LeftColumn}], IIf({test}, '=', '<>'), b.[{c.RightColumn}])")
            isFirst = False
        Next
        If isAggA Then _sb.AppendLine(MakeGroupByColumns(recon.LeftReconSource.Aggregations(0).GroupByColumns, "a", True))
        If isAggB Then _sb.AppendLine(MakeGroupByColumns(recon.RightReconSource.Aggregations(0).GroupByColumns, "b", True))
        If Not isAggA Then _sb.AppendLine(",[idA] = a.[rekonid]")
        If Not isAggB Then _sb.AppendLine(", [idB] = b.[rekonid]")
        _sb.AppendLine("INTO [Differ]")
        _sb.AppendLine($"FROM {aTable}, {bTable}")
        _sb.AppendLine("WHERE")
        isFirst = True
        For Each c As Comparision In recon.CompletenessComparisions
            If Not isFirst Then
                _sb.Append("And ")
            End If
            Select Case c.ComparisionTest
                Case ComparisionType.TextCaseEquals
                    _sb.AppendLine($"ISNULL(a.[{c.LeftColumn}],'') = ISNULL(b.[{c.RightColumn}],'')")
                Case ComparisionType.TextEquals
                    _sb.AppendLine($"LOWER(ISNULL(a.[{c.LeftColumn}],'')) = LOWER(ISNULL(b.[{c.RightColumn}],''))")
                Case ComparisionType.NumberEquals
                    _sb.AppendLine($"CONVERT(DECIMAL(14,{c.Percision}), ISNULL(a.[{c.LeftColumn}],0)) = CONVERT(DECIMAL(14,{c.Percision}), ISNULL(b.[{c.RightColumn}],0))")
                Case ComparisionType.DateEquals
                    _sb.AppendLine($"CONVERT(DATE, a.[{c.LeftColumn}]) = CONVERT(DATE, b.[{c.RightColumn}])")
                Case ComparisionType.DateTimeEquals
                    _sb.AppendLine($"CONVERT(DATETIME, a.[{c.LeftColumn}]) = CONVERT(DATETIME, b.[{c.RightColumn}])")
            End Select
            isFirst = False
        Next
        If Not String.IsNullOrWhiteSpace(recon.LeftReconSource.Where) And Not isAggA Then
            _sb.AppendLine($"AND {recon.LeftReconSource.Where.Replace("x!", "a")}")
        End If
        If Not String.IsNullOrWhiteSpace(recon.RightReconSource.Where) And Not isAggB Then
            _sb.AppendLine($"AND {recon.RightReconSource.Where.Replace("x!", "b")}")
        End If
        If isAggA Then
            _sb.AppendLine("AND")
            _sb.AppendLine(MakeNotExists(recon, "Match", "a"))
        End If
        If isAggB Then
            _sb.AppendLine("AND")
            _sb.AppendLine(MakeNotExists(recon, "Match", "b"))

        End If
        _sb.AppendLine()
        _sb.AppendLine("SELECT * FROM [Differ]")
        Return _sb.ToString
    End Function

    Private Shared Function MakeGroupBy(reconSource As ReconSource, cteTable As String, aorb As String) As String
        Dim isFirst As Boolean = True
        Dim sb As New StringBuilder 'don't use _sb
        sb.AppendLine()

        For Each agg As Aggregate In reconSource.Aggregations 'Only Tested for one
            sb.AppendLine($";WITH {cteTable} AS")
            sb.AppendLine("(")
            sb.AppendLine("SELECT")
            sb.AppendLine(MakeGroupByColumns(agg.GroupByColumns, aorb))
            For Each aop As AggregateOperation In agg.AggregateOperations
                sb.AppendLine($",[{aop.AggregateColumn}] = {aop.Operation.ToString}({aorb}.[{aop.SourceColumn}])")
            Next
            sb.AppendLine($"FROM [dbo].[{reconSource.ReconTable}] {aorb}")
            If Not String.IsNullOrWhiteSpace(reconSource.Where) Then
                sb.AppendLine($"WHERE {reconSource.Where.Replace("x!", aorb)}")
            End If
            sb.AppendLine("GROUP BY ")
            sb.AppendLine(MakeGroupByColumns(agg.GroupByColumns, aorb))
            sb.AppendLine(")")
        Next
        Return sb.ToString
    End Function

    Private Shared Function MakeGroupByColumns(groupbycols As String(), aorb As String, Optional isLeadingComma As Boolean = False)
        'If Not String.IsNullOrEmpty(aorb) Then aorb += "."
        'Dim isFirst As Boolean = True

        Return IIf(isLeadingComma, ",", String.Empty) + String.Join(",", groupbycols.ToList().Select(Function(s) $"{aorb}.[{s}]"))
        'Dim gb As New StringBuilder
        'For Each gbc In groupbycols
        'If Not isFirst Then
        'gb.Append(",")
        'End If
        'gb.AppendLine($"{aorb}.[{gbc}]")
        'isFirst = False
        'Next
        'Return gb.ToString
    End Function
    Public Shared Function GetLeftSelect(recon As Reconciliation) As String

        _sb.Clear()
        _sb.AppendLine("IF(OBJECT_ID('Rekonator..[Left]') IS NOT NULL) DROP TABLE [Left];")
        _sb.AppendLine()
        _sb.AppendLine("SELECT a.*")
        _sb.AppendLine("INTO [Left]")
        _sb.AppendLine($"FROM [dbo].[{recon.LeftReconSource.ReconTable}] a")
        _sb.AppendLine("WHERE")
        _sb.AppendLine(MakeNotExists(recon, "Match", "a"))
        _sb.AppendLine("AND")
        _sb.AppendLine(MakeNotExists(recon, "Differ", "a"))
        _sb.AppendLine()
        _sb.AppendLine("SELECT * FROM [Left] a")
        If Not String.IsNullOrWhiteSpace(recon.LeftReconSource.Where) And recon.LeftReconSource.Aggregations IsNot Nothing Then
            _sb.AppendLine($"WHERE {recon.LeftReconSource.Where.Replace("x!.", "a.")}")
        End If
        Return _sb.ToString
    End Function
    Public Shared Function GetRightSelect(recon As Reconciliation) As String
        Dim isAggB As Boolean = (recon.RightReconSource.Aggregations IsNot Nothing)
        _sb.Clear()
        _sb.AppendLine("IF(OBJECT_ID('Rekonator..[Right]') IS NOT NULL) DROP TABLE [Right];")
        _sb.AppendLine()

        If isAggB Then
            _sb.AppendLine("SELECT")
            _sb.AppendLine(MakeGroupByColumns(recon.RightReconSource.Aggregations(0).GroupByColumns, "b"))
            _sb.AppendLine("INTO [Right]")
            _sb.AppendLine($"FROM [dbo].[{recon.RightReconSource.ReconTable}] b")
            _sb.AppendLine("WHERE")
            If Not String.IsNullOrWhiteSpace(recon.RightReconSource.Where) Then
                _sb.AppendLine($"{recon.RightReconSource.Where.Replace("x!.", "b.")}")
                _sb.AppendLine("AND")
            End If
            _sb.AppendLine(MakeNotExists(recon, "Match", "b"))
            _sb.AppendLine("AND")
            _sb.AppendLine(MakeNotExists(recon, "Differ", "b"))
            _sb.AppendLine("GROUP BY ")
            _sb.AppendLine(MakeGroupByColumns(recon.RightReconSource.Aggregations(0).GroupByColumns, "b"))
            _sb.AppendLine()
            _sb.AppendLine("SELECT * FROM [Right] b")
        Else
            _sb.AppendLine("SELECT b.*")
            _sb.AppendLine("INTO [Right]")
            _sb.AppendLine($"FROM [dbo].[{recon.RightReconSource.ReconTable}] b")
            _sb.AppendLine("WHERE")
            _sb.AppendLine("NOT EXISTS (SELECT * FROM [Match] m WHERE {m.IdB = b.rekonid)")
            _sb.AppendLine("AND")
            _sb.AppendLine("NOT EXISTS (SELECT * FROM [Differ] d WHERE {d.IdB = b.rekonid)")
            _sb.AppendLine()
            _sb.AppendLine("SELECT * FROM [Right] b")
            If Not String.IsNullOrWhiteSpace(recon.RightReconSource.Where) Then
                _sb.AppendLine($"WHERE {recon.RightReconSource.Where.Replace("x!.", "b.")}")
            End If

        End If


        Return _sb.ToString

    End Function

    Private Shared Function MakeNotExists(recon As Reconciliation, tableName As String, aorb As String) As String
        'use aorb.  a for left, b for right.  don't use isggA or isggB since left won't have b alias and right won't have a alias
        Dim isFirst As Boolean = True
        Dim sb As New StringBuilder 'don't use _sb
        Dim mord As String = Left(tableName, 1).ToLower
        Dim isAggA As Boolean = (recon.LeftReconSource.Aggregations IsNot Nothing)
        Dim isAggB As Boolean = (recon.RightReconSource.Aggregations IsNot Nothing)

        sb.Append($"NOT EXISTS (SELECT * FROM [{tableName}] {mord} WHERE ")
        isFirst = True
        If isAggA Then
            For Each agg As Aggregate In recon.LeftReconSource.Aggregations
                For Each gbc In agg.GroupByColumns
                    If Not isFirst Then
                        sb.Append("AND ")
                    End If
                    sb.AppendLine($"{mord}.[{gbc}] = {aorb}.[{gbc}]")
                    isFirst = False
                Next
            Next
        End If
        If isAggB Then
            For Each agg As Aggregate In recon.RightReconSource.Aggregations
                For Each gbc In agg.GroupByColumns
                    If Not isFirst Then
                        sb.Append("AND ")
                    End If
                    sb.AppendLine($"{mord}.[{gbc}] = {aorb}.[{gbc}]")
                    isFirst = False
                Next
            Next
        End If
        sb.AppendLine(")")

        Return sb.ToString
    End Function



End Class
