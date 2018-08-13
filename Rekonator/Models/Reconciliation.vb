Imports System.Text
Imports Rekonator

<Serializable()>
Public Class Reconciliation
    Property ReconciliationName As String = "(New Reconciliation)"
    Property LeftReconSource As New ReconSource
    Property RightReconSource As New ReconSource
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
        _sb.AppendLine("--Use Rekonator")

        Dim cteTable As String = String.Empty
        Dim isAggA As Boolean = (recon.LeftReconSource.Aggregations IsNot Nothing)
        Dim aTable As String = $"[dbo].[{recon.LeftReconSource.ReconTable}] a"
        If isAggA Then
            cteTable = $"cte_{recon.LeftReconSource.ReconTable}_grp"
            _sb.AppendLine(MakeGroupBy(recon.LeftReconSource, cteTable, "a", isFirst))
            aTable = $"{cteTable} a"
            isFirst = False
        End If
        Dim isAggB As Boolean = (recon.RightReconSource.Aggregations IsNot Nothing)
        Dim bTable As String = $"[dbo].[{recon.RightReconSource.ReconTable}] b"
        If isAggB Then
            cteTable = $"cte_{recon.RightReconSource.ReconTable}_grp"
            _sb.AppendLine(MakeGroupBy(recon.RightReconSource, cteTable, "b", isFirst))
            bTable = $"{cteTable} b"
        End If

        _sb.AppendLine("SELECT")
        isFirst = True
        For Each c As Comparision In recon.CompletenessComparisions.Union(recon.MatchingComparisions)
            If Not isFirst Then
                _sb.Append(",")
            End If
            _sb.AppendLine($"[{c.LeftColumn}={c.RightColumn}] = a.[{c.LeftColumn}]")
            isFirst = False
        Next
        If isAggA Then _sb.AppendLine(MakeGroupByColumns(recon.LeftReconSource.Aggregations(0).GroupByColumns, "a", recon.LeftReconSource.ColumnPrefix, True))
        If isAggB Then _sb.AppendLine(MakeGroupByColumns(recon.RightReconSource.Aggregations(0).GroupByColumns, "b", recon.RightReconSource.ColumnPrefix, True))
        If Not isAggA Then _sb.AppendLine(",[idA] = a.[rekonid]")
        If Not isAggB Then _sb.AppendLine(", [idB] = b.[rekonid]")
        _sb.AppendLine("INTO [Match]")
        _sb.AppendLine($"FROM {aTable}, {bTable}")
        _sb.AppendLine("WHERE")
        _sb.AppendLine(MakeWhereComparision(recon.CompletenessComparisions, recon.LeftReconSource.ColumnPrefix, recon.RightReconSource.ColumnPrefix))
        _sb.Append("AND ")
        _sb.AppendLine(MakeWhereComparision(recon.MatchingComparisions, recon.LeftReconSource.ColumnPrefix, recon.RightReconSource.ColumnPrefix))

        If Not String.IsNullOrWhiteSpace(recon.LeftReconSource.WhereClause) And Not isAggA Then
            _sb.AppendLine($"AND {recon.LeftReconSource.WhereClause.Replace("x!", "a")}")
        End If
        If Not String.IsNullOrWhiteSpace(recon.RightReconSource.WhereClause) And Not isAggB Then
            _sb.AppendLine($"AND {recon.RightReconSource.WhereClause.Replace("x!", "b")}")
        End If

        _sb.AppendLine()
        _sb.AppendLine("SELECT * FROM [Match]")
        Return _sb.ToString
    End Function

    'Private Shared Function MakeDropTable(tableName As String) As String
    '    Dim mdt As New StringBuilder
    '    mdt.AppendLine("USE;")
    '    mdt.AppendLine($"IF (OBJECT_ID('Rekonator.[dbo].[{tableName}]') IS NOT NULL) DROP TABLE Rekonator.[dbo].[{tableName}];")
    '    mdt.AppendLine("COMMIT TRANSACTION;")
    '    mdt.AppendLine()
    '    Return mdt.ToString
    'End Function

    Public Shared Function GetDifferSelect(recon As Reconciliation) As String
        Dim isFirst As Boolean = True
        _sb.Clear()
        _sb.AppendLine("--Use Rekonator")

        Dim cteTable As String = String.Empty
        Dim isAggA As Boolean = (recon.LeftReconSource.Aggregations IsNot Nothing)
        Dim aTable As String = $"[dbo].[{recon.LeftReconSource.ReconTable}] a"
        If isAggA Then
            cteTable = $"cte_{recon.LeftReconSource.ReconTable}_grp"
            _sb.AppendLine(MakeGroupBy(recon.LeftReconSource, cteTable, "a", isFirst))
            aTable = $"{cteTable} a"
            isFirst = False
        End If
        Dim isAggB As Boolean = (recon.RightReconSource.Aggregations IsNot Nothing)
        Dim bTable As String = $"[dbo].[{recon.RightReconSource.ReconTable}] b"
        If isAggB Then
            cteTable = $"cte_{recon.RightReconSource.ReconTable}_grp"
            _sb.AppendLine(MakeGroupBy(recon.RightReconSource, cteTable, "b", isFirst))
            bTable = $"{cteTable} b"
        End If

        _sb.AppendLine("SELECT")
        isFirst = True
        For Each c As Comparision In recon.CompletenessComparisions.Union(recon.MatchingComparisions)
            If Not isFirst Then
                _sb.Append(",")
            End If

            Dim test As String = String.Empty
            Dim diff As String = String.Empty
            Dim aPrefix As String = recon.LeftReconSource.ColumnPrefix
            Dim bPrefix As String = recon.RightReconSource.ColumnPrefix
            Select Case c.ComparisionTest
                Case ComparisionType.TextCaseEquals
                    test = $"ISNULL(a.[{aPrefix}{c.LeftColumn}],'') = ISNULL(b.[{bPrefix}{c.RightColumn}],'')"
                Case ComparisionType.TextEquals
                    test = $"LOWER(ISNULL(a.[{aPrefix}{c.LeftColumn}],'')) = LOWER(ISNULL(b.[{bPrefix}{c.RightColumn}],''))"
                Case ComparisionType.NumberEquals
                    test = $"CONVERT(DECIMAL(14,{c.Percision}), ISNULL(a.[{aPrefix}{c.LeftColumn}],0)) = CONVERT(DECIMAL(14,{c.Percision}), ISNULL(b.[{bPrefix}{c.RightColumn}],0))"
                    diff = $"CONVERT(DECIMAL(14,{c.Percision}), (ISNULL(a.[{aPrefix}{c.LeftColumn}],0) - ISNULL(b.[{bPrefix}{c.RightColumn}],0)))"
                Case ComparisionType.DateEquals
                    test = $"CONVERT(DATE, a.[{aPrefix}{c.LeftColumn}]) = CONVERT(DATE, b.[{bPrefix}{c.RightColumn}])"
                Case ComparisionType.DateTimeEquals
                    test = $"CONVERT(DATETIME, a.[{aPrefix}{c.LeftColumn}]) = CONVERT(DATETIME, b.[{bPrefix}{c.RightColumn}])"
            End Select

            _sb.Append($"[{c.LeftColumn}:{c.RightColumn}] = CONCAT(a.[{c.LeftColumn}], IIf({test}, '=', '<>'), b.[{bPrefix}{c.RightColumn}]")
            If String.IsNullOrEmpty(diff) Then
                _sb.AppendLine(")")
            Else
                _sb.AppendLine($", IIf({test}, NULL, CONCAT(':', {diff})))")
            End If
            isFirst = False
        Next
        If isAggA Then _sb.AppendLine(MakeGroupByColumns(recon.LeftReconSource.Aggregations(0).GroupByColumns, "a", recon.LeftReconSource.ColumnPrefix, True))
        If isAggB Then _sb.AppendLine(MakeGroupByColumns(recon.RightReconSource.Aggregations(0).GroupByColumns, "b", recon.RightReconSource.ColumnPrefix, True))
        If Not isAggA Then _sb.AppendLine(",[idA] = a.[rekonid]")
        If Not isAggB Then _sb.AppendLine(", [idB] = b.[rekonid]")
        _sb.AppendLine("INTO [Differ]")
        _sb.AppendLine($"FROM {aTable}, {bTable}")
        _sb.AppendLine("WHERE")
        _sb.AppendLine(MakeWhereComparision(recon.CompletenessComparisions, recon.LeftReconSource.ColumnPrefix, recon.RightReconSource.ColumnPrefix))


        If Not String.IsNullOrWhiteSpace(recon.LeftReconSource.WhereClause) And Not isAggA Then
            _sb.AppendLine($"AND {recon.LeftReconSource.WhereClause.Replace("x!", "a")}")
        End If
        If Not String.IsNullOrWhiteSpace(recon.RightReconSource.WhereClause) And Not isAggB Then
            _sb.AppendLine($"AND {recon.RightReconSource.WhereClause.Replace("x!", "b")}")
        End If
        If isAggA Then
            _sb.AppendLine("AND")
            _sb.AppendLine(MakeNotExists1(recon.LeftReconSource, "Match", "a", recon.LeftReconSource.ColumnPrefix))
        End If
        If isAggB Then
            _sb.AppendLine("AND")
            _sb.AppendLine(MakeNotExists1(recon.RightReconSource, "Match", "b", recon.RightReconSource.ColumnPrefix))

        End If
        _sb.AppendLine()
        _sb.AppendLine("SELECT * FROM [Differ]")
        Return _sb.ToString
    End Function

    Private Shared Function MakeWhereComparision(completenessComparisions As List(Of Comparision), aPrefix As String, bPrefix As String) As String
        Dim isFirst As Boolean = True
        Dim cpc As New StringBuilder
        For Each c As Comparision In completenessComparisions
            If Not isFirst Then
                cpc.Append("And ")
            End If
            Dim lCol As String = $"a.[{aPrefix}{c.LeftColumn}]"
            Dim rCol As String = $"b.[{bPrefix}{c.RightColumn}]"
            If Not String.IsNullOrWhiteSpace(c.RightFunction) Then
                'RightFunction="SUBSTRING(b.[{c.RightColumn}], 11,  LEN(b.[{c.RightColumn}]) -10)"
                rCol = c.RightFunction.Replace("{RightColumn}", rCol)
            End If
            Select Case c.ComparisionTest
                Case ComparisionType.TextCaseEquals
                    cpc.AppendLine($"ISNULL({lCol},'') = ISNULL({rCol},'')")
                Case ComparisionType.TextEquals
                    cpc.AppendLine($"LOWER(ISNULL({lCol},'')) = LOWER(ISNULL({rCol},''))")
                Case ComparisionType.NumberEquals
                    cpc.AppendLine($"CONVERT(DECIMAL(14,{c.Percision}), ISNULL({lCol},0)) = CONVERT(DECIMAL(14,{c.Percision}), ISNULL({rCol},0))")
                Case ComparisionType.DateEquals
                    cpc.AppendLine($"CONVERT(DATE, {lCol}) = CONVERT(DATE, {rCol})")
                Case ComparisionType.DateTimeEquals
                    cpc.AppendLine($"CONVERT(DATETIME, {lCol}) = CONVERT(DATETIME, {rCol})")
            End Select
            isFirst = False
        Next
        Return cpc.ToString
    End Function

    Private Shared Function MakeGroupBy(reconSource As ReconSource, cteTable As String, aorb As String, isWITH As Boolean) As String
        Dim isFirst As Boolean = True
        Dim prefix As String = reconSource.ColumnPrefix
        Dim sb As New StringBuilder 'don't use _sb
        sb.AppendLine()

        For Each agg As Aggregate In reconSource.Aggregations 'Only Tested for one
            If isWITH Then
                sb.AppendLine($";WITH {cteTable} AS")
            Else
                sb.AppendLine($",{cteTable} AS")
            End If
            sb.AppendLine("(")
            sb.AppendLine("SELECT")
            sb.AppendLine(MakeGroupByColumns(agg.GroupByColumns, aorb, prefix))
            For Each aop As AggregateOperation In agg.AggregateOperations
                sb.AppendLine($",[{prefix}{aop.AggregateColumn}] = {aop.Operation.ToString}({aorb}.[{prefix}{aop.SourceColumn}])")
            Next
            sb.AppendLine($"FROM [dbo].[{reconSource.ReconTable}] {aorb}")
            If Not String.IsNullOrWhiteSpace(reconSource.WhereClause) Then
                sb.AppendLine($"WHERE {reconSource.WhereClause.Replace("x!", aorb)}")
            End If
            sb.AppendLine("GROUP BY ")
            sb.AppendLine(MakeGroupByColumns(agg.GroupByColumns, aorb, prefix))
            sb.AppendLine(")")
        Next
        Return sb.ToString
    End Function

    Private Shared Function MakeGroupByColumns(groupbycols As String(), aorb As String, Optional prefix As String = "", Optional isLeadingComma As Boolean = False)
        'If Not String.IsNullOrEmpty(aorb) Then aorb += "."
        'Dim isFirst As Boolean = True

        Return IIf(isLeadingComma, ",", String.Empty) + String.Join(",", groupbycols.ToList().Select(Function(s) $"{aorb}.[{prefix}{s}]"))
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
        Dim isAggA As Boolean = (recon.LeftReconSource.Aggregations IsNot Nothing)
        Dim prefix As String = recon.LeftReconSource.ColumnPrefix

        _sb.Clear()
        _sb.AppendLine("--Use Rekonator")

        If isAggA Then
            _sb.AppendLine("SELECT")
            _sb.AppendLine(MakeGroupByColumns(recon.LeftReconSource.Aggregations(0).GroupByColumns, "a", prefix))
            For Each aop As AggregateOperation In recon.LeftReconSource.Aggregations(0).AggregateOperations
                _sb.AppendLine($",[{aop.AggregateColumn}] = {aop.Operation.ToString}(a.[{prefix}{aop.SourceColumn}])")
            Next
            _sb.AppendLine("INTO [Left]")
            _sb.AppendLine($"FROM [dbo].[{recon.LeftReconSource.ReconTable}] a")
            _sb.AppendLine("WHERE")
            If Not String.IsNullOrWhiteSpace(recon.LeftReconSource.WhereClause) Then
                _sb.AppendLine($"{recon.LeftReconSource.WhereClause.Replace("x!.", "a.")}")
                _sb.AppendLine("AND")
            End If
            _sb.AppendLine(MakeNotExists1(recon.LeftReconSource, "Match", "a", prefix))
            _sb.AppendLine("AND")
            _sb.AppendLine(MakeNotExists1(recon.LeftReconSource, "Differ", "a", prefix))
            _sb.AppendLine("GROUP BY ")
            _sb.AppendLine(MakeGroupByColumns(recon.LeftReconSource.Aggregations(0).GroupByColumns, "a", prefix))
            _sb.AppendLine()
            _sb.AppendLine("SELECT * FROM [Left] a")
        Else
            _sb.AppendLine("SELECT a.*")
            _sb.AppendLine("INTO [Left]")
            _sb.AppendLine($"FROM [dbo].[{recon.LeftReconSource.ReconTable}] a")
            _sb.AppendLine("WHERE")
            _sb.AppendLine("NOT EXISTS (SELECT * FROM [Match] m WHERE m.IdA = a.rekonid)")
            _sb.AppendLine("AND")
            _sb.AppendLine("NOT EXISTS (SELECT * FROM [Differ] d WHERE d.IdA = a.rekonid)")
            _sb.AppendLine()
            _sb.AppendLine("SELECT * FROM [Left] a")
            If Not String.IsNullOrWhiteSpace(recon.LeftReconSource.WhereClause) Then
                _sb.AppendLine($"WHERE {recon.LeftReconSource.WhereClause.Replace("x!.", "a.")}")
            End If
        End If
        Return _sb.ToString
    End Function
    Public Shared Function GetRightSelect(recon As Reconciliation) As String
        Dim isAggB As Boolean = (recon.RightReconSource.Aggregations IsNot Nothing)
        Dim prefix As String = recon.RightReconSource.ColumnPrefix

        _sb.Clear()
        _sb.AppendLine("--Use Rekonator")

        If isAggB Then
            _sb.AppendLine("SELECT")
            _sb.AppendLine(MakeGroupByColumns(recon.RightReconSource.Aggregations(0).GroupByColumns, "b", prefix))
            For Each aop As AggregateOperation In recon.RightReconSource.Aggregations(0).AggregateOperations
                _sb.AppendLine($",[{aop.AggregateColumn}] = {aop.Operation.ToString}(b.[{prefix}{aop.SourceColumn}])")
            Next
            _sb.AppendLine("INTO [Right]")
            _sb.AppendLine($"FROM [dbo].[{recon.RightReconSource.ReconTable}] b")
            _sb.AppendLine("WHERE")
            If Not String.IsNullOrWhiteSpace(recon.RightReconSource.WhereClause) Then
                _sb.AppendLine($"{recon.RightReconSource.WhereClause.Replace("x!.", "b.")}")
                _sb.AppendLine("AND")
            End If
            _sb.AppendLine(MakeNotExists1(recon.RightReconSource, "Match", "b", prefix))
            _sb.AppendLine("AND")
            _sb.AppendLine(MakeNotExists1(recon.RightReconSource, "Differ", "b", prefix))
            _sb.AppendLine("GROUP BY ")
            _sb.AppendLine(MakeGroupByColumns(recon.RightReconSource.Aggregations(0).GroupByColumns, "b", prefix))
            _sb.AppendLine()
            _sb.AppendLine("SELECT * FROM [Right] b")
        Else
            _sb.AppendLine("SELECT b.*")
            _sb.AppendLine("INTO [Right]")
            _sb.AppendLine($"FROM [dbo].[{recon.RightReconSource.ReconTable}] b")
            _sb.AppendLine("WHERE")
            _sb.AppendLine("NOT EXISTS (SELECT * FROM [Match] m WHERE m.IdB = b.rekonid)")
            _sb.AppendLine("AND")
            _sb.AppendLine("NOT EXISTS (SELECT * FROM [Differ] d WHERE d.IdB = b.rekonid)")
            _sb.AppendLine()
            _sb.AppendLine("SELECT * FROM [Right] b")
            If Not String.IsNullOrWhiteSpace(recon.RightReconSource.WhereClause) Then
                _sb.AppendLine($"WHERE {recon.RightReconSource.WhereClause.Replace("x!.", "b.")}")
            End If

        End If


        Return _sb.ToString

    End Function

    Private Shared Function MakeNotExists2(recon As Reconciliation, tableName As String, aorb As String) As String
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

    Private Shared Function MakeNotExists1(reconSource As ReconSource, tableName As String, aorb As String, prefix As String) As String
        'use aorb.  a for left, b for right.  don't use isggA or isggB since left won't have b alias and right won't have a alias
        Dim isFirst As Boolean = True
        Dim sb As New StringBuilder 'don't use _sb
        Dim mord As String = Left(tableName, 1).ToLower

        sb.Append($"NOT EXISTS (SELECT * FROM [{tableName}] {mord} WHERE ")
        isFirst = True
        For Each agg As Aggregate In reconSource.Aggregations
            For Each gbc In agg.GroupByColumns
                If Not isFirst Then
                    sb.Append("AND ")
                End If
                sb.AppendLine($"ISNULL({mord}.[{prefix}{gbc}],'') = ISNULL({aorb}.[{prefix}{gbc}],'')")
                isFirst = False
            Next
        Next
        sb.AppendLine(")")

        Return sb.ToString
    End Function

End Class
