Imports System.Text

<Serializable()>
Public Class Reconciliation
    Property ReconciliationName As String
    Property LeftReconSource As ReconSource
    Property RightReconSource As ReconSource
    Property CompletenessComparisions As List(Of Comparision)
    Property MatchingComparisions As List(Of Comparision)

    Private Shared _sb As New StringBuilder

    Public Shared Reconciliations As New List(Of Reconciliation)
    Public Shared Sub Add(reconciliationName As String,
                          leftDataSource As ReconSource,
                          rightDataSource As ReconSource,
                          completenessComparision As List(Of Comparision),
                          matchingComparision As List(Of Comparision))
        Reconciliations.Add(New Reconciliation With {
                            .ReconciliationName = reconciliationName,
                            .LeftReconSource = leftDataSource,
                            .RightReconSource = rightDataSource,
                            .CompletenessComparisions = completenessComparision,
                            .MatchingComparisions = matchingComparision}
                            )
    End Sub
    Public Shared Function GetReconciliation(reconciliationName As String) As Reconciliation
        Return Reconciliations.Where(Function(w) w.ReconciliationName = reconciliationName).FirstOrDefault
    End Function

    Public Shared Function GetMatchSelect(recon As Reconciliation) As String
        Dim isFirst As Boolean = True
        _sb.Clear
        _sb.AppendLine("IF(OBJECT_ID('Rekonator..[Match]') IS NOT NULL) DROP TABLE [Match];")
        _sb.AppendLine()
        _sb.AppendLine("SELECT")
        For Each c As Comparision In recon.CompletenessComparisions.Union(recon.MatchingComparisions)
            If Not isFirst Then
                _sb.Append(",")
            End If
            _sb.AppendLine($"[{c.LeftColumn}={c.RightColumn}] = a.[{c.LeftColumn}]")
            isFirst = False
        Next
        _sb.AppendLine(",[idA] = a.[rekonid], [idB] = b.[rekonid]")
        _sb.AppendLine("INTO [Match]")
        _sb.AppendLine($"FROM [dbo].[{recon.LeftReconSource.ReconTable}] a, [dbo].[{recon.RightReconSource.ReconTable}] b")
        _sb.AppendLine("WHERE")

        isFirst = True
        For Each c As Comparision In recon.CompletenessComparisions.Union(recon.MatchingComparisions)
            If Not isFirst Then
                _sb.Append("AND ")
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
        If Not String.IsNullOrWhiteSpace(recon.LeftReconSource.Where) Then
            _sb.AppendLine($"AND {recon.LeftReconSource.Where.Replace("x!", "a")}")
        End If
        If Not String.IsNullOrWhiteSpace(recon.RightReconSource.Where) Then
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
        _sb.AppendLine(",[idA] = a.[rekonid], [idB] = b.[rekonid]")
        _sb.AppendLine("INTO [Differ]")
        _sb.AppendLine($"FROM [dbo].[{recon.LeftReconSource.ReconTable}] a, [dbo].[{recon.RightReconSource.ReconTable}] b")
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
        If Not String.IsNullOrWhiteSpace(recon.LeftReconSource.Where) Then
            _sb.AppendLine($"AND {recon.LeftReconSource.Where.Replace("x!", "a")}")
        End If
        If Not String.IsNullOrWhiteSpace(recon.RightReconSource.Where) Then
            _sb.AppendLine($"AND {recon.RightReconSource.Where.Replace("x!", "b")}")
        End If
        _sb.AppendLine("AND")
        _sb.AppendLine("NOT EXISTS (SELECT * FROM [Match] m WHERE m.IdA = a.rekonid)")

        _sb.AppendLine()
        _sb.AppendLine("SELECT * FROM [Differ]")
        Return _sb.ToString
    End Function

    Public Shared Function GetLeftSelect(recon As Reconciliation) As String
        _sb.Clear()
        _sb.AppendLine("IF(OBJECT_ID('Rekonator..[Left]') IS NOT NULL) DROP TABLE [Left];")
        _sb.AppendLine()
        _sb.AppendLine("SELECT a.*, [idA] = a.[rekonid]")
        _sb.AppendLine("INTO [Left]")
        _sb.AppendLine($"FROM [dbo].[{recon.LeftReconSource.ReconTable}] a")
        _sb.AppendLine("WHERE")
        _sb.AppendLine("NOT EXISTS (SELECT * FROM [Match] m WHERE m.[idA] = a.[rekonid])")
        _sb.AppendLine("AND")
        _sb.AppendLine("NOT EXISTS (SELECT * FROM [Differ] d WHERE d.[idA] = a.[rekonid])")
        _sb.AppendLine()
        _sb.AppendLine("SELECT * FROM [Left] a")
        If Not String.IsNullOrWhiteSpace(recon.LeftReconSource.Where) Then
            _sb.AppendLine($"WHERE {recon.LeftReconSource.Where.Replace("x!.", "a.")}")
        End If
        Return _sb.ToString
    End Function
    Public Shared Function GetRightSelect(recon As Reconciliation) As String
        _sb.Clear()
        _sb.AppendLine("IF(OBJECT_ID('Rekonator..[Right]') IS NOT NULL) DROP TABLE [Right];")
            _sb.AppendLine()
        _sb.AppendLine("SELECT b.*, [idB] = b.[rekonid]")
        _sb.AppendLine("INTO [Right]")
        _sb.AppendLine($"FROM [dbo].[{recon.RightReconSource.ReconTable}] b")
        _sb.AppendLine("WHERE")
        _sb.AppendLine("NOT EXISTS (SELECT * FROM [Match] m WHERE m.[idB] = b.[rekonid])")
        _sb.AppendLine("AND")
        _sb.AppendLine("NOT EXISTS (SELECT * FROM [Differ] d WHERE d.[idB] = b.[rekonid])")
        _sb.AppendLine()
        _sb.AppendLine("SELECT * FROM [Right] b")
        If Not String.IsNullOrWhiteSpace(recon.RightReconSource.Where) Then
            _sb.AppendLine($"WHERE {recon.RightReconSource.Where.Replace("x!.", "b.")}")
        End If

        Return _sb.ToString

    End Function

End Class
