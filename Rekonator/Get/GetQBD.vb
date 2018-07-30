Imports System.Data
Imports System.Reflection
Imports Dynamitey.Dynamic
Imports Interop.QBFC13

Public Class GetQBD
    Implements IDisposable

    Private _sessionManager As QBSessionManager
    Private _msgSetRequest As IMsgSetRequest
    Private _msgSetResponse As IMsgSetResponse
    Private _response As IResponse
    'Private _resultTable As DataTable

    Private _sql As SQL
    Private _fieldCount As Integer = 0

    Public Function LoadReport(reconSource As ReconSource, fromDate As DateTime, toDate As DateTime) As Boolean
        If Not ConnectToQB() Then Exit Function

        Dim msgSetResponse As IMsgSetResponse = Nothing
        Try
            _msgSetRequest.ClearRequests()
            _msgSetRequest.Attributes.OnError = ENRqOnError.roeContinue

            Dim QReport As IGeneralDetailReportQuery

            QReport = _msgSetRequest.AppendGeneralDetailReportQueryRq
            QReport.GeneralDetailReportType.SetValue(ENGeneralDetailReportType.gdrtTxnListByDate)
            'QReport.GeneralDetailReportType.SetValue(ENGeneralDetailReportType.gdrtProfitAndLossDetail)
            QReport.DisplayReport.SetValue(True)
            'QReport.ReportBasis.SetValue(ENReportBasis.rbAccrual)
            'QReport.ORReportPeriod.ReportDateMacro.SetValue(ENReportDateMacro.rdmLastQuarter)

            QReport.ORReportPeriod.ReportPeriod.FromReportDate.SetValue(fromDate)
            QReport.ORReportPeriod.ReportPeriod.ToReportDate.SetValue(toDate)
            QReport.IncludeColumnList.Add(ENIncludeColumn.icAccount)
            QReport.IncludeColumnList.Add(ENIncludeColumn.icDate)
            QReport.IncludeColumnList.Add(ENIncludeColumn.icRefNumber)
            QReport.IncludeColumnList.Add(ENIncludeColumn.icClass)
            QReport.IncludeColumnList.Add(ENIncludeColumn.icTxnNumber)
            QReport.IncludeColumnList.Add(ENIncludeColumn.icAmount)
            QReport.IncludeColumnList.Add(ENIncludeColumn.icTxnType)
            QReport.IncludeColumnList.Add(ENIncludeColumn.icTxnID)
            QReport.ReportTxnTypeFilter.TxnTypeFilterList.Add(ENTxnTypeFilter.ttfCreditMemo)
            QReport.ReportTxnTypeFilter.TxnTypeFilterList.Add(ENTxnTypeFilter.ttfInvoice)
            QReport.ReportTxnTypeFilter.TxnTypeFilterList.Add(ENTxnTypeFilter.ttfJournalEntry)
            QReport.ReportTxnTypeFilter.TxnTypeFilterList.Add(ENTxnTypeFilter.ttfCheck)
            _msgSetResponse = _sessionManager.DoRequests(_msgSetRequest)
            CloseQB()

            Dim headerList As List(Of String) = {"Account", "Date", "Number", "Class", "Trans #", "Amount", "Type", "TxnID"}.ToList
            Dim typeList As List(Of String) = {"String", "Date", "String", "String", "Integer", "Currency", "String", "String"}.ToList

            _fieldCount = headerList.Count
            _sql = New SQL(reconSource.ReconTable, _fieldCount, headerList, typeList)
            If Not _sql.CreateTable() Then
                Return False
            End If
            WalkGeneralDetailReportQueryRs()
            Return True

        Catch ex As Exception
            Application.ErrorMessage($"Error Loading Report:  {ex.Message}")
        End Try

        Return Nothing
    End Function

    Private Sub CloseQB()
        _sessionManager.EndSession()
        _sessionManager.CloseConnection()
    End Sub

    Public Function GetList(ByVal tableName As String,
                              Optional ByRef fields As String() = Nothing,
                              Optional ByRef criteria As Dictionary(Of String, String) = Nothing,
                              Optional ByRef iterationSize As Integer = 0,
                              Optional ByRef iterationRemaining As Integer = 0,
                              Optional ByRef iteratorID As String = "",
                              Optional ByVal includeLines As Boolean = False,
                              Optional ByVal includeLinkedTnx As Boolean = False) As DataTable

        If Not ConnectToQB() Then Exit Function

        Dim msgSetResponse As IMsgSetResponse = Nothing
        Try
            _msgSetRequest.ClearRequests()
            _msgSetRequest.Attributes.OnError = ENRqOnError.roeContinue

            Dim byListID As Boolean = False
            Dim QQuery As Object = Nothing

            If Not IsNothing(criteria) Then
                For Each d In criteria
                    If d.Key.ToUpper = "[MODIFIED]" Then
                        Dim lastModifiedTime As Date
                        lastModifiedTime = CDate(d.Value)
                        'Message($"Query Last Modified Time: {lastModifiedTime}")
                        Select Case tableName
                            Case "Item"
                                QQuery.ORListQueryWithOwnerIDAndClass.ListWithClassFilter.FromModifiedDate.SetValue(lastModifiedTime, False)
                            Case "Invoice"
                                'For Each member As String In Impromptu.GetMemberNames(QQuery)
                                '    Message(member)
                                'Next
                                InvokeGetChain(QQuery, String.Format(".OR{0}{1}Query.{0}Filter.ORDateRangeFilter.ModifiedDateRangeFilter.FromModifiedDate", tableName, If(tableName = "Charge", "Txn", ""))).SetValue(lastModifiedTime, False)
                        End Select
                    ElseIf d.Key.StartsWith("[TxnDate") Then
                        Dim txnDate As String = CDate(d.Value).ToString("yyyy-MM-dd")
                        'Message($"Query TxnDate: {txnDate}")
                        Dim fieldName As String = IIf(d.Key.EndsWith("To"), "To", "From")
                        Select Case tableName
                            Case "Item"
                                InvokeGetChain(QQuery, String.Format("OR{0}{1}Query.{0}Filter.ORDateRangeFilter.TxnDateRangeFilter.ORTxnDateRangeFilter.TxnDateFilter.{2}TxnDate", tableName, If(tableName = "Charge", "Txn", ""), fieldName)).SetValue(txnDate)
                        End Select
                    End If
                Next
            End If

            'Active Records
            If True Then
                Select Case tableName
                    Case "Item"
                        QQuery.ORListQueryWithOwnerIDAndClass.ListWithClassFilter.ActiveStatus.SetValue(ENActiveStatus.asActiveOnly)
                End Select
            End If

            If iterationSize > 0 Then
                If iteratorID = "" Then
                    QQuery.iterator.SetValue(ENiterator.itStart)
                Else
                    QQuery.iterator.SetValue(ENiterator.itContinue)
                    QQuery.iteratorID.SetValue(iteratorID)
                End If
                'Set the IterationSize
                Select Case tableName
                    Case "Item"
                        InvokeGetChain(QQuery, String.Format("OR{0}{1}Query.{0}Filter.MaxReturned", tableName, If(tableName = "Charge", "Txn", ""))).SetValue(iterationSize)
                End Select
            End If

            If iterationSize > 0 Then
                Select Case tableName
                    Case "Customer"
                        QQuery.ORCustomerListQuery.CustomerListFilter.MaxReturned.SetValue(iterationSize)
                        If iteratorID > "" Then
                            QQuery.ORCustomerListQuery.CustomerListFilter.ORNameFilter.NameRangeFilter.FromName.setvalue(iteratorID + " ")
                        End If
                    Case "ItemService"
                        QQuery.ORListQueryWithOwnerIDAndClass.ListWithClassFilter.MaxReturned.SetValue(iterationSize)
                        If iteratorID > "" Then
                            QQuery.ORListQueryWithOwnerIDAndClass.ListWithClassFilter.ORNameFilter.NameRangeFilter.FromName.setvalue(iteratorID + " ")
                        End If
                End Select
            End If

            ' Add the request to the message set request object
            Dim bIncludeCustomFields As Boolean = False
            If IsNothing(fields) OrElse fields.Count = 0 Then
                'if no fields were passed in the assume all fields including custom one
                'If _qbCustomFieldTables.Contains(tableName) Then bIncludeCustomFields = True
            Else
                'Can only call out fields if custom fields aren't needed
                If Not bIncludeCustomFields Then
                    For ndx = 0 To fields.Count - 1
                        Dim sField As String = fields(ndx)
                        Dim iPos As Integer = InStr(sField, "Ref")
                        If iPos > 1 Then sField = Left(fields(ndx), iPos + 2)

                        If fields(ndx).StartsWith("CustomField:") And Not bIncludeCustomFields Then
                            bIncludeCustomFields = True
                            sField = "DataExtRet"
                        End If
                        If fields(ndx).Contains(".") Then
                            For Each sPart As String In fields(ndx).Split(".")
                                QQuery.IncludeRetElementList.Add(sPart)
                            Next
                        Else
                            QQuery.IncludeRetElementList.Add(sField)
                        End If
                    Next
                End If
            End If

            If bIncludeCustomFields Then
                QQuery.OwnerIDList.Add("0")
            End If

            If includeLines Then QQuery.IncludeLineItems.SetValue(True)

            ' Do the request and get the response message set object
            msgSetResponse = _sessionManager.DoRequests(_msgSetRequest)

            ' Uncomment the following to view and save the request and response XML
            ' string requestXML = requestSet.ToXMLString();
            ' MessageBox.Show(requestXML);
            ' SaveXML(requestXML);
            ' string responseXML = responseSet.ToXMLString();
            ' MessageBox.Show(responseXML);
            ' SaveXML(responseXML);

            _response = msgSetResponse.ResponseList.GetAt(0)
            ' int statusCode = response.StatusCode;
            ' string statusMessage = response.StatusMessage;
            ' string statusSeverity = response.StatusSeverity;
            ' MessageBox.Show("Status:\nCode = " + statusCode + "\nMessage = " + statusMessage + "\nSeverity = " + statusSeverity);

            'Dim customerRetList As ICustomerRetList = response.Detail
            'For ndx As Integer = 0 To (customerRetList.Count - 1)
            'Dim customerRet As ICustomerRet = customerRetList.GetAt(ndx)
            'Debug.Print(customerRet.ListID.GetValue)
            'Next


            'Return TryCast(response.Detail, IItemServiceRetList)
            If iterationSize > 0 Then
                iteratorID = _response.iteratorID
                iterationRemaining = _response.iteratorRemainingCount
            End If
            If iterationSize > 0 Then
                If IsNothing(_response.Detail) Then
                    iterationRemaining = 0
                Else
                    iterationRemaining = 9999
                    'IteratorID = Chr(Asc(IteratorID) + 1)
                    'IteratorID = response.Detail
                End If
            End If
            'Message($"Query {tableName} returned with a status of: {_response.StatusMessage} ({_response.StatusCode})")
            If _response.StatusCode = 500 Then
                'MessageLink("Request File", CurrentSolutionFolder + "\" + _Solution.SolutionName + ".xml", _requestSet.ToXMLString())
            End If
            Return _response.Detail
        Catch ex As Exception
            If iterationSize = 1 Then
                iterationRemaining = 0
                'ErrorMessage($"Error Querying {tableName}: {ex.Message}")
                'MessageLink("Request File", CurrentSolutionFolder + "\" + _Solution.SolutionName + ".xml", _requestSet.ToXMLString())
                'If Not IsNothing(msgSetResponse) Then ErrorMessage(msgSetResponse.ToXMLString())
            End If
            Return Nothing
        End Try

        'http://webcache.googleusercontent.com/search?q=cache:Kt6QrI8f9JsJ:developer.intuit.com/qbSDK-current/samples/qbdt/vb/qbfcC:\Users\Peter\Documents\Clients\A Test\DataExtensions/modDataExtSample.bas+qbfc+DataExtRetList&cd=2&hl=en&ct=clnk&gl=us
        ' Add the DataExtDefQuery request
        'Dim dataExtDefQuery As IDataExtDefQuery
        'QQuery = _requestSet.AppendDataExtDefQueryRq()

        ' set the OwnerID for the Data Ext we want returned
        ' specify an ownerID of "0" for custom fields
        'QQuery.ORDataExtDefQuery.OwnerIDList.Add("0")


    End Function

    Public Function ConnectToQB() As Boolean
        Try
            Dim _qbfcVersion As String = String.Empty

            'http://consolibyte.com/wiki/doku.php/quickbooks_desktop_integration
            Dim appID As String = String.Empty

            _sessionManager = New QBSessionManager
            _sessionManager.OpenConnection2(appID, "Rekonator", ENConnectionType.ctLocalQBD)
            _sessionManager.BeginSession(String.Empty, ENOpenMode.omDontCare)

            _msgSetRequest = GetLatestMsgRequestSet()
            'Message("QuickBooks:  " + GetProduct() + " (QBFC Version " + _qbfcVersion + ")")
            Return True
        Catch ex As Exception
            'ErrorMessage($"Error Opening Connection: {ex.Message}, Connection Type {[Enum].GetName(GetType(ENConnectionType), _qbConnectionType)}")
            If ex.InnerException IsNot Nothing Then
                'ErrorMessage(ex.InnerException.ToString)
            End If
            Return False
        End Try
        Return True
    End Function

    Private Function GetLatestMsgRequestSet() As IMsgSetRequest
        Dim supportedVersion As Double = QBFCLatestVersion()

        Dim qbXMLMajorVer As Short = 0
        Dim qbXMLMinorVer As Short = 0

        If supportedVersion >= 14.0 Then
            qbXMLMajorVer = 14
            qbXMLMinorVer = 0
        ElseIf supportedVersion >= 13.0 Then
            qbXMLMajorVer = 13
            qbXMLMinorVer = 0
        ElseIf supportedVersion >= 12.0 Then
            qbXMLMajorVer = 12
            qbXMLMinorVer = 0
        ElseIf supportedVersion >= 11.0 Then
            qbXMLMajorVer = 11
            qbXMLMinorVer = 0
        ElseIf supportedVersion >= 10.0 Then
            qbXMLMajorVer = 10
            qbXMLMinorVer = 0
        ElseIf supportedVersion >= 9.0 Then
            qbXMLMajorVer = 9
            qbXMLMinorVer = 0
        ElseIf supportedVersion >= 8.0 Then
            qbXMLMajorVer = 8
            qbXMLMinorVer = 0
        ElseIf supportedVersion >= 7.0 Then
            qbXMLMajorVer = 7
            qbXMLMinorVer = 0
        ElseIf supportedVersion >= 6.0 Then
            qbXMLMajorVer = 6
            qbXMLMinorVer = 0
        ElseIf supportedVersion >= 5.0 Then
            qbXMLMajorVer = 5
            qbXMLMinorVer = 0
        ElseIf supportedVersion >= 4.0 Then
            qbXMLMajorVer = 4
            qbXMLMinorVer = 0
        ElseIf supportedVersion >= 3.0 Then
            qbXMLMajorVer = 3
            qbXMLMinorVer = 0
        ElseIf supportedVersion >= 2.0 Then
            qbXMLMajorVer = 2
            qbXMLMinorVer = 0
        ElseIf supportedVersion >= 1.1 Then
            qbXMLMajorVer = 1
            qbXMLMinorVer = 1
        Else
            qbXMLMajorVer = 1
            qbXMLMinorVer = 0
            'ErrorMessage("It seems that you are running QuickBooks 2002 Release 1. We strongly recommend that you use QuickBooks' online update feature to obtain the latest fixes and enhancements")
        End If

        _msgSetRequest = _sessionManager.CreateMsgSetRequest("US", qbXMLMajorVer, qbXMLMinorVer)
        Return _msgSetRequest
    End Function

    Private Function QBFCLatestVersion() As Double
        Dim requestSet = _sessionManager.CreateMsgSetRequest("US", 1, 0)
        requestSet.AppendHostQueryRq()
        Dim queryResponse As IMsgSetResponse = _sessionManager.DoRequests(requestSet)

        _response = queryResponse.ResponseList.GetAt(0)
        Dim hostResponse As IHostRet = TryCast(_response.Detail, IHostRet)
        Dim supportedVersions As IBSTRList = TryCast(hostResponse.SupportedQBXMLVersionList, IBSTRList)

        Dim i As Integer
        Dim vers As Double
        Dim lastVers As Double = 0
        Dim svers As String = Nothing

        For i = 0 To supportedVersions.Count - 1
            svers = supportedVersions.GetAt(i)
            vers = Convert.ToDouble(svers)
            If vers > lastVers Then
                lastVers = vers
            End If
        Next
        If lastVers >= 13 Then lastVers = 13
        Return lastVers
    End Function


    Private Sub WalkGeneralDetailReportQueryRs()
        Try
            If _msgSetResponse Is Nothing Then Return
            Dim responseList As IResponseList = _msgSetResponse.ResponseList
            If responseList Is Nothing Then Return

            For i As Integer = 0 To responseList.Count - 1
                Dim response As IResponse = responseList.GetAt(i)

                If response.StatusCode >= 0 Then

                    If response.Detail IsNot Nothing Then
                        Dim responseType As ENResponseType = CType(response.Type.GetValue(), ENResponseType)

                        If responseType = ENResponseType.rtGeneralDetailReportQueryRs Then
                            Dim ReportRet As IReportRet = CType(response.Detail, IReportRet)
                            'GetReportColumns(ReportRet)
                            WalkReportRet(ReportRet)
                        End If
                    End If
                End If
            Next

        Catch ex As Exception

        End Try
    End Sub

    'Private Sub GetReportColumns(ByVal ReportRet As IReportRet)
    '    Try
    '        Dim headerText As String = String.Empty
    '        Dim colDesc As IColDesc

    '        If (Not ReportRet.ColDescList Is Nothing) Then
    '            For index = 0 To ReportRet.ColDescList.Count - 1
    '                colDesc = ReportRet.ColDescList.GetAt(index)
    '                If (Not colDesc Is Nothing) Then
    '                    If (Not colDesc.ColTitleList Is Nothing) Then
    '                        If (colDesc.ColTitleList.Count >= 1) Then
    '                            Dim colTitle As IColTitle
    '                            colTitle = colDesc.ColTitleList.GetAt(0)
    '                            If (Not colTitle Is Nothing) Then
    '                                If (Not colTitle.titleRow Is Nothing) And
    '                                   (Not colTitle.value Is Nothing) Then
    '                                    Dim x = colTitle.titleRow
    '                                    Debug.Print(x)
    '                                    headerText = colTitle.value.GetValue()
    '                                    Dim qbType As String = colDesc.dataType.GetAsString()
    '                                    Select Case qbType
    '                                        Case "STRTYPE"
    '                                            '_resultTable.Columns.Add(headerText, GetType(String))
    '                                        Case "AMTTYPE"
    '                                            '_resultTable.Columns.Add(headerText, GetType(Single))
    '                                        Case "DATETYPE"
    '                                            '_resultTable.Columns.Add(headerText, GetType(Date))
    '                                        Case Else
    '                                            '_resultTable.Columns.Add(headerText, GetType(String))
    '                                    End Select
    '                                End If
    '                            End If
    '                        End If
    '                    End If
    '                End If
    '            Next
    '        End If
    '    Catch ex As Exception

    '    End Try
    'End Sub

    Private Sub WalkReportRet(ByVal ReportRet As IReportRet)
        Try
            If ReportRet Is Nothing Then Return

            If ReportRet.ReportData IsNot Nothing Then

                If ReportRet.ReportData.ORReportDataList IsNot Nothing Then

                    For idx As Integer = 0 To ReportRet.ReportData.ORReportDataList.Count - 1
                        If idx Mod 250 = 0 Then Application.Message(idx.ToString)
                        Dim ORReportData As IORReportData = ReportRet.ReportData.ORReportDataList.GetAt(idx)

                        If ORReportData.DataRow IsNot Nothing Then

                            If ORReportData.DataRow IsNot Nothing Then

                                If ORReportData.DataRow.RowData IsNot Nothing Then
                                    'For Each member As String In GetMemberNames(ORReportData.DataRow.RowData)
                                    '    Debug.Print(member)
                                    'Next
                                End If

                                If ORReportData.DataRow.ColDataList IsNot Nothing Then
                                    Dim rowList As New List(Of Object)

                                    Dim colCount As Integer = ORReportData.DataRow.ColDataList.Count
                                    If colCount <> _fieldCount Then Continue For
                                    For i36 As Integer = 0 To colCount - 1

                                        Dim ColData As IColData = ORReportData.DataRow.ColDataList.GetAt(i36)
                                        Dim dataValue As String = ColData.value.GetValue()
                                        rowList.Add(dataValue)
                                    Next
                                    If Not _sql.InsertRow(rowList) Then
                                        Application.ErrorMessage($"Insert from QB Report failed at: {idx}")
                                        Exit Sub
                                    End If
                                End If
                            End If
                        End If

                        'If ORReportData.TextRow IsNot Nothing Then

                        '    If ORReportData.TextRow IsNot Nothing Then
                        '        'For Each member As String In GetMemberNames(ORReportData.TextRow)
                        '        '    Debug.Print(member)
                        '        'Next
                        '    End If
                        'End If

                        'If ORReportData.SubtotalRow IsNot Nothing Then

                        '    If ORReportData.SubtotalRow IsNot Nothing Then

                        '        If ORReportData.SubtotalRow.RowData IsNot Nothing Then
                        '        End If

                        '        If ORReportData.SubtotalRow.ColDataList IsNot Nothing Then

                        '            For i37 As Integer = 0 To ORReportData.SubtotalRow.ColDataList.Count - 1
                        '                Dim ColData As IColData = ORReportData.SubtotalRow.ColDataList.GetAt(i37)
                        '            Next
                        '        End If
                        '    End If
                        'End If

                        'If ORReportData.TotalRow IsNot Nothing Then

                        '    If ORReportData.TotalRow IsNot Nothing Then

                        '        If ORReportData.TotalRow.RowData IsNot Nothing Then
                        '        End If

                        '        If ORReportData.TotalRow.ColDataList IsNot Nothing Then

                        '            For i38 As Integer = 0 To ORReportData.TotalRow.ColDataList.Count - 1
                        '                Dim ColData As IColData = ORReportData.TotalRow.ColDataList.GetAt(i38)
                        '            Next
                        '        End If
                        '    End If
                        'End If
                    Next
                End If
            End If

        Catch ex As Exception

        End Try
    End Sub

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
