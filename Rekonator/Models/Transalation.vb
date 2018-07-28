Public Class Translation
    Property TranslationName As String
    Property TableName As String

    Public Shared Translations As New List(Of Translation)
    Public Shared Sub Add(translationName As String, tableName As String)
        Translations.Add(New Translation With {.TranslationName = translationName, .TableName = tableName})
    End Sub
    Public Shared Function GetDataSource(translationName As String) As Translation
        Return Translations.Where(Function(w) w.TranslationName = translationName).FirstOrDefault
    End Function

    'Select Case dataSourceName
    'Case "QuickBooks"
    '            Translation.Add("Items", "Item")
    '            Translation.Add("Purchase Orders", "PurchaseOrder")
    '            Translation.Add("Invoices", "Invoice")
    '            Translation.Add("Payments", "Payment")
    '            Translation.Add("P/L Report", "PLReport")
    '            Translation.Add("AR Summary", "AR Summary Report")
    '            Translation.Add("AR Detail", "AR Detail Report")
    '    End Select
    'End Sub
End Class
