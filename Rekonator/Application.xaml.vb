Class Application

    ' Application-level events, such as Startup, Exit, and DispatcherUnhandledException
    ' can be handled in this file.
    Delegate Sub MessageDelegate(messageText As String, isError As Boolean)
    Public Shared Property ConnectionString As String = "Data Source=localhost;Initial Catalog=Rekonator;User ID=sa;Password=Summ!t29"
    Public Shared MessageFunc As MessageDelegate

    Public Shared Sub Message(messageText)
        Application.Current.Dispatcher.BeginInvoke(Sub() MessageFunc.Invoke(messageText, False))
    End Sub
    Public Shared Sub ErrorMessage(messageText)
        Application.Current.Dispatcher.BeginInvoke(Sub() MessageFunc.Invoke(messageText, True))
    End Sub
End Class
