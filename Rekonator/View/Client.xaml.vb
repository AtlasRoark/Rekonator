


Imports System.Threading
Imports System.Windows.Threading

Partial Class Client

    Private _dispatcherTimer As DispatcherTimer
    Private _uiDispatcher As Dispatcher
    Private _value As Integer = 0
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

    End Sub

    Private Sub BtUpload_Click(sender As Object, e As RoutedEventArgs)
        BtCancel.Visibility = Visibility.Hidden
        BtUpload.Visibility = Visibility.Hidden
        PbUpload.Visibility = Visibility.Visible
        _uiDispatcher = Application.Current.Dispatcher
        Task.Factory.StartNew(Sub() Thread.Sleep(2000)).ContinueWith(Sub(a) LaunchWeb())


    End Sub

    Private Sub BtCancel_Click(sender As Object, e As RoutedEventArgs)
        Dim w As New MainWindow
        w.Show()
        Me.Close()
    End Sub

    ''' <summary>
    ''' Start the timer.
    ''' </summary>
    Sub Start()
        _value = 0
        _dispatcherTimer = New DispatcherTimer
        AddHandler _dispatcherTimer.Tick, AddressOf dispatcherTimer_Tick
        _dispatcherTimer.Interval = New TimeSpan(0, 0, 0, 500)
        _dispatcherTimer.Start()
    End Sub



    Private Sub dispatcherTimer_Tick(ByVal sender As Object, ByVal e As EventArgs)
        _value += 25
        PbUpload.Value += 25
        If _value > 1000 Then
            _dispatcherTimer.Stop()
            _dispatcherTimer = Nothing
            Debug.Print("hi")
            Debug.Print(_uiDispatcher.Thread.ManagedThreadId.ToString)

            CommandManager.InvalidateRequerySuggested()
            _uiDispatcher.Invoke(Sub() LaunchWeb())


        End If
        ' Forcing the CommandManager to raise the RequerySuggested event
        CommandManager.InvalidateRequerySuggested()
    End Sub

    Private Sub LaunchWeb()
        _uiDispatcher.BeginInvoke(Sub()
                                      BtCancel.Visibility = Visibility.Visible
                                      BtUpload.Visibility = Visibility.Visible
                                      PbUpload.Visibility = Visibility.Hidden
                                      Dim webAddress As String = "http://localhost:53780/#/Report"
                                      Process.Start(webAddress)
                                  End Sub)

    End Sub




End Class
