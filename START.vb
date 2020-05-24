Imports System.ComponentModel
Imports System.Threading
Imports PLANDAY.OvertimeForm

Public Class START
    Private WithEvents watcher As New IdleWatcher
    Private closing_enabled As Boolean

    Private Sub START_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            ClosingForm.StartPosition = FormStartPosition.Manual
            ClosingForm.Location = New Point(Me.Location.X + Me.Width / 2 - ClosingForm.Width / 2, Me.Location.Y + Me.Height / 2 - ClosingForm.Height / 2)
            ClosingForm.Show(Me)
            bgw_closer.RunWorkerAsync()
        End If
    End Sub

    Private Sub OVERTIME_button_Click(sender As Object, e As EventArgs) Handles OVERTIME_button.Click
        Me.Hide()
        ''''loader'''
        Dim th As System.Threading.Thread = New Threading.Thread(Sub() Loader(Me.Location.X, Me.Location.Y, Me.Width, Me.Height))
        th.SetApartmentState(ApartmentState.STA)
        th.Start()
        OvertimeForm.Show()
        watcher.Enabled = False
        closing_enabled = False
        Me.Close()
    End Sub

    Private Sub Loader(x As Integer, y As Integer, width As Integer, height As Integer)
        myWorkingForm = New LoadingForm()
        myWorkingForm.StartPosition = FormStartPosition.Manual
        myWorkingForm.Location = New Point(x + width / 2 - myWorkingForm.Width / 2, y + height / 2 - myWorkingForm.Height / 2)
        System.Windows.Forms.Application.Run(myWorkingForm)
    End Sub

    Private Sub Closer(x As Integer, y As Integer, width As Integer, height As Integer)
        myWorkingForm = New ClosingForm()
        myWorkingForm.StartPosition = FormStartPosition.Manual
        myWorkingForm.Location = New Point(x + width / 2 - myWorkingForm.Width / 2, y + height / 2 - myWorkingForm.Height / 2)
        System.Windows.Forms.Application.Run(myWorkingForm)
    End Sub

    Private Sub PAYROLL_button_Click(sender As Object, e As EventArgs) Handles PAYROLL_button.Click
        Me.Hide()
        ''''loader'''
        Dim th As System.Threading.Thread = New Threading.Thread(Sub() Loader(Me.Location.X, Me.Location.Y, Me.Width, Me.Height))
        th.SetApartmentState(ApartmentState.STA)
        th.Start()
        ExportForm.Show()
        watcher.Enabled = False
        closing_enabled = False
        Me.Close()
    End Sub

    Private Sub START_Load(sender As Object, e As EventArgs) Handles Me.Load
        watcher.Timeout = 10000 ' 10 second timeout
        watcher.Enabled = True
        closing_enabled = True
    End Sub

    Private Sub START_MouseMove(sender As Object, e As MouseEventArgs) Handles Me.MouseMove
        watcher.Refresh()
    End Sub

    Private Sub watcher_idle(sender As Object) Handles watcher.Idle
        Me.Invoke(Sub() Me.BringToFront())
        ' This will depend on your implementation
        If closing_enabled Then
            ''''closer'''
            Dim th As System.Threading.Thread = New Threading.Thread(Sub() Closer(Me.Location.X, Me.Location.Y, Me.Width, Me.Height))
            th.SetApartmentState(ApartmentState.STA)
            th.Start()
            System.Threading.Thread.Sleep(5500)
            myWorkingForm.BeginInvoke(Sub() myWorkingForm.Close())
        End If
        Application.Exit()
    End Sub

    Private Sub OVERTIME_button_MouseMove(sender As Object, e As MouseEventArgs) Handles OVERTIME_button.MouseMove
        watcher.Refresh()
    End Sub

    Private Sub PAYROLL_button_MouseMove(sender As Object, e As MouseEventArgs) Handles PAYROLL_button.MouseMove
        watcher.Refresh()
    End Sub

    Private Sub TOOLS_button_MouseMove(sender As Object, e As MouseEventArgs) Handles TOOLS_button.MouseMove
        watcher.Refresh()
    End Sub

    Private Sub bgw_closer_DoWork(sender As Object, e As DoWorkEventArgs) Handles bgw_closer.DoWork
        System.Threading.Thread.Sleep(5000)
        bgw_closer.ReportProgress(100)
    End Sub

    Private Sub bgw_closer_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles bgw_closer.RunWorkerCompleted
        Application.Exit()
    End Sub

End Class