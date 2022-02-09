Imports System.ComponentModel
Imports System.Threading
Imports PLANDAY.OvertimeForm
Public Class PLANDAY_portal
    Private Sub DE_button_Click(sender As Object, e As EventArgs) Handles DE_button.Click
        OvertimeForm.PLANDAY_portal = "DE"
        START.closing_enabled = False
        Me.Close()
    End Sub

    Private Sub HQ_button_Click(sender As Object, e As EventArgs) Handles HQ_button.Click
        OvertimeForm.PLANDAY_portal = "HQ"
        START.closing_enabled = False
        Me.Close()
    End Sub

    Private Sub PLANDAY_portal_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            ''''closer'''
            Me.Hide()
            Dim th As System.Threading.Thread = New Threading.Thread(Sub() START.Closer(Me.Location.X, Me.Location.Y, Me.Width, Me.Height))
            th.SetApartmentState(ApartmentState.STA)
            th.Start()
            System.Threading.Thread.Sleep(5500)
            myWorkingForm.BeginInvoke(Sub() myWorkingForm.Close())
            Application.Exit()
        End If
    End Sub

    Private Sub PLANDAY_portal_MouseMove(sender As Object, e As MouseEventArgs) Handles Me.MouseMove
        START.watcher.Refresh()
    End Sub

    Private Sub DE_button_MouseMove(sender As Object, e As MouseEventArgs) Handles DE_button.MouseMove
        START.watcher.Refresh()
    End Sub

    Private Sub HQ_button_MouseMove(sender As Object, e As MouseEventArgs) Handles HQ_button.MouseMove
        START.watcher.Refresh()
    End Sub

End Class