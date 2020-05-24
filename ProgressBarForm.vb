
Public Class ProgressBarForm
    ' The delegate
    Delegate Sub SetLabelText_Delegate(ByVal Label As Label, ByVal text As String, ByVal number As Long)

    ' The delegates subroutine.
    Private Sub SetLabelText_ThreadSafe(ByVal Label As Label, ByVal text As String, ByVal number As Long)
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.
        ' If these threads are different, it returns true.
        If Label.InvokeRequired Then
            Dim MyDelegate As New SetLabelText_Delegate(AddressOf SetLabelText_ThreadSafe)
            Me.Invoke(MyDelegate, New Object() {Label, text})
        Else
            Label.Text = text
            ActionNumber = number
        End If
    End Sub

    Public TotalCount As Long
    Private ActionNumber As Long

    Protected Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim param As CreateParams = MyBase.CreateParams
            param.ClassStyle = param.ClassStyle Or &H200
            Return param
        End Get
    End Property
    Delegate Sub PStep(ByVal value As Integer)
    Public Sub ProgressStep(ByVal value As Integer)

        If Me.InvokeRequired Then

            Me.Invoke(New PStep(AddressOf ProgressStep))
        Else

            Me.ProgressBar1.PerformStep()
        End If
    End Sub

    Private Sub ProgressBarForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        With ProgressBar1
            .Style = ProgressBarStyle.Continuous
            .Step = 1
            .Minimum = 0
            .Maximum = TotalCount
            .Value = 0
        End With
        Label1.Text = "0 %"
        Label2.Text = "Starting..."
    End Sub

    Public Sub NextAction(ByVal number As Long, ByVal name As String)
        ProgressStep(ActionNumber)
        SetLabelText_ThreadSafe(Me.Label1, 100 * Math.Round((ProgressBar1.Value / ProgressBar1.Maximum), 2, MidpointRounding.AwayFromZero) & " %", number)
        SetLabelText_ThreadSafe(Me.Label2, "Action " & ActionNumber & " of " & ProgressBar1.Maximum & " | " & name, number)
    End Sub

    Public Sub Complete()
        Me.Label2.Text = "Complete | This Windows will close in 3 seconds"
        Me.Refresh()
        Threading.Thread.Sleep(1000)
        Me.Label2.Text = "Complete | This Windows will close in 2 seconds"
        Me.Refresh()
        Threading.Thread.Sleep(1000)
        Me.Label2.Text = "Complete | This Windows will close in 1 second"
        Me.Refresh()
        Threading.Thread.Sleep(1000)
        Me.Close()
    End Sub

End Class