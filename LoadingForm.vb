Public Class LoadingForm
    Private Sub Loading_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.BackColor = Color.Blue
        Me.TransparencyKey = Color.Blue
        If (Not Owner Is Nothing) Then
            Location = New Point(Owner.Location.X + Owner.Width / 2 - Width / 2, Owner.Location.Y + Owner.Height / 2 - Height / 2 + 100)
        End If
    End Sub
End Class