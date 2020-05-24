<System.Diagnostics.DebuggerStepThrough()>
Public Class TransparentPicturebox
    Inherits PictureBox
    Public Sub New()
        Me.SetStyle(ControlStyles.Opaque, True)
        Me.SetStyle(ControlStyles.OptimizedDoubleBuffer, False)
    End Sub
    Protected Overrides ReadOnly Property CreateParams() As System.Windows.Forms.CreateParams
        Get
            Dim cp As CreateParams = MyBase.CreateParams
            cp.ExStyle = cp.ExStyle Or &H20  ' Turn on WS_EX_TRANSPARENT
            Return cp
        End Get
    End Property
End Class