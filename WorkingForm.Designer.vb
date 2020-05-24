<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class WorkingForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(WorkingForm))
        Me.TransparentPicturebox1 = New PLANDAY.TransparentPicturebox()
        CType(Me.TransparentPicturebox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TransparentPicturebox1
        '
        Me.TransparentPicturebox1.BackColor = System.Drawing.Color.Transparent
        Me.TransparentPicturebox1.Image = CType(resources.GetObject("TransparentPicturebox1.Image"), System.Drawing.Image)
        Me.TransparentPicturebox1.Location = New System.Drawing.Point(0, 0)
        Me.TransparentPicturebox1.Name = "TransparentPicturebox1"
        Me.TransparentPicturebox1.Size = New System.Drawing.Size(156, 155)
        Me.TransparentPicturebox1.TabIndex = 0
        Me.TransparentPicturebox1.TabStop = False
        '
        'WorkingForm
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Inherit
        Me.ClientSize = New System.Drawing.Size(157, 157)
        Me.ControlBox = False
        Me.Controls.Add(Me.TransparentPicturebox1)
        Me.DoubleBuffered = True
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "WorkingForm"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "workingForm"
        CType(Me.TransparentPicturebox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents TransparentPicturebox1 As TransparentPicturebox
End Class
