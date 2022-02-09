<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ClosingForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ClosingForm))
        Me.TransparentPicturebox1 = New PLANDAY.TransparentPicturebox()
        CType(Me.TransparentPicturebox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TransparentPicturebox1
        '
        Me.TransparentPicturebox1.Image = CType(resources.GetObject("TransparentPicturebox1.Image"), System.Drawing.Image)
        Me.TransparentPicturebox1.Location = New System.Drawing.Point(0, 0)
        Me.TransparentPicturebox1.Margin = New System.Windows.Forms.Padding(4)
        Me.TransparentPicturebox1.Name = "TransparentPicturebox1"
        Me.TransparentPicturebox1.Size = New System.Drawing.Size(423, 272)
        Me.TransparentPicturebox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.TransparentPicturebox1.TabIndex = 0
        Me.TransparentPicturebox1.TabStop = False
        '
        'ClosingForm
        '
        Me.AccessibleRole = System.Windows.Forms.AccessibleRole.None
        Me.AutoScaleDimensions = New System.Drawing.SizeF(192.0!, 192.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.AutoValidate = System.Windows.Forms.AutoValidate.Disable
        Me.CausesValidation = False
        Me.ClientSize = New System.Drawing.Size(426, 273)
        Me.ControlBox = False
        Me.Controls.Add(Me.TransparentPicturebox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "ClosingForm"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Closing..."
        Me.TopMost = True
        CType(Me.TransparentPicturebox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents TransparentPicturebox1 As TransparentPicturebox
End Class
