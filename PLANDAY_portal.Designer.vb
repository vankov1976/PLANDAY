<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class PLANDAY_portal
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(PLANDAY_portal))
        Me.DE_button = New System.Windows.Forms.Button()
        Me.HQ_button = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'DE_button
        '
        Me.DE_button.Cursor = System.Windows.Forms.Cursors.Hand
        Me.DE_button.FlatAppearance.BorderSize = 5
        Me.DE_button.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DE_button.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.DE_button.Location = New System.Drawing.Point(-3, 0)
        Me.DE_button.Margin = New System.Windows.Forms.Padding(4)
        Me.DE_button.Name = "DE_button"
        Me.DE_button.Size = New System.Drawing.Size(423, 90)
        Me.DE_button.TabIndex = 2
        Me.DE_button.TabStop = False
        Me.DE_button.Text = "MEININGER Germany"
        Me.DE_button.UseVisualStyleBackColor = True
        '
        'HQ_button
        '
        Me.HQ_button.Cursor = System.Windows.Forms.Cursors.Hand
        Me.HQ_button.FlatAppearance.BorderSize = 5
        Me.HQ_button.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.HQ_button.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.HQ_button.Location = New System.Drawing.Point(-3, 89)
        Me.HQ_button.Margin = New System.Windows.Forms.Padding(4)
        Me.HQ_button.Name = "HQ_button"
        Me.HQ_button.Size = New System.Drawing.Size(423, 90)
        Me.HQ_button.TabIndex = 3
        Me.HQ_button.TabStop = False
        Me.HQ_button.Text = "MEININGER HQ"
        Me.HQ_button.UseVisualStyleBackColor = True
        '
        'PLANDAY_portal
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(192.0!, 192.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.AutoValidate = System.Windows.Forms.AutoValidate.EnablePreventFocusChange
        Me.ClientSize = New System.Drawing.Size(416, 175)
        Me.ControlBox = False
        Me.Controls.Add(Me.HQ_button)
        Me.Controls.Add(Me.DE_button)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "PLANDAY_portal"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.TopMost = True
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents DE_button As Button
    Friend WithEvents HQ_button As Button
End Class
