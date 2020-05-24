<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class START
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(START))
        Me.OVERTIME_button = New System.Windows.Forms.Button()
        Me.TOOLS_button = New System.Windows.Forms.Button()
        Me.PAYROLL_button = New System.Windows.Forms.Button()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.bgw_closer = New System.ComponentModel.BackgroundWorker()
        Me.SuspendLayout()
        '
        'OVERTIME_button
        '
        Me.OVERTIME_button.Cursor = System.Windows.Forms.Cursors.Hand
        Me.OVERTIME_button.FlatAppearance.BorderSize = 5
        Me.OVERTIME_button.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.OVERTIME_button.Image = CType(resources.GetObject("OVERTIME_button.Image"), System.Drawing.Image)
        Me.OVERTIME_button.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.OVERTIME_button.Location = New System.Drawing.Point(0, 0)
        Me.OVERTIME_button.Name = "OVERTIME_button"
        Me.OVERTIME_button.Size = New System.Drawing.Size(317, 69)
        Me.OVERTIME_button.TabIndex = 1
        Me.OVERTIME_button.TabStop = False
        Me.OVERTIME_button.Text = "Overtime"
        Me.ToolTip1.SetToolTip(Me.OVERTIME_button, "Calculate and adjust employee overtime accounts")
        Me.OVERTIME_button.UseVisualStyleBackColor = True
        '
        'TOOLS_button
        '
        Me.TOOLS_button.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.TOOLS_button.Cursor = System.Windows.Forms.Cursors.Hand
        Me.TOOLS_button.FlatAppearance.BorderSize = 5
        Me.TOOLS_button.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TOOLS_button.Image = CType(resources.GetObject("TOOLS_button.Image"), System.Drawing.Image)
        Me.TOOLS_button.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.TOOLS_button.Location = New System.Drawing.Point(0, 136)
        Me.TOOLS_button.Name = "TOOLS_button"
        Me.TOOLS_button.Size = New System.Drawing.Size(317, 69)
        Me.TOOLS_button.TabIndex = 3
        Me.TOOLS_button.TabStop = False
        Me.TOOLS_button.Text = "Tools"
        Me.ToolTip1.SetToolTip(Me.TOOLS_button, "Not implemented yet :-)")
        Me.TOOLS_button.UseVisualStyleBackColor = True
        '
        'PAYROLL_button
        '
        Me.PAYROLL_button.Cursor = System.Windows.Forms.Cursors.Hand
        Me.PAYROLL_button.FlatAppearance.BorderSize = 5
        Me.PAYROLL_button.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PAYROLL_button.Image = CType(resources.GetObject("PAYROLL_button.Image"), System.Drawing.Image)
        Me.PAYROLL_button.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.PAYROLL_button.Location = New System.Drawing.Point(0, 68)
        Me.PAYROLL_button.Name = "PAYROLL_button"
        Me.PAYROLL_button.Size = New System.Drawing.Size(317, 69)
        Me.PAYROLL_button.TabIndex = 2
        Me.PAYROLL_button.TabStop = False
        Me.PAYROLL_button.Text = "Payroll Export"
        Me.ToolTip1.SetToolTip(Me.PAYROLL_button, "Create Excel payroll reports")
        Me.PAYROLL_button.UseVisualStyleBackColor = True
        '
        'bgw_closer
        '
        Me.bgw_closer.WorkerReportsProgress = True
        '
        'START
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(144.0!, 144.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.AutoValidate = System.Windows.Forms.AutoValidate.EnablePreventFocusChange
        Me.ClientSize = New System.Drawing.Size(316, 204)
        Me.ControlBox = False
        Me.Controls.Add(Me.PAYROLL_button)
        Me.Controls.Add(Me.TOOLS_button)
        Me.Controls.Add(Me.OVERTIME_button)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "START"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.TopMost = True
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents OVERTIME_button As Button
    Friend WithEvents TOOLS_button As Button
    Friend WithEvents PAYROLL_button As Button
    Friend WithEvents ToolTip1 As ToolTip
    Friend WithEvents bgw_closer As System.ComponentModel.BackgroundWorker
End Class
