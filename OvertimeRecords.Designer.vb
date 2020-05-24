<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class OvertimeRecords
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
        Me.KW_Records = New System.Windows.Forms.ListView()
        Me.select_employee = New System.Windows.Forms.ComboBox()
        Me.SuspendLayout()
        '
        'KW_Records
        '
        Me.KW_Records.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.KW_Records.HideSelection = False
        Me.KW_Records.Location = New System.Drawing.Point(12, 58)
        Me.KW_Records.Name = "KW_Records"
        Me.KW_Records.Size = New System.Drawing.Size(924, 393)
        Me.KW_Records.TabIndex = 0
        Me.KW_Records.TabStop = False
        Me.KW_Records.UseCompatibleStateImageBehavior = False
        Me.KW_Records.View = System.Windows.Forms.View.Details
        '
        'select_employee
        '
        Me.select_employee.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.select_employee.FormattingEnabled = True
        Me.select_employee.IntegralHeight = False
        Me.select_employee.Location = New System.Drawing.Point(13, 14)
        Me.select_employee.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.select_employee.Name = "select_employee"
        Me.select_employee.Size = New System.Drawing.Size(286, 28)
        Me.select_employee.TabIndex = 1
        '
        'OvertimeRecords
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(948, 463)
        Me.Controls.Add(Me.select_employee)
        Me.Controls.Add(Me.KW_Records)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "OvertimeRecords"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "KW Stunden"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents KW_Records As ListView
    Friend WithEvents select_employee As ComboBox
End Class
