<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ExportForm
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ExportForm))
        Me.bgw_export_ALL = New System.ComponentModel.BackgroundWorker()
        Me.bgw_export_active = New System.ComponentModel.BackgroundWorker()
        Me.bgw_export_deactivated = New System.ComponentModel.BackgroundWorker()
        Me.CLOSE_button = New System.Windows.Forms.Button()
        Me.START_button = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.select_month = New System.Windows.Forms.ComboBox()
        Me.select_status = New System.Windows.Forms.ComboBox()
        Me.select_hotel = New System.Windows.Forms.ListView()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label_folder = New System.Windows.Forms.Label()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.YEAR_button = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'bgw_export_ALL
        '
        Me.bgw_export_ALL.WorkerReportsProgress = True
        '
        'bgw_export_active
        '
        Me.bgw_export_active.WorkerReportsProgress = True
        '
        'bgw_export_deactivated
        '
        Me.bgw_export_deactivated.WorkerReportsProgress = True
        '
        'CLOSE_button
        '
        Me.CLOSE_button.Cursor = System.Windows.Forms.Cursors.Hand
        Me.CLOSE_button.Location = New System.Drawing.Point(260, 125)
        Me.CLOSE_button.Name = "CLOSE_button"
        Me.CLOSE_button.Size = New System.Drawing.Size(136, 49)
        Me.CLOSE_button.TabIndex = 4
        Me.CLOSE_button.Text = "Close"
        Me.ToolTip1.SetToolTip(Me.CLOSE_button, "bye bye :-)")
        Me.CLOSE_button.UseVisualStyleBackColor = True
        '
        'START_button
        '
        Me.START_button.Cursor = System.Windows.Forms.Cursors.Hand
        Me.START_button.Location = New System.Drawing.Point(260, 70)
        Me.START_button.Name = "START_button"
        Me.START_button.Size = New System.Drawing.Size(136, 49)
        Me.START_button.TabIndex = 3
        Me.START_button.Text = "Export"
        Me.ToolTip1.SetToolTip(Me.START_button, "Do work")
        Me.START_button.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(12, 23)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(64, 20)
        Me.Label1.TabIndex = 15
        Me.Label1.Text = "Month:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(12, 86)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(101, 20)
        Me.Label2.TabIndex = 16
        Me.Label2.Text = "Employees:"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(12, 211)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(57, 20)
        Me.Label3.TabIndex = 17
        Me.Label3.Text = "Hotel:"
        '
        'select_month
        '
        Me.select_month.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.select_month.FormattingEnabled = True
        Me.select_month.IntegralHeight = False
        Me.select_month.Location = New System.Drawing.Point(16, 48)
        Me.select_month.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.select_month.Name = "select_month"
        Me.select_month.Size = New System.Drawing.Size(222, 28)
        Me.select_month.TabIndex = 0
        '
        'select_status
        '
        Me.select_status.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.select_status.FormattingEnabled = True
        Me.select_status.IntegralHeight = False
        Me.select_status.Location = New System.Drawing.Point(16, 111)
        Me.select_status.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.select_status.Name = "select_status"
        Me.select_status.Size = New System.Drawing.Size(222, 28)
        Me.select_status.TabIndex = 1
        '
        'select_hotel
        '
        Me.select_hotel.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.select_hotel.FullRowSelect = True
        Me.select_hotel.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable
        Me.select_hotel.HideSelection = False
        Me.select_hotel.Location = New System.Drawing.Point(16, 232)
        Me.select_hotel.Name = "select_hotel"
        Me.select_hotel.Size = New System.Drawing.Size(380, 366)
        Me.select_hotel.TabIndex = 20
        Me.select_hotel.TabStop = False
        Me.select_hotel.UseCompatibleStateImageBehavior = False
        Me.select_hotel.View = System.Windows.Forms.View.Details
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(12, 152)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(117, 20)
        Me.Label4.TabIndex = 21
        Me.Label4.Text = "Export folder:"
        '
        'Label_folder
        '
        Me.Label_folder.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Label_folder.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label_folder.ForeColor = System.Drawing.SystemColors.Highlight
        Me.Label_folder.Location = New System.Drawing.Point(11, 177)
        Me.Label_folder.Name = "Label_folder"
        Me.Label_folder.Size = New System.Drawing.Size(384, 28)
        Me.Label_folder.TabIndex = 22
        Me.Label_folder.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'YEAR_button
        '
        Me.YEAR_button.Cursor = System.Windows.Forms.Cursors.Hand
        Me.YEAR_button.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.YEAR_button.ForeColor = System.Drawing.SystemColors.Highlight
        Me.YEAR_button.Location = New System.Drawing.Point(260, 15)
        Me.YEAR_button.Name = "YEAR_button"
        Me.YEAR_button.Size = New System.Drawing.Size(136, 49)
        Me.YEAR_button.TabIndex = 2
        Me.YEAR_button.Text = "2020"
        Me.ToolTip1.SetToolTip(Me.YEAR_button, "Change year")
        Me.YEAR_button.UseVisualStyleBackColor = True
        '
        'ExportForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(144.0!, 144.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.ClientSize = New System.Drawing.Size(410, 612)
        Me.Controls.Add(Me.YEAR_button)
        Me.Controls.Add(Me.Label_folder)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.select_hotel)
        Me.Controls.Add(Me.select_status)
        Me.Controls.Add(Me.select_month)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.START_button)
        Me.Controls.Add(Me.CLOSE_button)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "ExportForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Export"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents bgw_export_ALL As System.ComponentModel.BackgroundWorker
    Friend WithEvents bgw_export_active As System.ComponentModel.BackgroundWorker
    Friend WithEvents bgw_export_deactivated As System.ComponentModel.BackgroundWorker
    Friend WithEvents CLOSE_button As Button
    Friend WithEvents START_button As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents select_month As ComboBox
    Friend WithEvents select_status As ComboBox
    Friend WithEvents select_hotel As ListView
    Friend WithEvents Label4 As Label
    Friend WithEvents Label_folder As Label
    Friend WithEvents ToolTip1 As ToolTip
    Friend WithEvents YEAR_button As Button
End Class
