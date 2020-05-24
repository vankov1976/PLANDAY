<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class ExportPDF
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
        Me.CREATE_button = New System.Windows.Forms.Button()
        Me.CLOSE_button = New System.Windows.Forms.Button()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label_folder = New System.Windows.Forms.Label()
        Me.Columns_list = New System.Windows.Forms.ListBox()
        Me.SuspendLayout()
        '
        'CREATE_button
        '
        Me.CREATE_button.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CREATE_button.Cursor = System.Windows.Forms.Cursors.Hand
        Me.CREATE_button.Location = New System.Drawing.Point(12, 12)
        Me.CREATE_button.Name = "CREATE_button"
        Me.CREATE_button.Size = New System.Drawing.Size(136, 49)
        Me.CREATE_button.TabIndex = 11
        Me.CREATE_button.Text = "Erstellen"
        Me.ToolTip1.SetToolTip(Me.CREATE_button, "PDF Datei erstellen und speichern")
        Me.CREATE_button.UseVisualStyleBackColor = True
        '
        'CLOSE_button
        '
        Me.CLOSE_button.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CLOSE_button.Cursor = System.Windows.Forms.Cursors.Hand
        Me.CLOSE_button.Location = New System.Drawing.Point(156, 12)
        Me.CLOSE_button.Name = "CLOSE_button"
        Me.CLOSE_button.Size = New System.Drawing.Size(136, 49)
        Me.CLOSE_button.TabIndex = 12
        Me.CLOSE_button.Text = "Beenden"
        Me.ToolTip1.SetToolTip(Me.CLOSE_button, "bye bye :-)")
        Me.CLOSE_button.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(12, 71)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(107, 20)
        Me.Label1.TabIndex = 14
        Me.Label1.Text = "Speicherort:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(12, 121)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(76, 20)
        Me.Label2.TabIndex = 15
        Me.Label2.Text = "Spalten:"
        '
        'Label_folder
        '
        Me.Label_folder.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Label_folder.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label_folder.ForeColor = System.Drawing.SystemColors.Highlight
        Me.Label_folder.Location = New System.Drawing.Point(12, 97)
        Me.Label_folder.Name = "Label_folder"
        Me.Label_folder.Size = New System.Drawing.Size(280, 20)
        Me.Label_folder.TabIndex = 16
        '
        'Columns_list
        '
        Me.Columns_list.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Columns_list.FormattingEnabled = True
        Me.Columns_list.ItemHeight = 20
        Me.Columns_list.Location = New System.Drawing.Point(12, 148)
        Me.Columns_list.Name = "Columns_list"
        Me.Columns_list.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        Me.Columns_list.Size = New System.Drawing.Size(280, 284)
        Me.Columns_list.TabIndex = 17
        '
        'ExportPDF
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(304, 446)
        Me.Controls.Add(Me.Columns_list)
        Me.Controls.Add(Me.Label_folder)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.CLOSE_button)
        Me.Controls.Add(Me.CREATE_button)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "ExportPDF"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "ExportPDF"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents CREATE_button As Button
    Friend WithEvents ToolTip1 As ToolTip
    Friend WithEvents CLOSE_button As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label_folder As Label
    Friend WithEvents Columns_list As ListBox
End Class
