<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class OvertimeForm
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(OvertimeForm))
        Me.select_employees = New System.Windows.Forms.ListView()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.PDF_button = New System.Windows.Forms.Button()
        Me.EXCEL_button = New System.Windows.Forms.Button()
        Me.DELETE_button = New System.Windows.Forms.Button()
        Me.WRITE_button = New System.Windows.Forms.Button()
        Me.FT_2 = New System.Windows.Forms.Label()
        Me.FT_1 = New System.Windows.Forms.Label()
        Me.FT_label = New System.Windows.Forms.Label()
        Me.loaded = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.shiftsApproved = New System.Windows.Forms.CheckBox()
        Me.active_employees = New System.Windows.Forms.CheckBox()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.select_kw = New System.Windows.Forms.ComboBox()
        Me.select_hotel = New System.Windows.Forms.ComboBox()
        Me.LOAD_button = New System.Windows.Forms.Button()
        Me.CLOSE_button = New System.Windows.Forms.Button()
        Me.bgw = New System.ComponentModel.BackgroundWorker()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.bgw_loading = New System.ComponentModel.BackgroundWorker()
        Me.bgw_write_overtime = New System.ComponentModel.BackgroundWorker()
        Me.ToolTip2 = New System.Windows.Forms.ToolTip(Me.components)
        Me.MODUS_button = New System.Windows.Forms.CheckBox()
        Me.Frame1.SuspendLayout()
        Me.SuspendLayout()
        '
        'select_employees
        '
        Me.select_employees.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.select_employees.BackColor = System.Drawing.SystemColors.Window
        Me.select_employees.CheckBoxes = True
        Me.select_employees.HideSelection = False
        Me.select_employees.Location = New System.Drawing.Point(12, 163)
        Me.select_employees.Name = "select_employees"
        Me.select_employees.ShowItemToolTips = True
        Me.select_employees.Size = New System.Drawing.Size(1022, 392)
        Me.select_employees.TabIndex = 0
        Me.select_employees.TabStop = False
        Me.select_employees.UseCompatibleStateImageBehavior = False
        Me.select_employees.View = System.Windows.Forms.View.Details
        '
        'Frame1
        '
        Me.Frame1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Frame1.Controls.Add(Me.PDF_button)
        Me.Frame1.Controls.Add(Me.EXCEL_button)
        Me.Frame1.Controls.Add(Me.DELETE_button)
        Me.Frame1.Controls.Add(Me.WRITE_button)
        Me.Frame1.Controls.Add(Me.FT_2)
        Me.Frame1.Controls.Add(Me.FT_1)
        Me.Frame1.Controls.Add(Me.FT_label)
        Me.Frame1.Controls.Add(Me.loaded)
        Me.Frame1.Controls.Add(Me.Label6)
        Me.Frame1.Controls.Add(Me.shiftsApproved)
        Me.Frame1.Controls.Add(Me.active_employees)
        Me.Frame1.Controls.Add(Me.CheckBox1)
        Me.Frame1.Controls.Add(Me.Label5)
        Me.Frame1.Controls.Add(Me.Label4)
        Me.Frame1.Controls.Add(Me.select_kw)
        Me.Frame1.Controls.Add(Me.select_hotel)
        Me.Frame1.Location = New System.Drawing.Point(12, -2)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Size = New System.Drawing.Size(882, 158)
        Me.Frame1.TabIndex = 1
        Me.Frame1.TabStop = False
        '
        'PDF_button
        '
        Me.PDF_button.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.PDF_button.BackgroundImage = CType(resources.GetObject("PDF_button.BackgroundImage"), System.Drawing.Image)
        Me.PDF_button.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.PDF_button.Cursor = System.Windows.Forms.Cursors.Hand
        Me.PDF_button.Location = New System.Drawing.Point(737, 85)
        Me.PDF_button.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.PDF_button.Name = "PDF_button"
        Me.PDF_button.Size = New System.Drawing.Size(68, 68)
        Me.PDF_button.TabIndex = 6
        Me.ToolTip2.SetToolTip(Me.PDF_button, "PDF Export")
        Me.PDF_button.UseVisualStyleBackColor = True
        '
        'EXCEL_button
        '
        Me.EXCEL_button.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.EXCEL_button.BackgroundImage = CType(resources.GetObject("EXCEL_button.BackgroundImage"), System.Drawing.Image)
        Me.EXCEL_button.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.EXCEL_button.Cursor = System.Windows.Forms.Cursors.Hand
        Me.EXCEL_button.Location = New System.Drawing.Point(737, 14)
        Me.EXCEL_button.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.EXCEL_button.Name = "EXCEL_button"
        Me.EXCEL_button.Size = New System.Drawing.Size(68, 68)
        Me.EXCEL_button.TabIndex = 5
        Me.ToolTip2.SetToolTip(Me.EXCEL_button, "EXCEL Export")
        Me.EXCEL_button.UseVisualStyleBackColor = True
        '
        'DELETE_button
        '
        Me.DELETE_button.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DELETE_button.BackgroundImage = CType(resources.GetObject("DELETE_button.BackgroundImage"), System.Drawing.Image)
        Me.DELETE_button.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.DELETE_button.Cursor = System.Windows.Forms.Cursors.Hand
        Me.DELETE_button.Location = New System.Drawing.Point(808, 85)
        Me.DELETE_button.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.DELETE_button.Name = "DELETE_button"
        Me.DELETE_button.Size = New System.Drawing.Size(68, 68)
        Me.DELETE_button.TabIndex = 8
        Me.ToolTip2.SetToolTip(Me.DELETE_button, "Die schon übertragenen Überstunden stornieren")
        Me.DELETE_button.UseVisualStyleBackColor = True
        Me.DELETE_button.Visible = False
        '
        'WRITE_button
        '
        Me.WRITE_button.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.WRITE_button.BackgroundImage = CType(resources.GetObject("WRITE_button.BackgroundImage"), System.Drawing.Image)
        Me.WRITE_button.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.WRITE_button.Cursor = System.Windows.Forms.Cursors.Hand
        Me.WRITE_button.Location = New System.Drawing.Point(808, 14)
        Me.WRITE_button.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.WRITE_button.Name = "WRITE_button"
        Me.WRITE_button.Size = New System.Drawing.Size(68, 68)
        Me.WRITE_button.TabIndex = 7
        Me.ToolTip2.SetToolTip(Me.WRITE_button, "Überstunden in PLANDAY übertragen")
        Me.WRITE_button.UseVisualStyleBackColor = True
        Me.WRITE_button.Visible = False
        '
        'FT_2
        '
        Me.FT_2.AutoSize = True
        Me.FT_2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FT_2.ForeColor = System.Drawing.Color.Red
        Me.FT_2.Location = New System.Drawing.Point(594, 129)
        Me.FT_2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.FT_2.Name = "FT_2"
        Me.FT_2.Size = New System.Drawing.Size(0, 20)
        Me.FT_2.TabIndex = 11
        '
        'FT_1
        '
        Me.FT_1.AutoSize = True
        Me.FT_1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FT_1.ForeColor = System.Drawing.Color.Red
        Me.FT_1.Location = New System.Drawing.Point(489, 130)
        Me.FT_1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.FT_1.Name = "FT_1"
        Me.FT_1.Size = New System.Drawing.Size(0, 20)
        Me.FT_1.TabIndex = 10
        '
        'FT_label
        '
        Me.FT_label.AutoSize = True
        Me.FT_label.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FT_label.Location = New System.Drawing.Point(448, 130)
        Me.FT_label.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.FT_label.Name = "FT_label"
        Me.FT_label.Size = New System.Drawing.Size(37, 20)
        Me.FT_label.TabIndex = 9
        Me.FT_label.Text = "FT:"
        Me.FT_label.Visible = False
        '
        'loaded
        '
        Me.loaded.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.loaded.ForeColor = System.Drawing.Color.Red
        Me.loaded.Location = New System.Drawing.Point(378, 130)
        Me.loaded.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.loaded.Name = "loaded"
        Me.loaded.Size = New System.Drawing.Size(62, 20)
        Me.loaded.TabIndex = 8
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(298, 130)
        Me.Label6.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(84, 20)
        Me.Label6.TabIndex = 7
        Me.Label6.Text = "Geladen:"
        Me.Label6.Visible = False
        '
        'shiftsApproved
        '
        Me.shiftsApproved.AutoSize = True
        Me.shiftsApproved.Checked = True
        Me.shiftsApproved.CheckState = System.Windows.Forms.CheckState.Checked
        Me.shiftsApproved.Location = New System.Drawing.Point(303, 40)
        Me.shiftsApproved.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.shiftsApproved.Name = "shiftsApproved"
        Me.shiftsApproved.Size = New System.Drawing.Size(384, 24)
        Me.shiftsApproved.TabIndex = 3
        Me.shiftsApproved.Text = "Nur Mitarbeiter mit allen KW Schichten genehmigt"
        Me.shiftsApproved.UseVisualStyleBackColor = True
        '
        'active_employees
        '
        Me.active_employees.AutoSize = True
        Me.active_employees.Checked = True
        Me.active_employees.CheckState = System.Windows.Forms.CheckState.Checked
        Me.active_employees.Location = New System.Drawing.Point(303, 97)
        Me.active_employees.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.active_employees.Name = "active_employees"
        Me.active_employees.Size = New System.Drawing.Size(279, 24)
        Me.active_employees.TabIndex = 4
        Me.active_employees.Text = "Nur in PLANDAY aktive Mitarbeiter"
        Me.active_employees.UseVisualStyleBackColor = True
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Location = New System.Drawing.Point(10, 128)
        Me.CheckBox1.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(61, 24)
        Me.CheckBox1.TabIndex = 2
        Me.CheckBox1.Text = "Alle"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(6, 72)
        Me.Label5.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(58, 20)
        Me.Label5.TabIndex = 3
        Me.Label5.Text = "KW:"
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(6, 14)
        Me.Label4.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(186, 18)
        Me.Label4.TabIndex = 2
        Me.Label4.Text = "Hotel:"
        '
        'select_kw
        '
        Me.select_kw.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.select_kw.FormattingEnabled = True
        Me.select_kw.IntegralHeight = False
        Me.select_kw.Location = New System.Drawing.Point(8, 94)
        Me.select_kw.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.select_kw.Name = "select_kw"
        Me.select_kw.Size = New System.Drawing.Size(286, 28)
        Me.select_kw.TabIndex = 1
        '
        'select_hotel
        '
        Me.select_hotel.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.select_hotel.FormattingEnabled = True
        Me.select_hotel.IntegralHeight = False
        Me.select_hotel.Location = New System.Drawing.Point(8, 37)
        Me.select_hotel.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.select_hotel.Name = "select_hotel"
        Me.select_hotel.Size = New System.Drawing.Size(286, 28)
        Me.select_hotel.TabIndex = 0
        '
        'LOAD_button
        '
        Me.LOAD_button.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LOAD_button.Cursor = System.Windows.Forms.Cursors.Hand
        Me.LOAD_button.Location = New System.Drawing.Point(900, 55)
        Me.LOAD_button.Name = "LOAD_button"
        Me.LOAD_button.Size = New System.Drawing.Size(136, 49)
        Me.LOAD_button.TabIndex = 10
        Me.LOAD_button.Text = "Daten laden"
        Me.ToolTip2.SetToolTip(Me.LOAD_button, "Überstunden für die KW laden")
        Me.LOAD_button.UseVisualStyleBackColor = True
        '
        'CLOSE_button
        '
        Me.CLOSE_button.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CLOSE_button.Cursor = System.Windows.Forms.Cursors.Hand
        Me.CLOSE_button.Location = New System.Drawing.Point(900, 106)
        Me.CLOSE_button.Name = "CLOSE_button"
        Me.CLOSE_button.Size = New System.Drawing.Size(136, 49)
        Me.CLOSE_button.TabIndex = 11
        Me.CLOSE_button.Text = "Beenden"
        Me.ToolTip2.SetToolTip(Me.CLOSE_button, "Bye bye :-)")
        Me.CLOSE_button.UseVisualStyleBackColor = True
        '
        'bgw
        '
        Me.bgw.WorkerReportsProgress = True
        Me.bgw.WorkerSupportsCancellation = True
        '
        'bgw_loading
        '
        Me.bgw_loading.WorkerReportsProgress = True
        '
        'bgw_write_overtime
        '
        Me.bgw_write_overtime.WorkerReportsProgress = True
        '
        'MODUS_button
        '
        Me.MODUS_button.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.MODUS_button.Appearance = System.Windows.Forms.Appearance.Button
        Me.MODUS_button.Cursor = System.Windows.Forms.Cursors.Hand
        Me.MODUS_button.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MODUS_button.Location = New System.Drawing.Point(900, 5)
        Me.MODUS_button.Name = "MODUS_button"
        Me.MODUS_button.Size = New System.Drawing.Size(136, 49)
        Me.MODUS_button.TabIndex = 9
        Me.MODUS_button.Text = "Senden AUS"
        Me.MODUS_button.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.MODUS_button.UseVisualStyleBackColor = True
        '
        'OvertimeForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1046, 568)
        Me.Controls.Add(Me.MODUS_button)
        Me.Controls.Add(Me.CLOSE_button)
        Me.Controls.Add(Me.LOAD_button)
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me.select_employees)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MinimumSize = New System.Drawing.Size(1068, 610)
        Me.Name = "OvertimeForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Überstunden"
        Me.Frame1.ResumeLayout(False)
        Me.Frame1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Private Sub OvertimeForm_Load(sender As Object, e As EventArgs) Handles Me.Load
        ImChangingStuff = True

        current_width = Me.select_employees.Width - 30
        Me.select_employees.GetType().GetProperty("DoubleBuffered", Reflection.BindingFlags.Instance Or Reflection.BindingFlags.NonPublic).SetValue(Me.select_employees, True)

        If CreateGraphics().DpiX <= 96 Then
            normalDPI = True
        Else
            normalDPI = False
        End If

        initialize()

        If ownDepartment = "" Then ownDepartment = get_ownDepartment()

        With Me.select_hotel
            For Each hotel_name In departments.Keys
                If ownDepartment = "NONE" Then Exit For
                If ownDepartment <> "ALL" Then
                    .Items.Add(ownDepartment)
                    Exit For
                End If
                .Items.Add(departments(hotel_name))
            Next
        End With

        Me.select_hotel.SelectedIndex = 0
        hotel_id = getDepartmentID(select_hotel.SelectedItem)

        With Me.select_employees
            .Columns.Clear()
        End With

        'Add the column headers
        Me.select_employees.Columns.Add(text:="Mitarbeiter", width:=294)
        Me.select_employees.Columns.Add(text:="Soll", width:=60, textAlign:=HorizontalAlignment.Right)
        Me.select_employees.Columns.Add(text:="IST", width:=60, textAlign:=HorizontalAlignment.Right)
        Me.select_employees.Columns.Add(text:="Überstunden", width:=116, textAlign:=HorizontalAlignment.Right)
        Me.select_employees.Columns.Add(text:="Zu übertragen", width:=116, textAlign:=HorizontalAlignment.Right)
        Me.select_employees.Columns.Add(text:="Übertragen", width:=116, textAlign:=HorizontalAlignment.Right)
        Me.select_employees.Columns.Add(text:="von PLANDAY", width:=116, textAlign:=HorizontalAlignment.Right)
        Me.select_employees.Columns.Add(text:="Kontostand", width:=116, textAlign:=HorizontalAlignment.Right)

        Me.PDF_button.Left = Me.DELETE_button.Left
        Me.EXCEL_button.Left = Me.WRITE_button.Left
        Me.LOAD_button.Select()

        For Each c As ColumnHeader In select_employees.Columns
            c.Tag = SortOrder.None
        Next

        For i = 0 To Me.select_employees.Columns.Count - 1
            column_widths(i) = Me.select_employees.Columns(i).Width
        Next

        If normalDPI Then
            Dim totalColumnWidth As Double = 0
            Me.Width = Me.Width + 80
            For i = 0 To select_employees.Columns.Count - 1
                totalColumnWidth += Convert.ToInt32(column_widths(i))
            Next
            If totalColumnWidth > select_employees.Width - 30 Then
                column_widths(0) = 294
                column_widths(1) = 60
                column_widths(2) = 60
                column_widths(3) = 116
                column_widths(4) = 116
                column_widths(5) = 116
                column_widths(6) = 116
                column_widths(7) = 116
            End If
            For i = 0 To select_employees.Columns.Count - 1
                Dim colPercentage As Double = (Convert.ToInt32(column_widths(i)) / totalColumnWidth)
                select_employees.Columns(i).Width = CInt(colPercentage * (select_employees.Width - 30))
            Next
        End If

        Dim start_monday As Date = PreviousMonday(Date.Today).AddDays(-52 * 7)

        With Me.select_kw
            Do While start_monday < PreviousMonday(Date.Today)
                If fGetKW(CDate(start_monday)) < 10 Then
                    .Items.Add("KW" & fGetKW(start_monday) & "           " & start_monday & "-" & start_monday.AddDays(6))
                Else
                    .Items.Add("KW" & fGetKW(start_monday) & "         " & start_monday & "-" & start_monday.AddDays(6))
                End If
                start_monday = start_monday.AddDays(7)
            Loop
        End With

        Me.select_kw.SelectedIndex = Me.select_kw.Items.Count - 1
        OvertimeForm.start_monday = CDate(Strings.Left(Strings.Right(select_kw.SelectedItem, 21), 10))
        current_year = Year(OvertimeForm.start_monday)

        Me.select_hotel.Select()
        ''finish loader''
        myWorkingForm.BeginInvoke(Sub() myWorkingForm.Close())

        ImChangingStuff = False
    End Sub
    Public Shared Function PreviousMonday(ByVal dateValue As DateTime) As DateTime
        Dim dayOffset As Integer
        Select Case dateValue.DayOfWeek
            Case DayOfWeek.Sunday : dayOffset = 6
            Case DayOfWeek.Monday : dayOffset = 0
            Case DayOfWeek.Tuesday : dayOffset = 1
            Case DayOfWeek.Wednesday : dayOffset = 2
            Case DayOfWeek.Thursday : dayOffset = 3
            Case DayOfWeek.Friday : dayOffset = 4
            Case DayOfWeek.Saturday : dayOffset = 5
        End Select

        Return dateValue.AddDays(-1 * dayOffset)
    End Function

    Public Shared Function fGetKW(dt As DateTime) As Integer
        Dim cal As Globalization.Calendar = Globalization.CultureInfo.InvariantCulture.Calendar
        Dim d As DayOfWeek = cal.GetDayOfWeek(dt)

        If (d >= DayOfWeek.Monday) AndAlso (d <= DayOfWeek.Wednesday) Then
            dt = dt.AddDays(3)
        End If

        Return cal.GetWeekOfYear(dt, Globalization.CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday)

    End Function

    Private Sub select_employees_ColumnClick(sender As Object, e As ColumnClickEventArgs) Handles select_employees.ColumnClick
        If ImChangingStuff Then Exit Sub
        ImChangingStuff = True

        Dim iSortOrder As SortOrder = CType(Me.select_employees.Columns(e.Column).Tag, SortOrder)
        Dim lvcs As New ListViewColumnSorter

        If iSortOrder = SortOrder.Ascending Then
            Me.select_employees.Columns(e.Column).Tag = SortOrder.Descending
            lvcs.SortingOrder = SortOrder.Descending
        Else
            Me.select_employees.Columns(e.Column).Tag = SortOrder.Ascending
            lvcs.SortingOrder = SortOrder.Ascending
        End If
        lvcs.ColumnIndex = e.Column
        Me.select_employees.ListViewItemSorter = lvcs

        ImChangingStuff = False
    End Sub

    Private Sub CheckBox1_Click(sender As Object, e As EventArgs) Handles CheckBox1.Click
        Dim column_count As Integer
        Dim Item As ListViewItem
        If ImChangingStuff Then Exit Sub

        column_count = Me.select_employees.Columns.Count
        If CheckBox1.Checked = True Then

            For i = 0 To Me.select_employees.Items.Count - 1
                Item = Me.select_employees.Items.Item(i)
                Item.Checked = True
                Item.Font = New Font(Item.Font, FontStyle.Bold)
                Item.SubItems(3).Font = New Font(Item.SubItems(3).Font, FontStyle.Bold)
                On Error Resume Next
                If Val(Item.SubItems(4).Text) > 0 Then
                    Item.SubItems(4).ForeColor = Color.FromArgb(0, 140, 60)
                    Item.SubItems(4).Font = New Font(Item.SubItems(4).Font, FontStyle.Bold)
                End If
                If Val(Item.SubItems(4).Text) < 0 Then
                    Item.SubItems(4).ForeColor = Color.Red
                    Item.SubItems(4).Font = New Font(Item.SubItems(4).Font, FontStyle.Bold)
                End If
            Next
            Me.CheckBox1.Font = New Font(Me.CheckBox1.Font, FontStyle.Bold)
        Else

            For i = 0 To Me.select_employees.Items.Count - 1
                Item = Me.select_employees.Items.Item(i)
                Item.Checked = False
                Item.Font = New Font(Item.Font, FontStyle.Regular)
                Item.SubItems(3).Font = New Font(Item.SubItems(3).Font, FontStyle.Regular)
                Item.SubItems(4).ForeColor = Color.Black
                Item.SubItems(4).Font = New Font(Item.SubItems(4).Font, FontStyle.Regular)
            Next
            Me.CheckBox1.Font = New Font(CheckBox1.Font, FontStyle.Regular)
        End If
    End Sub


    Private Sub select_employees_ItemChecked(sender As Object, e As ItemCheckedEventArgs) Handles select_employees.ItemChecked
        Dim column_count As Integer
        Dim Item As ListViewItem
        If ImChangingStuff Then Exit Sub
        ImChangingStuff = True

        Item = e.Item
        column_count = Me.select_employees.Columns.Count
        If Item.Checked Then
            Item.Font = New Font(Item.Font, FontStyle.Bold)
            Item.SubItems(3).Font = New Font(Item.SubItems(3).Font, FontStyle.Bold)

            If Val(Item.SubItems(4).Text) > 0 Then
                Item.SubItems(4).ForeColor = Color.FromArgb(0, 140, 60)
                Item.SubItems(4).Font = New Font(Item.SubItems(4).Font, FontStyle.Bold)
            End If
            If Val(Item.SubItems(4).Text) < 0 Then
                Item.SubItems(4).ForeColor = Color.Red
                Item.SubItems(4).Font = New Font(Item.SubItems(4).Font, FontStyle.Bold)
            End If
            If select_employees.CheckedItems.Count = select_employees.Items.Count Then
                CheckBox1.Checked = True
                CheckBox1.Font = New Font(CheckBox1.Font, FontStyle.Bold)
            End If
        Else
            Item.Font = New Font(Item.Font, FontStyle.Regular)
            Item.SubItems(3).Font = New Font(Item.SubItems(3).Font, FontStyle.Regular)
            Item.SubItems(4).ForeColor = Color.Black
            Item.SubItems(4).Font = New Font(Item.SubItems(4).Font, FontStyle.Regular)
            CheckBox1.Checked = False
            CheckBox1.Font = New Font(CheckBox1.Font, FontStyle.Regular)
        End If

        ImChangingStuff = False
    End Sub

    Private Sub select_employees_SizeChanged(sender As Object, e As EventArgs) Handles select_employees.SizeChanged
        If WidthChanging Or ImChangingStuff Then Exit Sub

        ImChangingStuff = True
        select_employees.BeginUpdate()
        Dim totalColumnWidth As Double = 0
        For i = 0 To select_employees.Columns.Count - 1
            totalColumnWidth += Convert.ToInt32(column_widths(i))
        Next
        If totalColumnWidth > select_employees.Width - 30 Then
            column_widths(0) = 294
            column_widths(1) = 60
            column_widths(2) = 60
            column_widths(3) = 116
            column_widths(4) = 116
            column_widths(5) = 116
            column_widths(6) = 116
            column_widths(7) = 116
        End If
        For i = 0 To select_employees.Columns.Count - 1
            Dim colPercentage As Double = (Convert.ToInt32(column_widths(i)) / totalColumnWidth)
            select_employees.Columns(i).Width = CInt(colPercentage * (select_employees.Width - 30))
        Next
        select_employees.EndUpdate()

        ImChangingStuff = False
    End Sub

    Private Sub select_employees_ColumnWidthChanging(sender As Object, e As ColumnWidthChangingEventArgs) Handles select_employees.ColumnWidthChanging
        WidthChanging = True
    End Sub

    Friend WithEvents select_employees As ListView
    Friend WithEvents Frame1 As GroupBox
    Friend WithEvents LOAD_button As Button
    Friend WithEvents CLOSE_button As Button
    Friend WithEvents select_kw As ComboBox
    Friend WithEvents select_hotel As ComboBox
    Friend WithEvents Label5 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents CheckBox1 As CheckBox
    Friend WithEvents Label6 As Label
    Friend WithEvents shiftsApproved As CheckBox
    Friend WithEvents active_employees As CheckBox
    Friend WithEvents loaded As Label
    Friend WithEvents FT_label As Label
    Friend WithEvents FT_2 As Label
    Friend WithEvents FT_1 As Label
    Friend WithEvents DELETE_button As Button
    Friend WithEvents WRITE_button As Button
    Friend WithEvents EXCEL_button As Button
    Friend WithEvents bgw As System.ComponentModel.BackgroundWorker
    Friend WithEvents ToolTip1 As ToolTip
    Friend WithEvents bgw_loading As System.ComponentModel.BackgroundWorker
    Private WidthChanging As Boolean
    Friend WithEvents bgw_write_overtime As System.ComponentModel.BackgroundWorker
    Friend WithEvents ToolTip2 As ToolTip
    Friend WithEvents PDF_button As Button
    Friend WithEvents MODUS_button As CheckBox
End Class
